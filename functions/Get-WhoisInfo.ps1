function Get-WhoisInfo()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DomainNameOrIPAddress,
        [Parameter(Mandatory = $false)]
        [string]$WhoisServer,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 10
    )

    $functionName = $MyInvocation.MyCommand.Name
    # Resolve log file from common scopes with a safe fallback in case bootstrap didn't set it
    try
    {
        if (-not (Get-Variable -Name LogFile -Scope 0 -ErrorAction SilentlyContinue))
        {
            if (Get-Variable -Name LogFile -Scope Script -ErrorAction SilentlyContinue) { $LogFile = $Script:LogFile }
            elseif (Get-Variable -Name LogFile -Scope Global -ErrorAction SilentlyContinue) { $LogFile = $Global:LogFile }
            else { $LogFile = Join-Path $env:TEMP 'Autopilot.log' }
        }
    }
    catch { $LogFile = Join-Path $env:TEMP 'Autopilot.log' }
    # Helper: simple check if input is an IP address
    function Test-IsIPAddress()
    {
        [CmdletBinding()]
        param(
            [string]$InputString
        )
        $ip = $null
        return [System.Net.IPAddress]::TryParse($InputString, [ref]$ip)
    }

    # Helper: perform a WHOIS query against a server
    function Invoke-WhoisQuery()
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$Server,
            [Parameter(Mandatory = $true)][string]$Query,
            [Parameter(Mandatory = $false)][int]$Timeout = 10
        )
        Write-Log -LogFile $LogFile -Module $functionName -Message ("Connecting to $Server:43 to query '$Query'") -LogLevel "Verbose"
        $client = $null
        $stream = $null
        $reader = $null
        $writer = $null
        try
        {
            $client = New-Object System.Net.Sockets.TcpClient
            $ar = $client.BeginConnect($Server, 43, $null, $null)
            if (-not $ar.AsyncWaitHandle.WaitOne($Timeout * 1000, $false))
            {
                throw "WHOIS connection to $Server timed out after $Timeout seconds"
            }
            $client.EndConnect($ar) | Out-Null
            $stream = $client.GetStream()
            $stream.ReadTimeout = $Timeout * 1000
            $stream.WriteTimeout = $Timeout * 1000
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::ASCII)
            $writer.NewLine = "`r`n"
            $writer.AutoFlush = $true
            # Some WHOIS servers are picky; ensure just the query is sent
            [void]$writer.WriteLine($Query)
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
            $responseBuilder = New-Object System.Text.StringBuilder
            while (-not $reader.EndOfStream)
            {
                $line = $reader.ReadLine()
                [void]$responseBuilder.AppendLine($line)
            }
            $resp = $responseBuilder.ToString()
            Write-Log -LogFile $LogFile -Module $functionName -Message ("Received WHOIS response from $Server (length: $($resp.Length))")
            return $resp
        }
        catch
        {
            Write-Log -LogFile $LogFile -Module $functionName -Message ("WHOIS query to $Server failed: $($_.Exception.Message)") -LogLevel "Error"
            return $null
        }
        finally
        {
            if ($reader) { $reader.Dispose() }
            if ($writer) { $writer.Dispose() }
            if ($stream) { $stream.Dispose() }
            if ($client) { $client.Close() }
        }
    }

    # Helper: determine default WHOIS server for domains via IANA, with fallback to whois-servers.net convention
    function Get-WhoisServerForDomain()
    {
        [CmdletBinding()]
        param([string]$Domain)
        try
        {
            $tld = ($Domain -split '\\.')[-1]
            if (-not $tld) { return $null }
            $ianaResp = Invoke-WhoisQuery -Server 'whois.iana.org' -Query $tld -Timeout $TimeoutSeconds
            if ($null -ne $ianaResp)
            {
                foreach ($l in ($ianaResp -split "`n"))
                {
                    $line = $l.Trim()
                    if ($line -match '^(?i)whois:\s*(?<srv>[^\s]+)')
                    {
                        return $Matches['srv'].Trim()
                    }
                }
            }
            # Fallback pattern used broadly by clients
            return ("{0}.whois-servers.net" -f $tld)
        }
        catch
        {
            Write-Log -LogFile $LogFile -Module $functionName -Message ("Failed to resolve WHOIS server via IANA: $($_.Exception.Message)") -LogLevel "Warning"
            return $null
        }
    }

    # Helper: resolve referral if present
    function Resolve-WhoisReferral()
    {
        [CmdletBinding()]
        param([string]$Response, [string]$OriginalQuery, [string]$CurrentServer)
        if (-not $Response) { return @($Response, $CurrentServer) }
        $referralServer = $null
        foreach ($l in ($Response -split "`n"))
        {
            $line = $l.Trim()
            if ($line -match '^(?i)ReferralServer:\s*whois://(?<srv>[^\s/:]+)')
            {
                $referralServer = $Matches['srv']
                break
            }
            if (-not $referralServer -and $line -match '^(?i)whois:\s*(?<srv>[^\s]+)$')
            {
                # Some registries use just 'whois: host'
                $referralServer = $Matches['srv']
                break
            }
        }
        if ($referralServer)
        {
            Write-Log -LogFile $LogFile -Module $functionName -Message ("Following WHOIS referral to $referralServer for '$OriginalQuery'") -LogLevel "Information"
            $refResp = Invoke-WhoisQuery -Server $referralServer -Query $OriginalQuery -Timeout $TimeoutSeconds
            if ($refResp) { return @($refResp, $referralServer) }
        }
        return @($Response, $CurrentServer)
    }

    # Decide server and perform query
    $serverToUse = $WhoisServer
    $isIP = Test-IsIPAddress -InputString $DomainNameOrIPAddress
    if (-not $serverToUse)
    {
        if ($isIP)
        {
            $serverToUse = 'whois.arin.net'
        }
        else
        {
            $serverToUse = Get-WhoisServerForDomain -Domain $DomainNameOrIPAddress
            if (-not $serverToUse) { $serverToUse = 'whois.verisign-grs.com' }
        }
    }

    Write-Log -LogFile $LogFile -Module $functionName -Message ("WHOIS lookup starting. Query: $DomainNameOrIPAddress, Server: $serverToUse, Timeout: $TimeoutSeconds s") -LogLevel "Information"

    $raw = Invoke-WhoisQuery -Server $serverToUse -Query $DomainNameOrIPAddress -Timeout $TimeoutSeconds
    $usedServer = $serverToUse
    $followed = Resolve-WhoisReferral -Response $raw -OriginalQuery $DomainNameOrIPAddress -CurrentServer $usedServer
    if ($followed -and $followed.Count -eq 2)
    {
        $raw = $followed[0]
        $usedServer = $followed[1]
    }

    if (-not $raw)
    {
        Write-Log -LogFile $LogFile -Module $functionName -Message ("WHOIS lookup failed or returned no data for '$DomainNameOrIPAddress'") -LogLevel "Error"
        return @{ 
            Query           = $DomainNameOrIPAddress
            WhoisServerUsed = $usedServer
            Error           = "Lookup failed or returned no data"
            RawText         = ''
        }
    }

    # Build clean result with deduplication and consolidated notices
    $result = @{}
    $result['Query'] = $DomainNameOrIPAddress
    $result['WhoisServerUsed'] = $usedServer
    $result['RawText'] = $raw

    # Temporary accumulators
    $kv = @{}
    $noticeBuilder = New-Object System.Text.StringBuilder
    $termsBuilder = New-Object System.Text.StringBuilder
    $registryInfoBuilder = New-Object System.Text.StringBuilder
    $statusInfoUrl = $null
    $lastUpdate = $null
    $inNotice = $false
    $inTerms = $false
    $inRegistryInfo = $false

    $lines = $raw -split "`n"
    foreach ($l in $lines)
    {
        $lineRaw = $l
        $line = $l.Trim()

        # Capture metadata
        if ($line -match '^(?i)>>>' ) { continue }
        if ($line -match '^(?i)>>>\s*Last update of whois database:\s*(.+)$') { $lastUpdate = $Matches[1].Trim(); continue }
        if ($line -match '^(?i)For more information on Whois status codes, please visit\s+(?<url>\S+)$') { $statusInfoUrl = $Matches['url']; continue }

        # NOTICE block
        if ($line -match '^(?i)NOTICE:') { $inNotice = $true }
        if ($inNotice)
        {
            if (-not [string]::IsNullOrWhiteSpace($line)) { [void]$noticeBuilder.AppendLine($lineRaw.TrimEnd()) } else { $inNotice = $false }
            continue
        }

        # TERMS OF USE block
        if ($line -match '^(?i)TERMS OF USE:') { $inTerms = $true }
        if ($inTerms)
        {
            if (-not [string]::IsNullOrWhiteSpace($line)) { [void]$termsBuilder.AppendLine($lineRaw.TrimEnd()) } else { $inTerms = $false }
            continue
        }

        # Registry info block
        if ($line -match '^(?i)The Registry database contains ONLY') { $inRegistryInfo = $true }
        if ($inRegistryInfo)
        {
            if (-not [string]::IsNullOrWhiteSpace($line)) { [void]$registryInfoBuilder.AppendLine($lineRaw.TrimEnd()) } else { $inRegistryInfo = $false }
            continue
        }

        # Skip comments/empty
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.StartsWith('%') -or $line.StartsWith('#')) { continue }
        if ($line -match '^(?i)for more information') { continue }

        # Parse key: value
        $idx = $line.IndexOf(':')
        if ($idx -lt 1) { continue }
        $key = $line.Substring(0, $idx).Trim()
        $value = $line.Substring($idx + 1).Trim()
        if (-not $key) { continue }

        if (-not $kv.ContainsKey($key)) { $kv[$key] = New-Object System.Collections.Generic.List[string] }
        if (-not $kv[$key].Contains($value)) { $kv[$key].Add($value) }
    }

    # Materialize deduped
    foreach ($k in $kv.Keys)
    {
        $vals = $kv[$k]
        if ($vals.Count -eq 1) { $result[$k] = $vals[0] }
        else { $result[$k] = $vals.ToArray() }
    }

    # Consolidated notices object
    $notices = @{}
    if ($noticeBuilder.Length -gt 0) { $notices['Notice'] = $noticeBuilder.ToString().Trim() }
    if ($termsBuilder.Length -gt 0) { $notices['TermsOfUse'] = $termsBuilder.ToString().Trim() }
    if ($registryInfoBuilder.Length -gt 0) { $notices['RegistryDatabaseContains'] = $registryInfoBuilder.ToString().Trim() }
    if ($statusInfoUrl) { $notices['StatusCodesInfoUrl'] = $statusInfoUrl }
    if ($lastUpdate) { $notices['LastUpdate'] = $lastUpdate }
    if ($notices.Keys.Count -gt 0) { $result['Notices'] = $notices }

    Write-Log -LogFile $LogFile -Module $functionName -Message ("WHOIS lookup complete for '$DomainNameOrIPAddress' using $usedServer. Parsed keys: $($result.Keys.Count)") -LogLevel "Information"
    return $result
}