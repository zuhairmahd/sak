function Get-LocalComputerInfo()
{
    <#
    .SYNOPSIS
        Retrieves client computer information from within a Citrix or RDP session.

    .DESCRIPTION
        This function gathers information about the CLIENT computer when running inside a 
        Citrix ICA or RDP virtual session. It detects the session type and uses appropriate 
        methods to retrieve client-side information including OS, hostname, IP address, and 
        other relevant details. Supports both Windows and IGEL (Linux) clients.

    .PARAMETER LogFile
        Optional. Full path to the log file. If not specified, logging is skipped.

    .PARAMETER Module
        Optional. Module name for logging context. Defaults to the function name.

    .PARAMETER WriteToConsole
        Optional. If specified, also writes log messages to the console.

    .EXAMPLE
        PS C:\> Get-LocalComputerInfo
        Retrieves and returns client computer information as a hashtable.

    .EXAMPLE
        PS C:\> Get-LocalComputerInfo -LogFile "C:\Logs\clientinfo.log" -WriteToConsole
        Retrieves client info with logging enabled to file and console.

    .OUTPUTS
        System.Collections.Hashtable
        Returns a hashtable with the following keys:
        - SessionType: "Citrix", "RDP", or "Local"
        - ClientName: Client hostname
        - ClientOS: Operating system name
        - ClientIPAddress: Client IP address
        - ClientVersion: Client software version (if available)
        - IsIGEL: Boolean indicating if client is IGEL OS
        - ServerName: Virtual session server name
        - ErrorOccurred: Boolean indicating if any errors occurred
        - ErrorMessages: Array of error messages (if any)

    .NOTES
        Author: MahmoudZ
        Date: 2024-12-03
        Version: 2.0
        
        Environment Variables Used:
        - Citrix: CLIENTNAME, CLIENTADDRESS, ICA_CLIENT_*, HDX_*
        - RDP: SESSIONNAME, CLIENTNAME (from Win32_TSEnvironment)
        
        Requirements:
        - PowerShell 5.1 or higher
        - Runs within a Citrix ICA or RDP session
        - Write-Log function must be available if logging is enabled
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$LogFile,
        [Parameter(Mandatory = $false)]
        [switch]$WriteToConsole
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    $enableLogging = -not [string]::IsNullOrEmpty($LogFile)
    $result = @{
        SessionType     = "Unknown"
        ClientName      = $null
        ClientOS        = $null
        ClientIPAddress = $null
        ClientVersion   = $null
        IsIGEL          = $false
        ServerName      = $env:COMPUTERNAME
        ErrorOccurred   = $false
        ErrorMessages   = @()
    }
    
    # Helper function to log messages
    function Write-FunctionLog
    {
        param(
            [string]$Message,
            [string]$LogLevel = "Information"
        )
        
        if ($enableLogging)
        {
            try
            {
                Write-Log -Message $Message -LogFile $LogFile -Module $Module -LogLevel $LogLevel -WriteToConsole:$WriteToConsole
            }
            catch
            {
                if ($WriteToConsole)
                {
                    Write-Warning "[$functionName] Failed to write log: $_"
                }
            }
        }
        elseif ($WriteToConsole)
        {
            Write-Verbose "[$functionName] $Message"
        }
    }
    
    Write-FunctionLog -Message "Starting client computer information retrieval" -LogLevel "Information"
    Write-FunctionLog -Message "Running on server: $($env:COMPUTERNAME)" -LogLevel "Verbose"
    
    try
    {
        # Detect session type
        Write-FunctionLog -Message "Detecting session type..." -LogLevel "Verbose"
        
        # Check for Citrix session
        $isCitrix = $false
        $isRDP = $false
        
        # Citrix detection via environment variables
        $citrixEnvVars = @('CLIENTNAME', 'ICA_CLIENT_NAME', 'HDX_RTAV_SESSION')
        foreach ($envVar in $citrixEnvVars)
        {
            if ([Environment]::GetEnvironmentVariable($envVar))
            {
                $isCitrix = $true
                Write-FunctionLog -Message "Citrix session detected via environment variable: $envVar" -LogLevel "Verbose"
                break
            }
        }
        
        # Additional Citrix detection via registry
        if (-not $isCitrix)
        {
            try
            {
                $citrixRegPath = "HKCU:\Software\Citrix\ICA Client"
                if (Test-Path $citrixRegPath)
                {
                    $isCitrix = $true
                    Write-FunctionLog -Message "Citrix session detected via registry" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-FunctionLog -Message "Error checking Citrix registry: $($_.Exception.Message)" -LogLevel "Debug"
            }
        }
        
        # RDP detection via SESSIONNAME environment variable
        $sessionName = $env:SESSIONNAME
        if (-not $isCitrix -and $sessionName -and $sessionName -ne "Console" -and $sessionName -match "RDP-")
        {
            $isRDP = $true
            Write-FunctionLog -Message "RDP session detected via SESSIONNAME: $sessionName" -LogLevel "Verbose"
        }
        
        # Set session type
        if ($isCitrix)
        {
            $result.SessionType = "Citrix"
            Write-FunctionLog -Message "Session type determined: Citrix" -LogLevel "Information"
        }
        elseif ($isRDP)
        {
            $result.SessionType = "RDP"
            Write-FunctionLog -Message "Session type determined: RDP" -LogLevel "Information"
        }
        else
        {
            $result.SessionType = "Local"
            Write-FunctionLog -Message "Session type determined: Local (no remote session detected)" -LogLevel "Information"
        }
        
        # Retrieve client information based on session type
        if ($isCitrix)
        {
            Write-FunctionLog -Message "Retrieving Citrix client information..." -LogLevel "Verbose"
            
            # Get client name
            $result.ClientName = [Environment]::GetEnvironmentVariable('CLIENTNAME')
            if (-not $result.ClientName)
            {
                $result.ClientName = [Environment]::GetEnvironmentVariable('ICA_CLIENT_NAME')
            }
            Write-FunctionLog -Message "Citrix client name: $($result.ClientName)" -LogLevel "Verbose"
            
            # Get client IP address
            $result.ClientIPAddress = [Environment]::GetEnvironmentVariable('CLIENTADDRESS')
            if (-not $result.ClientIPAddress)
            {
                $result.ClientIPAddress = [Environment]::GetEnvironmentVariable('ICA_CLIENT_ADDRESS')
            }
            Write-FunctionLog -Message "Citrix client IP: $($result.ClientIPAddress)" -LogLevel "Verbose"
            
            # Detect IGEL OS via various indicators
            $igelIndicators = @(
                [Environment]::GetEnvironmentVariable('IGEL_LICENSE'),
                [Environment]::GetEnvironmentVariable('IGEL_SYSTEM'),
                ([Environment]::GetEnvironmentVariable('CLIENTNAME') -match '^IGEL-'),
                ([Environment]::GetEnvironmentVariable('ICA_CLIENT_PRODUCTID') -match 'IGEL')
            )
            
            $result.IsIGEL = $igelIndicators -contains $true
            Write-FunctionLog -Message "IGEL OS detected: $($result.IsIGEL)" -LogLevel "Verbose"
            
            # Attempt to get client OS from Citrix environment variables
            $clientPlatform = [Environment]::GetEnvironmentVariable('ICA_CLIENT_PLATFORM')
            if ($clientPlatform)
            {
                $result.ClientOS = $clientPlatform
                Write-FunctionLog -Message "Client platform from ICA_CLIENT_PLATFORM: $clientPlatform" -LogLevel "Verbose"
            }
            elseif ($result.IsIGEL)
            {
                $result.ClientOS = "IGEL OS (Linux)"
                Write-FunctionLog -Message "Client OS determined as IGEL OS" -LogLevel "Verbose"
            }
            else
            {
                $result.ClientOS = "Unknown (Citrix Client)"
            }
            
            # Get client version
            $result.ClientVersion = [Environment]::GetEnvironmentVariable('ICA_CLIENT_VERSION')
            Write-FunctionLog -Message "Citrix client version: $($result.ClientVersion)" -LogLevel "Verbose"
            
            # Try to get more detailed information via WMI (if available)
            try
            {
                $citrixSession = Get-WmiObject -Class Citrix_VirtualChannel -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($citrixSession)
                {
                    Write-FunctionLog -Message "Additional Citrix session info retrieved via WMI" -LogLevel "Debug"
                }
            }
            catch
            {
                Write-FunctionLog -Message "Could not retrieve Citrix WMI info: $($_.Exception.Message)" -LogLevel "Debug"
            }
        }
        elseif ($isRDP)
        {
            Write-FunctionLog -Message "Retrieving RDP client information..." -LogLevel "Verbose"
            
            # Try to get client information from Win32_TSEnvironment
            try
            {
                $tsEnv = Get-WmiObject -Namespace "root\cimv2\TerminalServices" -Class Win32_TSEnvironment -ErrorAction Stop
                if ($tsEnv)
                {
                    $result.ClientName = $tsEnv.ClientName
                    $result.ClientIPAddress = $tsEnv.ClientIPAddress
                    Write-FunctionLog -Message "RDP client name from WMI: $($result.ClientName)" -LogLevel "Verbose"
                    Write-FunctionLog -Message "RDP client IP from WMI: $($result.ClientIPAddress)" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-FunctionLog -Message "Failed to query Win32_TSEnvironment: $($_.Exception.Message)" -LogLevel "Warning"
                $result.ErrorMessages += "WMI query failed: $($_.Exception.Message)"
            }
            
            # Fallback: Try environment variable
            if (-not $result.ClientName)
            {
                $result.ClientName = [Environment]::GetEnvironmentVariable('CLIENTNAME')
                Write-FunctionLog -Message "RDP client name from env var: $($result.ClientName)" -LogLevel "Verbose"
            }
            
            # Try to get client information from Win32_LogonSession
            try
            {
                $sessionId = (Get-Process -Id $PID).SessionId
                Write-FunctionLog -Message "Current session ID: $sessionId" -LogLevel "Debug"
                
                # Query qwinsta for client info (fallback method)
                $qwinstaOutput = qwinsta 2>$null
                if ($qwinstaOutput)
                {
                    $currentSessionLine = $qwinstaOutput | Where-Object { $_ -match "rdp-tcp#$sessionId" }
                    if ($currentSessionLine)
                    {
                        Write-FunctionLog -Message "Session info from qwinsta: $currentSessionLine" -LogLevel "Debug"
                    }
                }
            }
            catch
            {
                Write-FunctionLog -Message "Failed to query session info: $($_.Exception.Message)" -LogLevel "Debug"
            }
            
            # Determine client OS (RDP doesn't provide easy access to this)
            $result.ClientOS = "Unknown (RDP Client)"
            Write-FunctionLog -Message "RDP client OS: Cannot be determined from RDP session" -LogLevel "Verbose"
        }
        else
        {
            # Local session - get local computer info
            Write-FunctionLog -Message "Local session detected - retrieving local computer information" -LogLevel "Verbose"
            
            try
            {
                $computerInfo = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
                $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
                
                $result.ClientName = $computerInfo.Name
                $result.ClientOS = "$($osInfo.Caption) $($osInfo.Version)"
                Write-FunctionLog -Message "Local computer name: $($result.ClientName)" -LogLevel "Verbose"
                Write-FunctionLog -Message "Local OS: $($result.ClientOS)" -LogLevel "Verbose"
                
                # Get IP address
                $networkAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction Stop | 
                    Where-Object { $_.IPEnabled -eq $true }
                
                if ($networkAdapters)
                {
                    $result.ClientIPAddress = ($networkAdapters | Select-Object -First 1).IPAddress[0]
                    Write-FunctionLog -Message "Local IP address: $($result.ClientIPAddress)" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-FunctionLog -Message "Failed to query local computer info: $($_.Exception.Message)" -LogLevel "Error"
                $result.ErrorOccurred = $true
                $result.ErrorMessages += "Failed to query local WMI: $($_.Exception.Message)"
            }
        }
        
        # Validate that we got at least some information
        if (-not $result.ClientName -and -not $result.ClientIPAddress)
        {
            Write-FunctionLog -Message "Warning: Could not retrieve client name or IP address" -LogLevel "Warning"
            $result.ErrorMessages += "No client identification information could be retrieved"
            
            # Last resort: try hostname command
            try
            {
                $hostnameOutput = hostname 2>$null
                if ($hostnameOutput)
                {
                    $result.ClientName = $hostnameOutput.Trim()
                    Write-FunctionLog -Message "Fallback: Got hostname from hostname command: $($result.ClientName)" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-FunctionLog -Message "Hostname command also failed: $($_.Exception.Message)" -LogLevel "Debug"
            }
        }
        
        Write-FunctionLog -Message "Client information retrieval completed successfully" -LogLevel "Information"
        Write-FunctionLog -Message "Summary - Type: $($result.SessionType), Name: $($result.ClientName), IP: $($result.ClientIPAddress), OS: $($result.ClientOS)" -LogLevel "Information"
    }
    catch
    {
        $errorMsg = "Unexpected error in Get-LocalComputerInfo: $($_.Exception.Message)"
        Write-FunctionLog -Message $errorMsg -LogLevel "Error"
        $result.ErrorOccurred = $true
        $result.ErrorMessages += $errorMsg
        
        if ($WriteToConsole)
        {
            Write-Error "[$functionName] $errorMsg"
        }
    }
    
    # Return the result hashtable
    return $result
}                                   