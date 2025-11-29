function Get-FileHandle
{
    <#
    .SYNOPSIS
    Returns processes (name + PID) that have a lock on a given file.

    .DESCRIPTION
    Uses Sysinternals handle.exe to query for handles matching the specified file path
    and parses the output to return unique objects containing the process name and PID
    that currently hold a lock on the file. If no locks are found, returns an empty array.

    .PARAMETER FilePath
    The full path to the file to check for open handles.

    .OUTPUTS
    PSCustomObject[] with properties: ProcessName (string), PID (int)

    .EXAMPLE
    Get-FileHandle -FilePath "C:\\Temp\\my.log"
    #> 
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchString,
        [string]$HandlePath = "$pwd\handle.exe"
    )
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Searching for file handles matching: $SearchString"
    write-log -logFile $LogFile -module $functionName -message "Searching for file handles matching: $SearchString"
    Write-Verbose "[$functionName] Using handle.exe path: $HandlePath"
    write-log -logFile $LogFile -module $functionName -message "Using handle.exe path: $HandlePath"
    if ($null -eq $HandlePath -or $HandlePath -eq '')
    {
        Write-Verbose "[$functionName] Using default handle.exe path: $handleExe"
        write-log -logFile $LogFile -module $functionName -message "Using default handle.exe path: $handleExe"
        $handleExe = "$pwd\handle.exe"
    }
    else
    {
        Write-Verbose "[$functionName] Using custom handle.exe path: $HandlePath"
        write-log -logFile $LogFile -module $functionName -message "Using custom handle.exe path: $HandlePath"
        $handleExe = $HandlePath
    }

    #validate the path is valid.
    if (-not (Test-Path $handleExe))
    {
        Write-Error "handle.exe not found at $handleExe"
        Write-Verbose "[$functionName] handle.exe not found at $handleExe"
        write-log -logFile $LogFile -module $functionName -message "handle.exe not found at $handleExe"
        return @()
    }
    Write-Verbose "[$functionName] handle.exe found at: $handleExe"
    Write-log -logFile $LogFile -module $functionName -message "handle.exe found at: $handleExe"
    
    # Invoke handle.exe; it will filter by the provided path
    $arguments = @('-accepteula', '-nobanner', $SearchString)
    Write-Verbose "[$functionName] Invoking handle.exe with arguments: $arguments"
    write-log -logFile $LogFile -module $functionName -message "Invoking handle.exe with arguments: $arguments"
    $global:output = & $handleExe @arguments 2>&1
    Write-Verbose "[$functionName] handle.exe output: $($output -join "`n")"
    write-log -logFile $LogFile -module $functionName -message "handle.exe output: $($output -join "`n")"
    if (-not $output -or ($output -join "`n") -match 'No matching handles')
    {
        Write-Verbose "[$functionName] No matching handles found."
        write-log -logFile $LogFile -module $functionName -message "No matching handles found."
        return @()
    }
    $items = @()
    foreach ($line in $output)
    {
        Write-Verbose "[$functionName] Processing line: $line"
        write-log -logFile $LogFile -module $functionName -message "Processing line: $line"
        $m = [regex]::Match($line, '^\s*([\w\-.]+\.exe)\s+pid:\s*(\d+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success)
        {
            $procName = $m.Groups[1].Value
            $ProcId = [int]$m.Groups[2].Value
            Write-Verbose "[$functionName] Found matching process: $procName (PID: $ProcId)"
            write-log -logFile $LogFile -module $functionName -message "Found matching process: $procName (PID: $ProcId)"
            $items += [pscustomobject]@{
                ProcessName = $procName
                ProcessId   = $ProcId
            }
        }
    }
    if ($items.count -gt 1)
    {
        $items | Sort-Object -Property ProcessId, ProcessName -Unique
        Write-Verbose "[$functionName] Returning $($items.count) unique items: $($items | Format-Table -AutoSize | Out-String)"
        write-log -logFile $LogFile -module $functionName -message "Returning $($items.count) unique items: $($items | Format-Table -AutoSize | Out-String)"
    }
    return $items
}
