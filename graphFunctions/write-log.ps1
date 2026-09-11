function Write-Log()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [string]$Message,
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $true, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $true, ParameterSetName = 'FinishLogging')]
        [ValidateScript({
                $parentDir = Split-Path $_ -Parent
                if (-not (Test-Path $parentDir))
                {
                    try
                    {
                        New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                    }
                    catch
                    {
                        throw "Failed to create log directory: $_. Exception: $($_.Exception.Message)"
                    }
                }
                return $true
            })]
        [string]$LogFile,
        [Parameter(Mandatory = $true, ParameterSetName = 'Normal')]
        [string]$Module,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [switch]$WriteToConsole,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [ValidateSet("Verbose", "Debug", "Information", "Warning", "Error")]
        [string]$LogLevel = "Information",
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [switch]$CMTraceFormat,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [int]$MaxLogSizeMB = 10,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [switch]$PassThru,
        [Parameter(Mandatory = $true, ParameterSetName = 'StartLogging')]
        [switch]$StartLogging,
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [switch]$OverwriteLog,
        [Parameter(Mandatory = $true, ParameterSetName = 'FinishLogging')]
        [switch]$FinishLogging,
        [Parameter(Mandatory = $false, ParameterSetName = 'Normal')]
        [Parameter(Mandatory = $false, ParameterSetName = 'StartLogging')]
        [Parameter(Mandatory = $false, ParameterSetName = 'FinishLogging')]
        [ValidateSet('Error', 'Warning', 'Information', 'Verbose', 'Debug')]
        [string]$MinimumLogLevel
    )
    
    try
    {
        # Use global minimum log level if not provided
        if (-not $MinimumLogLevel -and $Global:MinimumLogLevel)
        {
            $MinimumLogLevel = $Global:MinimumLogLevel
        }
        elseif (-not $MinimumLogLevel)
        {
            $MinimumLogLevel = 'Information'
        }
        
        # Define log level hierarchy (higher numbers = more detailed logging)
        $logLevelHierarchy = @{
            'Error'       = 1
            'Warning'     = 2
            'Information' = 3
            'Verbose'     = 4
            'Debug'       = 5
        }
        
        # Handle StartLogging and FinishLogging switches
        if ($StartLogging -or $FinishLogging)
        {
            # Set default values when using StartLogging or FinishLogging
            $Module = $MyInvocation.MyCommand.Name
            $LogLevel = "Information"
            
            # Create separator line with appropriate message
            if ($StartLogging)
            {
                $separatorLine = "=" * 30 + " start of log session " + "=" * 30
            }
            else
            {
                $separatorLine = "=" * 30 + " end of log session " + "=" * 30
            }
            
            # Ensure log directory exists
            $logDir = Split-Path $LogFile -Parent
            if (-not (Test-Path $logDir))
            {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
            
            if ($OverwriteLog)
            {
                Remove-Item -Path $LogFile -Force -ErrorAction SilentlyContinue | Out-Null
            }   
            
            # Check for log rotation if file exists and is too large
            if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt ($MaxLogSizeMB * 1MB))
            {
                $archiveFile = $LogFile -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                Move-Item -Path $LogFile -Destination $archiveFile -Force
                Write-Verbose "Log file rotated to: $archiveFile"
            }
            
            if ($CMTraceFormat)
            {
                # For CMTrace format, still use the separator but in CMTrace format
                $cmTime = Get-Date -Format "HH:mm:ss.fff+000"
                $cmDate = Get-Date -Format "MM-dd-yyyy"
                $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
                $logEntry = "<![LOG[$separatorLine]LOG]!><time=`"$cmTime`" date=`"$cmDate`" component=`"$Module`" context=`"`" type=`"1`" thread=`"$thread`" file=`"`">"
            }
            else
            {
                # For standard format, just use the separator line without timestamp
                $logEntry = $separatorLine
            }
            
            # Use mutex for thread safety
            $mutexName = "Global\LogMutex_" + ($LogFile -replace '[\\/:*?"<>|]', '_')
            $mutex = $null
            $streamWriter = $null
            $fileStream = $null
            
            try
            {
                $mutex = New-Object System.Threading.Mutex($false, $mutexName)
                $mutex.WaitOne() | Out-Null
                
                # Use StreamWriter with FileShare.ReadWrite to allow concurrent access
                $fileStream = [System.IO.File]::Open(
                    $LogFile,
                    [System.IO.FileMode]::Append,
                    [System.IO.FileAccess]::Write,
                    [System.IO.FileShare]::ReadWrite
                )
                $streamWriter = New-Object System.IO.StreamWriter($fileStream, [System.Text.Encoding]::UTF8)
                $streamWriter.WriteLine($logEntry)
                $streamWriter.Flush()
            }
            catch [System.IO.IOException]
            {
                # If file is still locked, retry with exponential backoff
                $retryCount = 0
                $maxRetries = 5
                $success = $false
                
                while (-not $success -and $retryCount -lt $maxRetries)
                {
                    $retryCount++
                    Start-Sleep -Milliseconds (100 * [Math]::Pow(2, $retryCount))
                    
                    try
                    {
                        $fileStream = [System.IO.File]::Open(
                            $LogFile,
                            [System.IO.FileMode]::Append,
                            [System.IO.FileAccess]::Write,
                            [System.IO.FileShare]::ReadWrite
                        )
                        $streamWriter = New-Object System.IO.StreamWriter($fileStream, [System.Text.Encoding]::UTF8)
                        $streamWriter.WriteLine($logEntry)
                        $streamWriter.Flush()
                        $success = $true
                    }
                    catch [System.IO.IOException]
                    {
                        if ($retryCount -ge $maxRetries)
                        {
                            Write-Warning "Failed to write to log after $maxRetries retries: $($_.Exception.Message)"
                        }
                    }
                }
            }
            finally
            {
                if ($streamWriter)
                {
                    try
                    {
                        $streamWriter.Close()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to close StreamWriter: $($_.Exception.Message)"
                    }

                    try
                    {
                        $streamWriter.Dispose()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to dispose StreamWriter: $($_.Exception.Message)"
                    }
                }

                if ($fileStream)
                {
                    try
                    {
                        $fileStream.Close()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to close FileStream: $($_.Exception.Message)"
                    }

                    try
                    {
                        $fileStream.Dispose()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to dispose FileStream: $($_.Exception.Message)"
                    }
                }

                if ($mutex)
                {
                    try
                    {
                        $mutex.ReleaseMutex()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to release mutex: $($_.Exception.Message)"
                    }

                    try
                    {
                        $mutex.Dispose()
                    }
                    catch
                    {
                        Write-Warning "Write-Log: Failed to dispose mutex: $($_.Exception.Message)"
                    }
                }
            }
            
            # Write to console
            if ($WriteToConsole)
            {
                Write-Host $separatorLine
            }
            return
        }
        
        # Check if this log entry should be written based on minimum log level
        # Only continue if the current log level meets or exceeds the minimum threshold
        if (-not ($StartLogging -or $FinishLogging))
        {
            $currentLogLevelValue = $logLevelHierarchy[$LogLevel]
            $minimumLogLevelValue = $logLevelHierarchy[$MinimumLogLevel]
            
            if ($currentLogLevelValue -gt $minimumLogLevelValue)
            {
                # Current log level is more detailed than the minimum, skip logging to file
                # But still write to console streams
                switch ($LogLevel)
                {
                    "Error"
                    {
                        if ($WriteToConsole)
                        {
                            Write-Error "[$Module] $Message" -ErrorAction SilentlyContinue 
                        }
                    }
                    "Warning"
                    {
                        if ($WriteToConsole)
                        {
                            Write-Warning "[$Module] $Message" 
                        }
                    }
                    "Verbose"
                    {
                        if ($WriteToConsole)
                        {
                            Write-Verbose "[$Module] $Message" 
                        }
                    }
                    "Debug"
                    {
                        if ($WriteToConsole)
                        {
                            Write-Debug "[$Module] $Message" 
                        }
                    }
                    default
                    {
                        # For Information level, we don't output to console in this case
                    }
                }
                return
            }
        }
        
        # Ensure log directory exists
        $logDir = Split-Path $LogFile -Parent
        if (-not (Test-Path $logDir))
        {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Check for log rotation if file exists and is too large
        if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt ($MaxLogSizeMB * 1MB))
        {
            $archiveFile = $LogFile -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            Move-Item -Path $LogFile -Destination $archiveFile -Force
            Write-Verbose "Log file rotated to: $archiveFile"
        }
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
        
        # Get context in a cross-platform way
        try 
        {
            if ($IsWindows -or ($null -eq $IsWindows -and $env:OS -eq "Windows_NT"))
            {
                $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
            }
            else
            {
                $Context = $env:USER
            }
        }
        catch 
        {
            $Context = "Unknown"
        }
        
        if ($CMTraceFormat)
        {
            # True CMTrace format: 
            $cmTime = Get-Date -Format "HH:mm:ss.fff+000"
            $cmDate = Get-Date -Format "MM-dd-yyyy"
            $severity = switch ($LogLevel)
            {
                "Error"
                {
                    3 
                }
                "Warning"
                {
                    2 
                }
                default
                {
                    1 
                }
            }
            $logEntry = "<![LOG[$Message]LOG]!><time=`"$cmTime`" date=`"$cmDate`" component=`"$Module`" context=`"`" type=`"$severity`" thread=`"$thread`" file=`"`">"
        }
        else
        {
            # Enhanced standard format with thread ID
            $logEntry = "$timestamp [$LogLevel] [$Module] [Thread:$thread] [Context:$Context] $Message"
        }
        
        # Use mutex for thread safety in concurrent scenarios
        $mutexName = "Global\LogMutex_" + ($LogFile -replace '[\\/:*?"<>|]', '_')
        $mutex = $null
        $streamWriter = $null
        $fileStream = $null
        
        try
        {
            $mutex = New-Object System.Threading.Mutex($false, $mutexName)
            $mutex.WaitOne() | Out-Null
            
            # Use StreamWriter with FileShare.ReadWrite to allow concurrent access
            $fileStream = [System.IO.File]::Open(
                $LogFile,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $streamWriter = New-Object System.IO.StreamWriter($fileStream, [System.Text.Encoding]::UTF8)
            $streamWriter.WriteLine($logEntry)
            $streamWriter.Flush()
        }
        catch [System.IO.IOException]
        {
            # If file is still locked, retry with exponential backoff
            $retryCount = 0
            $maxRetries = 5
            $success = $false
            
            while (-not $success -and $retryCount -lt $maxRetries)
            {
                $retryCount++
                Start-Sleep -Milliseconds (100 * [Math]::Pow(2, $retryCount))
                
                try
                {
                    $fileStream = [System.IO.File]::Open(
                        $LogFile,
                        [System.IO.FileMode]::Append,
                        [System.IO.FileAccess]::Write,
                        [System.IO.FileShare]::ReadWrite
                    )
                    $streamWriter = New-Object System.IO.StreamWriter($fileStream, [System.Text.Encoding]::UTF8)
                    $streamWriter.WriteLine($logEntry)
                    $streamWriter.Flush()
                    $success = $true
                }
                catch [System.IO.IOException]
                {
                    if ($retryCount -ge $maxRetries)
                    {
                        Write-Warning "Failed to write to log after $maxRetries retries: $($_.Exception.Message)"
                    }
                }
            }
        }
        finally
        {
            if ($streamWriter) { $streamWriter.Close(); $streamWriter.Dispose() }
            if ($fileStream) { $fileStream.Close(); $fileStream.Dispose() }
            if ($mutex) { $mutex.ReleaseMutex(); $mutex.Dispose() }
        }
        
        # Write to appropriate PowerShell stream based on log level
        switch ($LogLevel)
        {
            "Error"
            {
                if ($WriteToConsole)
                {
                    Write-Error "[$Module] $Message" -ErrorAction SilentlyContinue 
                }
            }
            "Warning"
            {
                if ($WriteToConsole)
                {
                    Write-Warning "[$Module] $Message" 
                }
            }
            "Verbose"
            {
                if ($WriteToConsole)
                {
                    Write-Verbose "[$Module] $Message" 
                }
            }
            "Debug"
            {
                if ($WriteToConsole)
                {
                    Write-Debug "[$Module] $Message" 
                }
            }
            default
            {
                if ($WriteToConsole)
                {
                    Write-Verbose "Logged: $logEntry" 
                }
            }
        }
        
        # Return log entry if PassThru is specified
        if ($PassThru)
        {
            return [PSCustomObject]@{
                Timestamp = $timestamp
                LogLevel  = $LogLevel
                Module    = $Module
                Message   = $Message
                Thread    = $thread
                LogFile   = $LogFile
                Entry     = $logEntry
            }
        }
    }
    catch
    {
        Write-Error "Failed to write to log file '$LogFile': $_"
        # Fallback to console output
        Write-Host "$timestamp [$LogLevel] [$Module] $Message" -ForegroundColor $(
            switch ($LogLevel)
            {
                "Error"
                {
                    "Red" 
                }
                "Warning"
                {
                    "Yellow" 
                }
                "Debug"
                {
                    "Cyan" 
                }
                default
                {
                    "White" 
                }
            }
        )
    }
}
