function Close-Process()
{
    <#
.SYNOPSIS
    Closes one or more processes by name.

.DESCRIPTION
    The Close-Process function attempts to gracefully close processes by their name(s). 
    It first tries to close the main window of each process, waits for 5 seconds, and 
    if the process hasn't exited, it forcefully terminates the process using Kill().
    The function returns detailed information about which processes were closed gracefully, 
    forcefully, or failed to close.

.PARAMETER ProcessNames
    An array of process names to close. Do not include the .exe extension.

.EXAMPLE
    Close-Process -ProcessNames "notepad"
    Closes all instances of Notepad.

.EXAMPLE
    Close-Process -ProcessNames @("chrome", "firefox", "msedge")
    Closes all instances of Chrome, Firefox, and Edge browsers.

.OUTPUTS
    Returns a hashtable containing:
    - allProcessesClosed: Boolean indicating if all processes were successfully closed
    - processesToCloseCount: Total number of process names to close
    - processesClosedGracefully: Array of processes that closed gracefully
    - processesClosedGracefullyCount: Count of gracefully closed processes
    - processesForceClosed: Array of processes that were forcefully terminated
    - processesForceClosedCount: Count of forcefully closed processes
    - processesNotClosed: Array of processes that failed to close or weren't found
    - processesNotClosedCount: Count of processes not closed
    - message: Array of status messages

.NOTES
    This function uses write-log for logging. Ensure the $logFile variable is defined in the calling scope or it is accessible globally.
#>
    [CmdletBinding()                ]
    param (
        [string[]]$ProcessNames,
        [int]$timeout = 5
    )
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting to close $($ProcessNames.count) processes: $($ProcessNames -join ', ')"
    write-log -logFile $logFile -module $functionName -message "Starting to close $($ProcessNames.count) processes: $($ProcessNames -join ', ')"
    #eliminate processes not running.
    $ProcessNames = $ProcessNames | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue }                                 
    Write-Verbose "[$functionName] After filtering, $($ProcessNames.count) processes remain to be closed: $($ProcessNames -join ', ')"                  
    write-log -logFile $logFile -module $functionName -message "After filtering, $($ProcessNames.count) processes remain to be closed: $($ProcessNames -join ', ')"                             
    $returnObject = @{
        allProcessesClosed             = $false
        processesToCloseCount          = $ProcessNames.Count
        processesClosedGracefully      = @()
        processesClosedGracefullyCount = 0
        processesForceClosed           = @()
        processesForceClosedCount      = 0
        processesNotClosed             = @()
        processesNotClosedCount        = 0
        message                        = @()
    }
    # Track which process names had failures
    $failedProcessNames = @()
    foreach ($processName in $ProcessNames)
    {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        Write-Verbose "[$functionName] Found $($processes.Count) instances of process $processName"                         
        write-log -logFile $logFile -module $functionName -message "Found $($processes.Count) instances of process $processName"
        Write-Host "Closing $($processes.Count) instances of process $processName"      
        $processNameHadFailure = $false
        if ($processes)
        {
            Write-Verbose "[$functionName] Closing process: $processName"           
            write-log -logFile $logFile -module $functionName -message "Closing process: $processName"                                  
            foreach ($process in $processes)
            {
                Write-Verbose "[$functionName] Attempting to close process: $($process.Name) (ID: $($process.Id))"
                write-log -logFile $logFile -module $functionName -message "Attempting to close process: $($process.Name) (ID: $($process.Id))"                                                                         
                try
                {
                    Write-Verbose "[$functionName] Closing process: $($process.Name) (ID: $($process.Id))"                      
                    write-log -logFile $logFile -module $functionName -message "Closing process: $($process.Name) (ID: $($process.Id))"                                 
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Seconds $timeout       
                    if (!$process.HasExited)
                    {
                        Write-Host "Forcefully terminating process: $($process.Name) (ID: $($process.Id))"
                        write-log -logFile $logFile -module $functionName -message "Forcefully terminating process: $($process.Name) (ID: $($process.Id))"                                          
                        $process.Kill()
                        $returnObject.processesForceClosed += $process.Name                 
                        $returnObject.processesForceClosedCount++
                    }
                    else
                    {
                        Write-Host "Process closed gracefully: $($process.Name) (ID: $($process.Id))"
                        write-log -logFile $logFile -module $functionName -message "Process closed gracefully: $($process.Name) (ID: $($process.Id))"
                        $returnObject.processesClosedGracefully += $process.Name
                        $returnObject.processesClosedGracefullyCount++                        
                    }                               
                }
                catch
                {
                    Write-Warning "Failed to close process: $($process.Name) (ID: $($process.Id)). Error: $_"
                    $processNameHadFailure = $true
                    $returnObject.message += "Failed to close process: $($process.Name) (ID: $($process.Id)). Error: $_"                    
                }
            }
        }
        else
        {
            Write-Host "No running process found with name: $processName"
            write-log -logFile $logFile -module $functionName -message "No running process found with name: $processName"
            # Not running is not considered a failure
        }
        
        # After processing all instances of this process name, check if any failed
        if ($processNameHadFailure)
        {
            Write-Verbose "[$functionName] Process name $processName had failures during closing attempts." 
            write-log -logFile $logFile -module $functionName -message "Process name $processName had failures during closing attempts." -logLevel "Warning "
            $failedProcessNames += $processName
            if ($processName -notin $returnObject.processesNotClosed)
            {
                Write-Verbose "[$functionName] Adding process name $processName to processesNotClosed list."
                write-log -logFile $logFile -module $functionName -message "Adding process name $processName to processesNotClosed list." -logLevel "Information"   
                $returnObject.processesNotClosed += $processName
                $returnObject.processesNotClosedCount++
            }
        }
    }
    
    # Determine overall success: all processes closed if no failures occurred
    if ($failedProcessNames.Count -eq 0)
    {
        Write-Verbose "[$functionName] All processes closed successfully."
        $returnObject.message += "All processes closed successfully."                   
        $returnObject.allProcessesClosed = $true
        write-log -logFile $logFile -module $functionName -message "All processes closed successfully."
    }
    else
    {
        Write-Verbose "[$functionName] Some processes failed to close."             
        $returnObject.message += "Failed to close the following processes: $($failedProcessNames -join ', ')"
        $returnObject.allProcessesClosed = $false
        write-log -logFile $logFile -module $functionName -message "Failed to close the following processes: $($failedProcessNames -join ', ')"
    }                                           
    return $returnObject                                
}                                               

