function Invoke-ExternalProcess()
{
    <#
.SYNOPSIS
    Starts an external process, optionally waits for spawned processes to exit with a timeout, and returns status.

.DESCRIPTION
    The Invoke-ExternalProcess function launches an external executable with optional arguments, monitors its execution, and can wait for the process (and its children) to exit. It supports retry logic based on process output or exit code, and provides detailed status and logging. The function can use either direct invocation or PowerShell's native Start-Process cmdlet for process creation.

.PARAMETER FilePath
    The path to the executable to run.

.PARAMETER ArgumentList
    Arguments as a single string (will be split on spaces unless UseNativeStartProcess is specified).

.PARAMETER ProcessName
    Optional process name to use for detection/waiting; defaults to the executable name (without extension).

.PARAMETER SpecialRegExp
    Optional regex to detect a condition that triggers stopping matching processes and retrying the command.

.PARAMETER specialExitCode
    Optional exit code that triggers the same retry logic as SpecialRegExp. Default: 255.

.PARAMETER waitForExit
    When set, attempts to detect the newly started process IDs for ProcessName and waits until they exit.

.PARAMETER timeout
    Maximum seconds to wait when -waitForExit is set (default 1800). 0 or less waits indefinitely.

.PARAMETER UseNativeStartProcess
    When set, uses PowerShell's Start-Process cmdlet for process creation and output redirection.

.OUTPUTS
    PSCustomObject with:
        FilePath, Arguments, exitCode, exitCodeDescription, commandOutput, stoppedProcesses (array), success (bool),
        waitedOnProcessIds (array of Int32), timedOut (bool)
    If the executable cannot be launched (not found, access denied, etc.), exitCode will be non-zero,
    success will be $false, and commandOutput will contain the error message. No waiting is attempted.
    
    For msiexec.exe processes, exitCodeDescription will contain the meaning of the MSI exit code,
    and success will always be $true regardless of the exit code, as MSI operations may have
    various exit codes that still indicate successful completion (e.g., reboot required).

.EXAMPLE
    Invoke-ExternalProcess -FilePath "notepad.exe" -waitForExit

.NOTES
    Author: Zuhair Mahmoud
    Date: 06/12/2025
    This function is intended for use in automation scenarios where robust process launching and monitoring is required.
#>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter()]
        [string]$ArgumentList,
        [string]$ProcessName,
        [string]$SpecialRegExp,
        [int]$specialExitCode = 255,
        [switch]$waitForExit,
        [switch]$UseNativeStartProcess,
        [int]$timeout = 0
    )

    $functionName = $MyInvocation.MyCommand.Name
    #region write verbose and log received parameters.
    Write-Verbose "[$functionName] Received parameters:"
    Write-Verbose "[$functionName] FilePath: $FilePath"
    Write-Verbose "[$functionName] ArgumentList: $ArgumentList"
    Write-Verbose "[$functionName] SpecialExitCode: $specialExitCode"
    Write-Verbose "[$functionName] ProcessName: $ProcessName"
    Write-Verbose "[$functionName] SpecialRegExp: $SpecialRegExp"
    Write-Verbose "[$functionName] UseNativeStartProcess: $UseNativeStartProcess"
    Write-Verbose "[$functionName] waitForExit: $waitForExit"
    write-log -logFile $LogFile -module $functionName -Message "Received parameters: FilePath=$FilePath, ArgumentList=$ArgumentList, ProcessName=$ProcessName, SpecialRegExp=$SpecialRegExp, SpecialExitCode=$specialExitCode, UseNativeStartProcess=$UseNativeStartProcess, waitForExit=$waitForExit" -LogLevel "Information"
    #endregion

    # Prepare arguments array if provided
    $argsText = if ($null -ne $ArgumentList)
    {
        if (-not $UseNativeStartProcess)
        {
            Write-Verbose "[$functionName] ArgumentList is a string. Converting to an array."
            write-log -logFile $LogFile -module $functionName -Message "ArgumentList is a string. Converting to an array." -LogLevel "Verbose"
            @($ArgumentList -split ' ')
        }
        else
        {
            Write-Verbose "[$functionName] Using native Powershell process. Using as-is."
            write-log -logFile $LogFile -module $functionName -Message "Using Powershell native start process. Using as-is." -LogLevel "Information"
            $ArgumentList
        }
    }
    else
    {
        Write-Verbose "[$functionName] No arguments provided."
        write-log -LogFile $LogFile -Module $functionName -Message "No arguments provided." -LogLevel "Information"
    }
    Write-Verbose "[$functionName] Processed $($argsText.Count) arguments."
    write-log -logFile $LogFile -module $functionName -Message "Processed $($argsText.Count) arguments." -LogLevel "Verbose"
    
    # Resolve process name for detection/wait
    $procName = if ($ProcessName)
    {
        $ProcessName 
    }
    else
    {
        [System.IO.Path]::GetFileNameWithoutExtension($FilePath) 
    }
    
    # Return object scaffold
    $returnObject = [PSCustomObject]@{
        FilePath            = $FilePath
        Arguments           = $argsText
        exitCode            = $null
        exitCodeDescription = $null
        commandOutput       = $null
        stoppedProcesses    = @()
        success             = $false
        errorMessage        = $null
        waitedOnProcessIds  = @()
        timedOut            = $false
    }

    # Preflight: attempt to resolve the executable and detect obvious not-found conditions
    $skipInvoke = $false
    $exeForCheck = $FilePath.Trim('"')
    $resolved = $null
    try
    {
        $resolved = Get-Command -Name $exeForCheck -ErrorAction Stop 
    }
    catch
    {
        $resolved = $null 
    }
    if (-not $resolved)
    {
        $looksLikePath = ($exeForCheck -match '[\\/]' -or $exeForCheck -match '\.(exe|cmd|bat|ps1|msi)$')
        if ($looksLikePath -and -not (Test-Path -LiteralPath $exeForCheck -PathType Leaf))
        {
            $skipInvoke = $true
            $returnObject.exitCode = 2  # ERROR_FILE_NOT_FOUND
            $returnObject.commandOutput = "Executable not found: $FilePath"
            $returnObject.success = $false
            Write-Error "[$functionName] Executable not found: $FilePath"
            write-log -logFile $LogFile -module $functionName -Message "Executable not found: $FilePath" -LogLevel "Error"
        }
    }
    $invokeSucceeded = $false
    $proc = $null
    $lastPreStartPids = @()
    $lastLaunchStart = $null
    if (-not $UseNativeStartProcess)
    {
        if (-not $skipInvoke)
        {
            # Snapshot existing PIDs to detect new ones created by the launch
            $preStartPids = @()
            try
            {
                $preStartPids = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) 
            }
            catch
            {
                $preStartPids = @() 
            }
            $launchStart = Get-Date
            Write-Log -LogFile $LogFile -Module $scriptName -Message "Starting process: $FilePath $argsText" -LogLevel "Information"
            Write-Verbose "[$functionName] Starting process: $FilePath with arguments $argsText"
            try
            {
                # Do not change how the process is launched
                $proc = & $FilePath $argsText
                $invokeSucceeded = $true
            }
            catch
            {
                $invokeSucceeded = $false
                $nativeCode = $null
                if ($_.Exception -is [System.ComponentModel.Win32Exception])
                {
                    $nativeCode = $_.Exception.NativeErrorCode 
                }
                elseif ($_.Exception.InnerException -is [System.ComponentModel.Win32Exception])
                {
                    $nativeCode = $_.Exception.InnerException.NativeErrorCode 
                }
                $errMsg = $_.Exception.Message
                $returnObject.exitCode = if ($nativeCode)
                {
                    [int]$nativeCode 
                }
                else
                {
                    1 
                }
                $returnObject.commandOutput = $errMsg
                $returnObject.success = $false
                Write-Error "[$functionName] Failed to launch '$FilePath': $errMsg"
                write-log -logFile $LogFile -module $functionName -Message "Failed to launch '$FilePath': $errMsg (Exit=$($returnObject.exitCode))" -LogLevel "Error"
            }
            # Track snapshot info for waiting logic
            $lastPreStartPids = $preStartPids
            $lastLaunchStart = $launchStart
            if ($invokeSucceeded)
            {
                Write-Verbose "[$functionName] Command output: $proc"
                write-log -logFile $LogFile -module $functionName -Message "Command output: $proc" -LogLevel "Information"
                Write-Verbose "[$functionName] Exit code: $lastExitCode"
                write-log -logFile $LogFile -module $functionName -Message "Exit code: $lastExitCode" -LogLevel "Information"
            }
        }
        # If requested, wait for the detected processes to exit (only if launch succeeded)
        if ($waitForExit -and $invokeSucceeded)
        {
            Write-Verbose "[$functionName] -waitForExit specified. Beginning detection and wait for '$procName'."
            write-log -logFile $LogFile -module $functionName -Message "-waitForExit specified. Detecting new '$procName' PIDs and waiting with timeout=$timeout seconds." -LogLevel "Information"
            # Discover newly created PIDs for procName since last launch
            $discoveryDeadline = (Get-Date).AddSeconds(10)
            $newPids = @()
            $discoveryPass = 0
            do
            {
                $discoveryPass++
                $candidates = @()
                try
                {
                    $candidates = @(Get-Process -Name $procName -ErrorAction SilentlyContinue) 
                }
                catch
                {
                    $candidates = @() 
                }
                Write-Verbose "[$functionName] Discovery pass #$($discoveryPass): found $($candidates.Count) candidate process(es) for '$procName'."
                if ($candidates.Count -gt 0)
                {
                    $filtered = @()
                    foreach ($c in $candidates)
                    {
                        $include = $true
                        $stStr = ''
                        try
                        {
                            $stStr = $c.StartTime
                            if ($null -ne $lastLaunchStart -and $c.StartTime -lt $lastLaunchStart)
                            {
                                $include = $false 
                            }
                        }
                        catch
                        { 
                            # some system processes throw when accessing StartTime
                        }
                        if ($include -and ($lastPreStartPids -contains $c.Id))
                        {
                            $include = $false 
                        }
                        Write-Verbose "[$functionName] Candidate PID=$($c.Id) Name=$($c.ProcessName) StartTime=$stStr Include=$include"
                        if ($include)
                        {
                            $filtered += $c.Id 
                        }
                    }
                    $newPids = @($filtered | Select-Object -Unique)
                }

                if ($newPids.Count -eq 0)
                {
                    Start-Sleep -Milliseconds 300 
                }
            } while ($newPids.Count -eq 0 -and (Get-Date) -lt $discoveryDeadline)
            if ($newPids.Count -gt 0)
            {
                Write-Verbose "[$functionName] Detected new '$procName' PIDs to wait on: $($newPids -join ', ')"
                write-log -logFile $LogFile -module $functionName -Message "Detected new '$procName' PIDs: $($newPids -join ', ')" -LogLevel "Information"
                # Track all PIDs to wait on and expand with child processes spawned by these
                $trackedPids = @($newPids | Select-Object -Unique)
                $returnObject.waitedOnProcessIds = $trackedPids

                $deadline = $null
                if ($timeout -gt 0)
                {
                    $deadline = (Get-Date).AddSeconds([int]$timeout) 
                }

                $tickLast = Get-Date
                $childDiscoverLast = Get-Date
                while ($true)
                {
                    $running = @()
                    try
                    {
                        $running = @(Get-Process -Id $trackedPids -ErrorAction SilentlyContinue) 
                    }
                    catch
                    {
                        $running = @() 
                    }

                    # Periodically look for child processes of any still-running tracked PIDs
                    if ($running.Count -gt 0 -and (Get-Date) -ge $childDiscoverLast.AddSeconds(5))
                    {
                        $childDiscoverLast = Get-Date
                        try
                        {
                            $allProcs = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue)
                        }
                        catch
                        {
                            $allProcs = @()
                        }
                        if ($allProcs.Count -gt 0)
                        {
                            $runningIds = @($running | Select-Object -ExpandProperty Id)
                            $newChildren = @()
                            foreach ($wp in $allProcs)
                            {
                                $ppid = [int]$wp.ParentProcessId
                                if ($runningIds -contains $ppid)
                                {
                                    # Filter by creation time >= lastLaunchStart when available
                                    $ok = $true
                                    try
                                    {
                                        if ($lastLaunchStart)
                                        {
                                            $created = [System.Management.ManagementDateTimeConverter]::ToDateTime($wp.CreationDate)
                                            if ($created -lt $lastLaunchStart)
                                            {
                                                $ok = $false 
                                            }
                                        }
                                    }
                                    catch
                                    { 
                                    }
                                    if ($ok -and -not ($trackedPids -contains [int]$wp.ProcessId))
                                    {
                                        $newChildren += [int]$wp.ProcessId
                                        Write-Verbose "[$functionName] Adding child PID=$([int]$wp.ProcessId) Name=$($wp.Name) ParentPID=$ppid to tracking set."
                                        write-log -logFile $LogFile -module $functionName -Message "Tracking child PID=$([int]$wp.ProcessId) Name=$($wp.Name) ParentPID=$ppid" -LogLevel "Verbose"
                                    }
                                }
                            }
                            if ($newChildren.Count -gt 0)
                            {
                                $trackedPids = @(($trackedPids + $newChildren) | Select-Object -Unique)
                                $returnObject.waitedOnProcessIds = $trackedPids
                                Write-Verbose "[$functionName] Expanded tracking set to PIDs: $($trackedPids -join ', ')"
                            }
                        }
                    }

                    if ($running.Count -eq 0)
                    {
                        break 
                    }

                    if ($null -ne $deadline -and (Get-Date) -ge $deadline)
                    {
                        $returnObject.timedOut = $true
                        Write-Verbose "[$functionName] Timed out waiting for '$procName' after $timeout seconds. Still running PIDs: $($running.Id -join ', ')"
                        write-log -logFile $LogFile -module $functionName -Message "Timed out waiting for '$procName' after $timeout seconds. Still running PIDs: $($running.Id -join ', ')" -LogLevel "Warning"
                        break
                    }

                    if ((Get-Date) -ge $tickLast.AddSeconds(10))
                    {
                        $tickLast = Get-Date
                        Write-Verbose "[$functionName] Waiting on '$procName'... still running PIDs: $($running.Id -join ', ')"
                        write-log -logFile $LogFile -module $functionName -Message "Waiting on '$procName'... still running PIDs: $($running.Id -join ', ')" -LogLevel "Verbose"
                    }

                    Start-Sleep -Seconds 1
                }

                if (-not $returnObject.timedOut)
                {
                    Write-Verbose "[$functionName] All detected '$procName' processes (including children) have exited."
                    write-log -logFile $LogFile -module $functionName -Message "All detected '$procName' processes (including children) have exited." -LogLevel "Information"
                }
            }
            else
            {
                Write-Verbose "[$functionName] No new '$procName' processes detected to wait on. Likely completed inline."
                write-log -logFile $LogFile -module $functionName -Message "No new '$procName' processes detected to wait on."
            }
        }

        # Retry logic if special condition detected (only if we had a successful launch)
        if ($invokeSucceeded -and ($proc -match $SpecialRegExp -or $lastExitCode -eq $specialExitCode))
        {
            Write-Verbose "[$functionName] Special condition matched (regex/exitCode)."
            write-log -logFile $LogFile -module $functionName -Message "Special condition matched (regex/exitCode)." -LogLevel "Information"

            $processToStop = Get-FileHandle -SearchString $procName
            if ($processToStop -and $processToStop.Count -gt 0)
            {
                Write-Verbose "[$functionName] Found processes to stop: $($processToStop | Format-Table -AutoSize | Out-String)"
                write-log -logFile $LogFile -module $functionName -Message "Found processes to stop: $($processToStop | Format-Table -AutoSize | Out-String)" -LogLevel "Information"

                $stoppedProcesses = @()
                foreach ($process in $processToStop)
                {
                    $processObject = [PSCustomObject]@{ ProcessName = $process.ProcessName; ProcessId = $process.ProcessId }
                    $stoppedProcesses += $processObject
                    Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
                    Write-Verbose "[$functionName] Stopped process: $($process.ProcessName) (PID: $($process.ProcessId))"
                    write-log -logFile $LogFile -module $functionName -Message "Stopped process: $($process.ProcessName) (PID: $($process.ProcessId))" -LogLevel "Information"
                }
                $returnObject.stoppedProcesses = $stoppedProcesses

                Write-Verbose "[$functionName] Attempting to run the process again."
                write-log -logFile $LogFile -module $functionName -Message "Attempting to run the process again." -LogLevel "Information"

                # Refresh snapshot for the re-run
                $preStartPids2 = @()
                try
                {
                    $preStartPids2 = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) 
                }
                catch
                {
                    $preStartPids2 = @() 
                }
                $launchStart2 = Get-Date

                try
                {
                    $proc = & $FilePath $argsText
                    $invokeSucceeded = $true
                }
                catch
                {
                    $invokeSucceeded = $false
                    $nativeCode = $null
                    if ($_.Exception -is [System.ComponentModel.Win32Exception])
                    {
                        $nativeCode = $_.Exception.NativeErrorCode 
                    }
                    elseif ($_.Exception.InnerException -is [System.ComponentModel.Win32Exception])
                    {
                        $nativeCode = $_.Exception.InnerException.NativeErrorCode 
                    }
                    $errMsg = $_.Exception.Message
                    $returnObject.exitCode = if ($nativeCode)
                    {
                        [int]$nativeCode 
                    }
                    else
                    {
                        1 
                    }
                    $returnObject.commandOutput = $errMsg
                    $returnObject.success = $false
                    Write-Error "[$functionName] Failed to relaunch '$FilePath': $errMsg"
                    write-log -logFile $LogFile -module $functionName -Message "Failed to relaunch '$FilePath': $errMsg (Exit=$($returnObject.exitCode))" -LogLevel "Error"
                }

                $lastPreStartPids = $preStartPids2
                $lastLaunchStart = $launchStart2

                Write-Verbose "[$functionName] Command output: $proc"
                write-log -logFile $LogFile -module $functionName -Message "Command output: $proc" -LogLevel "Information"
                Write-Verbose "[$functionName] Exit code: $lastExitCode"
                write-log -logFile $LogFile -module $functionName -Message "Exit code: $lastExitCode" -LogLevel "Information"
            }
        }
    }
    if ($UseNativeStartProcess -and -not $skipInvoke)
    {
        Write-Verbose "[$functionName] Using native Start-Process to run the command."
        write-log -logFile $LogFile -module $functionName -Message "Using native Start-Process to run the command." -LogLevel "Information"
        try
        {
            $stdOutTempFile = "$env:TEMP\$((New-Guid).Guid)"
            Write-Verbose "[$functionName] Redirecting standard output to [$stdOutTempFile]"
            write-log -logFile $logFile -module $functionName -message "Redirecting standard output to [$stdOutTempFile]"
            $stdErrTempFile = "$env:TEMP\$((New-Guid).Guid)"
            Write-Verbose "[$functionName] Redirecting standard error to [$stdErrTempFile]"
            write-log -logFile $logFile -module $functionName -message "Redirecting standard error to [$stdErrTempFile]"
            $startProcessParams = @{
                FilePath               = $FilePath
                ArgumentList           = $argsText
                RedirectStandardError  = $stdErrTempFile
                RedirectStandardOutput = $stdOutTempFile
                Wait                   = $false
                PassThru               = $true
            }
            if ($PSCmdlet.ShouldProcess("Process [$($FilePath)]", "Run with args: [$($ArgumentList)]"))
            {
                Write-Verbose "[$functionName] Starting process with parameters: $($startProcessParams | Out-String)"
                write-log -logFile $logFile -module $functionName -message "Starting process with parameters: $($startProcessParams | Out-String)"
                $cmd = Start-Process @startProcessParams
                Write-Verbose "[$functionName] Process started with ID: $($cmd.Id)"
                write-log -logFile $logFile -module $functionName -message "Process started with ID: $($cmd.Id)"
                
                if ($waitForExit)
                {
                    # Wait for process to exit with timeout enforcement
                    $counter = 0
                    $lastOutputSize = 0
                    $lastErrorSize = 0
                    $effectiveTimeout = if ($timeout -le 0) { [int]::MaxValue } else { $timeout }
                    
                    Write-Verbose "[$functionName] Waiting for process to exit with timeout of $effectiveTimeout seconds..."
                    write-log -logFile $logFile -module $functionName -message "Waiting for process to exit with timeout of $effectiveTimeout seconds..."
                    
                    while (-not $cmd.HasExited -and $counter -lt $effectiveTimeout)
                    {
                        if ($counter % 10 -eq 0)
                        {
                            $elapsedTime = [TimeSpan]::FromSeconds($counter)
                            Write-Verbose "[$functionName] Waiting for process to exit... Elapsed time: $([int]$elapsedTime.TotalMinutes)m $($elapsedTime.Seconds)s."
                            write-log -logFile $logFile -module $functionName -message "Waiting for process to exit... Elapsed time: $counter seconds." -LogLevel "Verbose"
                        }
                        Start-Sleep -Seconds 1
                        $counter++
                        
                        # Check for new stdout content
                        if (Test-Path -Path $stdOutTempFile -PathType Leaf)
                        {
                            try
                            {
                                $currentContent = Get-Content -Path $stdOutTempFile -Raw -ErrorAction SilentlyContinue
                                if ($currentContent -and $currentContent.Length -gt $lastOutputSize)
                                {
                                    $newOutput = $currentContent.Substring($lastOutputSize)
                                    $lastOutputSize = $currentContent.Length
                                    if (-not [string]::IsNullOrWhiteSpace($newOutput))
                                    {
                                        Write-Host $newOutput -NoNewline
                                        write-log -logFile $logFile -module $functionName -message "Process output: $newOutput" -LogLevel "Information"
                                    }
                                }
                            }
                            catch
                            {
                                Write-Verbose "[$functionName] Could not read stdout: $($_.Exception.Message)"
                            }
                        }
                        
                        # Check for new stderr content
                        if (Test-Path -Path $stdErrTempFile -PathType Leaf)
                        {
                            try
                            {
                                $currentError = Get-Content -Path $stdErrTempFile -Raw -ErrorAction SilentlyContinue
                                if ($currentError -and $currentError.Length -gt $lastErrorSize)
                                {
                                    $newError = $currentError.Substring($lastErrorSize)
                                    $lastErrorSize = $currentError.Length
                                    if (-not [string]::IsNullOrWhiteSpace($newError))
                                    {
                                        Write-Warning $newError
                                        write-log -logFile $logFile -module $functionName -message "Process error: $newError" -LogLevel "Warning"
                                    }
                                }
                            }
                            catch
                            {
                                Write-Verbose "[$functionName] Could not read stderr: $($_.Exception.Message)"
                            }
                        }
                    }
                    
                    if ($counter -ge $effectiveTimeout -and -not $cmd.HasExited)
                    {
                        Write-Verbose "[$functionName] Process timed out after $timeout seconds. Attempting to stop process and children."
                        write-log -logFile $logFile -module $functionName -message "Process timed out after $timeout seconds. Attempting to stop process and children."                       
                        try
                        {
                            # Stop the main process
                            $cmd | Stop-Process -Force -ErrorAction Stop
                            Write-Verbose "[$functionName] Main process (PID: $($cmd.Id)) stopped due to timeout."
                            write-log -logFile $logFile -module $functionName -message "Main process (PID: $($cmd.Id)) stopped due to timeout."
                            
                            # Also attempt to stop any child processes
                            try
                            {
                                $children = Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.ParentProcessId -eq $cmd.Id }
                                foreach ($child in $children)
                                {
                                    Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue
                                    Write-Verbose "[$functionName] Stopped child process (PID: $($child.ProcessId))."
                                    write-log -logFile $logFile -module $functionName -message "Stopped child process (PID: $($child.ProcessId))."
                                }
                            }
                            catch
                            {
                                Write-Verbose "[$functionName] Could not enumerate/stop child processes: $($_.Exception.Message)"
                                write-log -logFile $logFile -module $functionName -message "Could not enumerate/stop child processes: $($_.Exception.Message)"
                            }                       
                        }
                        catch
                        {
                            Write-Error "[$functionName] Failed to stop process after timeout: $($_.Exception.Message)"
                            write-log -logFile $logFile -module $functionName -message "Failed to stop process after timeout: $($_.Exception.Message)" -LogLevel "Error"                       
                        }
                        $returnObject.timedOut = $true
                    }                                                                   
                    else 
                    {
                        Write-Verbose "[$functionName] Process completed within timeout."
                        write-log -logFile $logFile -module $functionName -message "Process completed within timeout."
                        
                        # Capture any remaining output after process exit
                        if (Test-Path -Path $stdOutTempFile -PathType Leaf)
                        {
                            try
                            {
                                $finalContent = Get-Content -Path $stdOutTempFile -Raw -ErrorAction SilentlyContinue
                                if ($finalContent -and $finalContent.Length -gt $lastOutputSize)
                                {
                                    $remainingOutput = $finalContent.Substring($lastOutputSize)
                                    if (-not [string]::IsNullOrWhiteSpace($remainingOutput))
                                    {
                                        Write-Host $remainingOutput -NoNewline
                                        write-log -logFile $logFile -module $functionName -message "Final process output: $remainingOutput" -LogLevel "Information"
                                    }
                                }
                            }
                            catch
                            {
                                Write-Verbose "[$functionName] Could not read final stdout: $($_.Exception.Message)"
                            }
                        }
                        
                        if (Test-Path -Path $stdErrTempFile -PathType Leaf)
                        {
                            try
                            {
                                $finalError = Get-Content -Path $stdErrTempFile -Raw -ErrorAction SilentlyContinue
                                if ($finalError -and $finalError.Length -gt $lastErrorSize)
                                {
                                    $remainingError = $finalError.Substring($lastErrorSize)
                                    if (-not [string]::IsNullOrWhiteSpace($remainingError))
                                    {
                                        Write-Warning $remainingError
                                        write-log -logFile $logFile -module $functionName -message "Final process error: $remainingError" -LogLevel "Warning"
                                    }
                                }
                            }
                            catch
                            {
                                Write-Verbose "[$functionName] Could not read final stderr: $($_.Exception.Message)"
                            }
                        }
                    }
                    $invokeSucceeded = $true
                }
                else
                {
                    # Not waiting for exit - process runs in background
                    Write-Verbose "[$functionName] Process started in background (not waiting for exit)."
                    write-log -logFile $logFile -module $functionName -message "Process started in background (not waiting for exit)."
                    $invokeSucceeded = $true
                }
                
                # Capture full output for return object
                $cmdOutput = Get-Content -Path $stdOutTempFile -Raw -ErrorAction SilentlyContinue
                $cmdError = Get-Content -Path $stdErrTempFile -Raw -ErrorAction SilentlyContinue
                
                # Determine exit code - if process timed out, use special code
                if ($returnObject.timedOut)
                {
                    $lastExitCode = -1  # Special exit code for timeout
                    Write-Verbose "[$functionName] Process was terminated due to timeout. Exit code set to -1."
                    write-log -logFile $logFile -module $functionName -message "Process was terminated due to timeout. Exit code set to -1." -LogLevel "Warning"
                    
                    # Display any remaining output from timed out process
                    if ($cmdOutput -or $cmdError)
                    {
                        Write-Verbose "[$functionName] Full captured output from timed out process (length: $($cmdOutput.Length) chars)"
                        Write-Verbose "[$functionName] Full captured error from timed out process (length: $($cmdError.Length) chars)"
                    }
                }
                else
                {
                    $lastExitCode = if ($null -ne $cmd.ExitCode) { $cmd.ExitCode } elseif ($cmdError) { 100 } else { 0 }
                }
                
                Write-Verbose "[$functionName] Total process output captured: $($cmdOutput.Length) bytes"
                Write-Verbose "[$functionName] Total process error captured: $($cmdError.Length) bytes"
                Write-Verbose "[$functionName] Process exit code: $lastExitCode"
                write-log -logFile $logFile -module $functionName -message "Process exit code: $lastExitCode"
                
                if ($lastExitCode -ne 0)
                {
                    if ($returnObject.timedOut)
                    {
                        Write-Host "Process [$($FilePath)] timed out after $timeout seconds and was terminated." -ForegroundColor Red
                        write-log -logFile $logFile -module $functionName -message "Process [$($FilePath)] timed out after $timeout seconds and was terminated." -LogLevel "Error"
                    }
                    else
                    {
                        Write-Host "Process [$($FilePath)] exited with code [$lastExitCode]." -ForegroundColor Red
                        write-log -logFile $logFile -module $functionName -message "Process [$($FilePath)] exited with code [$lastExitCode]." -LogLevel "Error"
                        if ($cmdError)
                        {
                            write-log -logFile $logFile -module $functionName -message "Process error details: $cmdError" -LogLevel "Error"
                        }
                    }
                }
                else
                {
                    Write-Verbose "[$functionName] Process completed successfully with exit code 0."
                    write-log -logFile $logFile -module $functionName -message "Process completed successfully with exit code 0." -LogLevel "Information"
                }
            }
        }
        catch
        {
            $invokeSucceeded = $false
            $returnObject.exitCode = 1
            $returnObject.commandOutput = $_.Exception.Message
            $returnObject.success = $false
            Write-Error "[$functionName] Failed to launch using Start-Process: $($_.Exception.Message)"
            write-log -logFile $LogFile -module $functionName -Message "Failed to launch using Start-Process: $($_.Exception.Message)" -LogLevel "Error"
        }
        finally
        {
            Write-Verbose "[$functionName] Cleaning up temporary files."
            write-log -logFile $logFile -module $functionName -message "Cleaning up temporary files." -LogLevel "Information"
            Remove-Item -Path $stdOutTempFile, $stdErrTempFile -Force -ErrorAction Ignore
            Write-Verbose "[$functionName] Temporary files cleaned up."
            write-log -logFile $logFile -module $functionName -message "Temporary files cleaned up." -LogLevel "Information"
        }
    }
    # Check if the process is msiexec.exe and handle its specific exit codes
    $isMsiexec = $false
    $executableName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    if ($executableName -eq "msiexec")
    {
        $isMsiexec = $true
        Write-Verbose "[$functionName] Detected msiexec.exe process, will handle MSI-specific exit codes."
        write-log -logFile $LogFile -module $functionName -Message "Detected msiexec.exe process, will handle MSI-specific exit codes." -LogLevel "Information"
    }
    # Define comprehensive MSI exit codes and their meanings
    $msiExitCodes = @{
        0    = "ERROR_SUCCESS - The action completed successfully"
        13   = "ERROR_INVALID_DATA - The data is invalid"
        87   = "ERROR_INVALID_PARAMETER - One of the parameters was invalid"
        120  = "ERROR_CALL_NOT_IMPLEMENTED - This value is returned when a custom action attempts to call a function that cannot be called from custom actions"
        1259 = "ERROR_APPHELP_BLOCK - If Windows Installer determines a product may be incompatible with the current operating system, it displays a dialog informing the user and asking whether to try to install anyway"
        1601 = "ERROR_INSTALL_SERVICE_FAILURE - The Windows Installer service could not be accessed"
        1602 = "ERROR_INSTALL_USEREXIT - User cancelled installation"
        1603 = "ERROR_INSTALL_FAILURE - Fatal error during installation"
        1604 = "ERROR_INSTALL_SUSPEND - Installation suspended, incomplete"
        1605 = "ERROR_UNKNOWN_PRODUCT - This action is only valid for products that are currently installed"
        1606 = "ERROR_UNKNOWN_FEATURE - Feature ID not registered"
        1607 = "ERROR_UNKNOWN_COMPONENT - Component ID not registered"
        1608 = "ERROR_UNKNOWN_PROPERTY - Unknown property"
        1609 = "ERROR_INVALID_HANDLE_STATE - Handle is in an invalid state"
        1610 = "ERROR_BAD_CONFIGURATION - The configuration data for this product is corrupt"
        1611 = "ERROR_INDEX_ABSENT - Component qualifier not present"
        1612 = "ERROR_INSTALL_SOURCE_ABSENT - The installation source for this product is not available"
        1613 = "ERROR_INSTALL_PACKAGE_VERSION - This installation package cannot be installed by the Windows Installer service"
        1614 = "ERROR_PRODUCT_UNINSTALLED - Product is uninstalled"
        1615 = "ERROR_BAD_QUERY_SYNTAX - SQL query syntax invalid or unsupported"
        1616 = "ERROR_INVALID_FIELD - Record field does not exist"
        1618 = "ERROR_INSTALL_ALREADY_RUNNING - Another installation is already in progress"
        1619 = "ERROR_INSTALL_PACKAGE_OPEN_FAILED - This installation package could not be opened"
        1620 = "ERROR_INSTALL_PACKAGE_INVALID - This installation package is invalid"
        1621 = "ERROR_INSTALL_UI_FAILURE - There was an error starting the Windows Installer service user interface"
        1622 = "ERROR_INSTALL_LOG_FAILURE - Error opening installation log file"
        1623 = "ERROR_INSTALL_LANGUAGE_UNSUPPORTED - This language of this installation package is not supported by your system"
        1624 = "ERROR_INSTALL_TRANSFORM_FAILURE - Error applying transforms"
        1625 = "ERROR_INSTALL_PACKAGE_REJECTED - This installation is forbidden by system policy"
        1626 = "ERROR_FUNCTION_NOT_CALLED - Function could not be executed"
        1627 = "ERROR_FUNCTION_FAILED - Function failed during execution"
        1628 = "ERROR_INVALID_TABLE - Invalid or unknown table specified"
        1629 = "ERROR_DATATYPE_MISMATCH - Data supplied is of wrong type"
        1630 = "ERROR_UNSUPPORTED_TYPE - Data of this type is not supported"
        1631 = "ERROR_CREATE_FAILED - The Windows Installer service failed to start"
        1632 = "ERROR_INSTALL_TEMP_UNWRITABLE - The Temp folder is on a drive that is full or inaccessible"
        1633 = "ERROR_INSTALL_PLATFORM_UNSUPPORTED - This installation package is not supported by this processor type"
        1634 = "ERROR_INSTALL_NOTUSED - Component not used on this computer"
        1635 = "ERROR_PATCH_PACKAGE_OPEN_FAILED - This update package could not be opened"
        1636 = "ERROR_PATCH_PACKAGE_INVALID - This update package is invalid"
        1637 = "ERROR_PATCH_PACKAGE_UNSUPPORTED - This update package cannot be processed by the Windows Installer service"
        1638 = "ERROR_PRODUCT_VERSION - Another version of this product is already installed"
        1639 = "ERROR_INVALID_COMMAND_LINE - Invalid command line argument"
        1640 = "ERROR_INSTALL_REMOTE_DISALLOWED - Only administrators have permission to add, remove, or configure server software during a Terminal services remote session"
        1641 = "ERROR_SUCCESS_REBOOT_INITIATED - The installer has initiated a restart"
        1642 = "ERROR_PATCH_TARGET_NOT_FOUND - The installer cannot install the upgrade patch because the program to be upgraded may be missing"
        1643 = "ERROR_PATCH_PACKAGE_REJECTED - The update package is not permitted by software restriction policy"
        1644 = "ERROR_INSTALL_TRANSFORM_REJECTED - One or more customizations are not permitted by software restriction policy"
        1645 = "ERROR_INSTALL_REMOTE_PROHIBITED - The Windows Installer does not permit installation from a Remote Desktop Connection"
        1646 = "ERROR_PATCH_REMOVAL_UNSUPPORTED - Uninstallation of the update package is not supported"
        1647 = "ERROR_UNKNOWN_PATCH - The update is not applied to this product"
        1648 = "ERROR_PATCH_NO_SEQUENCE - No valid sequence could be found for the set of updates"
        1649 = "ERROR_PATCH_REMOVAL_DISALLOWED - Update removal was disallowed by policy"
        1650 = "ERROR_INVALID_PATCH_XML - The XML update data is invalid"
        1651 = "ERROR_PATCH_MANAGED_ADVERTISED_PRODUCT - Windows Installer does not permit updating of managed advertised products"
        1652 = "ERROR_INSTALL_SERVICE_SAFEBOOT - The Windows Installer service is not accessible in Safe Mode"
        1653 = "ERROR_FAIL_FAST_EXCEPTION - A fail fast exception occurred"
        1654 = "ERROR_INSTALL_REJECTED - App that you are trying to run is not supported on this version of Windows"
        3010 = "ERROR_SUCCESS_REBOOT_REQUIRED - A restart is required to complete the install"
    }
    # Finalize result
    if (-not $skipInvoke -and $invokeSucceeded)
    {
        if ($isMsiexec -and $null -ne $lastExitCode)
        {
            # Get the exit code description if available
            $exitCodeDescription = if ($msiExitCodes.ContainsKey($lastExitCode))
            {
                $msiExitCodes[$lastExitCode] 
            }
            else
            {
                "Unknown MSI exit code: $lastExitCode" 
            }
            
            Write-Verbose "[$functionName] MSI process completed with exit code: $lastExitCode - $exitCodeDescription"
            write-log -logFile $LogFile -module $functionName -Message "MSI process completed with exit code: $lastExitCode - $exitCodeDescription" -LogLevel "Information"
            
            # For MSI processes, determine success based on common success codes
            # 0 = SUCCESS, 1641 = SUCCESS_REBOOT_INITIATED, 3010 = SUCCESS_REBOOT_REQUIRED
            $msiSuccess = $lastExitCode -eq 0 -or $lastExitCode -eq 1641 -or $lastExitCode -eq 3010
            
            $returnObject.exitCode = $lastExitCode
            $returnObject.exitCodeDescription = $exitCodeDescription
            $returnObject.commandOutput = $proc
            $returnObject.success = $msiSuccess
            $returnObject.errorMessage = if (-not $msiSuccess) { $exitCodeDescription } else { $null }
            
            if ($msiSuccess)
            {
                Write-Verbose "[$functionName] MSI installation completed successfully (exit code indicates success or reboot required)."
                write-log -logFile $LogFile -module $functionName -Message "MSI installation completed successfully (exit code: $lastExitCode)." -LogLevel "Information"
            }
            else
            {
                Write-Error "[$functionName] MSI installation failed with exit code: $lastExitCode - $exitCodeDescription"
                write-log -logFile $LogFile -module $functionName -Message "MSI installation failed with exit code: $lastExitCode - $exitCodeDescription" -LogLevel "Error"
            }
        }
        elseif ($null -eq $lastExitCode -or $lastExitCode -eq 0)
        {
            Write-Verbose "[$functionName] Process Executed Successfully."
            write-log -logFile $LogFile -module $functionName -Message "Process Executed Successfully." -LogLevel "Information"
            $returnObject.exitCode = 0
            $returnObject.commandOutput = $proc
            $returnObject.success = $true
        }
        else
        {
            Write-Error "[$functionName] Process failed with exit code: $lastExitCode"
            write-log -logFile $LogFile -module $functionName -Message "Process failed with exit code: $lastExitCode" -LogLevel "Error"
            $returnObject.exitCode = $lastExitCode
            $returnObject.commandOutput = $proc
            $returnObject.errorMessage = $cmdError
            $returnObject.stoppedProcesses = @()
            $returnObject.success = $false
        }
    }
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Process finished. ExitCode=$($returnObject.exitCode)" -LogLevel "Information"
    Write-Verbose "[$functionName] Process finished. ExitCode=$($returnObject.exitCode)"
    return $returnObject
}
