function Invoke-ExecutableWithUserEnvironmentVariables {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $false)]
        [string]$Arguments = "",
        [Parameter(Mandatory = $false)]
        [hashtable]$EnvironmentOverrides = @{},
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300
    )

    $functionName = $MyInvocation.MyCommand.Name
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    # Apply user environment spoofing
    foreach ($key in $EnvironmentOverrides.Keys) {
        $psi.EnvironmentVariables[$key] = $EnvironmentOverrides[$key]
    }

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    # Use StringBuilders with thread-safe event handlers to avoid buffer blocking
    $stdOutBuilder = New-Object System.Text.StringBuilder
    $stdErrBuilder = New-Object System.Text.StringBuilder

    $outEvent = Register-ObjectEvent -InputObject $proc -EventName "OutputDataReceived" -Action {
        if (-not [string]::IsNullOrEmpty($EventArgs.Data)) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stdOutBuilder

    $errEvent = Register-ObjectEvent -InputObject $proc -EventName "ErrorDataReceived" -Action {
        if (-not [string]::IsNullOrEmpty($EventArgs.Data)) {
            [void]$Event.MessageData.AppendLine($EventArgs.Data)
        }
    } -MessageData $stdErrBuilder

    try {
        $null = $proc.Start()

        # Begin asynchronous reads on background threads
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        # Wait with an explicit millisecond timeout (prevents infinite lock)
        $hasExited = $proc.WaitForExit($TimeoutSeconds * 1000)

        if (-not $hasExited) {
            $proc.Kill()
            Write-Log -Message "Process timed out after $TimeoutSeconds seconds and was terminated." -logFile $logFile -module $functionName -LogLevel "Warning"
            throw "Execution timed out."
        }

        # Small sleep to ensure async buffers flush
        Start-Sleep -Milliseconds 250

        return [PSCustomObject]@{
            ExitCode = $proc.ExitCode
            StdOut   = $stdOutBuilder.ToString().Trim()
            StdErr   = $stdErrBuilder.ToString().Trim()
        }
    }
    finally {
        # Clean up registered event subscribers
        Unregister-Event -SourceIdentifier $outEvent.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $errEvent.Name -ErrorAction SilentlyContinue
        $proc.Dispose()
    }
}
