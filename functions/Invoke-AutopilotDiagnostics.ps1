function Invoke-AutopilotDiagnostics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,
        [Parameter(Mandatory = $true)]
        [string]$fileName,
        [string]$accessToken
    )

    $functionName = $MyInvocation.MyCommand.Name
    $returnObject = @{
        success = $false
        message = $null
    }
    Write-Log -logFile $logFile -Module $functionName -Message "Starting Invoke-AutopilotDiagnostics with RootPath: $RootPath and fileName: $fileName"
    if (-not (Test-Path $fileName)) {
        $returnObject.message = "File not found at path: $fileName"
        Write-Log -logFile $logFile -Module $functionName -Message $returnObject.message -logLevel "ERROR"
        return $returnObject
    }

    $toolPath = Join-Path -Path $RootPath -ChildPath "tools\Get-AutopilotDiagnosticsCommunity.ps1"
    if (Test-Path -Path $toolPath) {
        Write-Log -logFile $logFile -Module $functionName -Message "Found tool at path: $toolPath" -logLevel "INFORMATION"
        $returnObject.message = & $toolPath -File $fileName -Bearer $accessToken -online
        $returnObject.success = $true
        Write-Log -logFile $logFile -Module $functionName -Message "Successfully invoked tool at path: $toolPath" -logLevel "INFORMATION"
    }
    else {
        $returnObject.message = "Tool not found at path: $toolPath"
        Write-Log -logFile $logFile -Module $functionName -Message $returnObject.message -logLevel "ERROR"
        return $returnObject
    }
    return $returnObject
}

