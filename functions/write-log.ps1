function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][string]$LogFile,
        [Parameter(Mandatory)][string]$Module,
        [ValidateSet('Information', 'Warning', 'Error', 'Debug', 'Verbose')][string]$LogLevel = 'Information',
        [int]$MaxLogSizeMB = 10,
        [bool]$NoEcho = $true
    )
    if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt ($MaxLogSizeMB * 1MB)) {
        Rename-Item -Path $LogFile -NewName ([IO.Path]::ChangeExtension($LogFile, 'bak')) -Force -ErrorAction SilentlyContinue
    }
    $logDir = Split-Path $LogFile -Parent
    if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
    $severity = switch ($LogLevel) { 'Error' { 3 } 'Warning' { 2 } default { 1 } }
    $entry = "<![LOG[$Message]LOG]!><time=`"$(Get-Date -Format 'HH:mm:ss.fff')`" date=`"$(Get-Date -Format 'MM-dd-yyyy')`" component=`"$Module`" context=`"`" type=`"$severity`" thread=`"$PID`" file=`"`">"
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Verbose "[$Module] $Message"
    switch ($LogLevel) {
        'Error' { if (-not $NoEcho) { Write-Host $Message -ForegroundColor Red } }
        'Warning' { if (-not $NoEcho) { Write-Host $Message -ForegroundColor Yellow } }
        'Verbose' { if (-not $NoEcho) { Write-Verbose $Message } }
        default { if (-not $NoEcho) { Write-Host $Message } }
    }
}
