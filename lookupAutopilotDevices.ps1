[cmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$serialNumber,
    [string]$inputFile = (Join-Path $PSScriptRoot 'output_hardware_hashes.csv'),
    [string]$outputPath = (Join-Path $PSScriptRoot 'exported_hardware_hashes.csv')
)

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
if (Test-Path -Path $inputFile) {
    $rows = Import-Csv -Path $inputFile
    if ($rows.Count -eq 0) {
        Write-Error "Input file is empty: $inputFile"
        exit 1
    }
}
else {
    Write-Error "Input file not found: $inputFile"
    exit 1
}

$suppliedSerialNumber = $serialNumber.trim()
if ($suppliedSerialNumber.StartsWith('w11-')) {
    Write-Host "Serial number starts with 'w11-', removing prefix for lookup." -ForegroundColor Yellow
    $suppliedSerialNumber = $suppliedSerialNumber.Substring(4)
}
$possibleDevices = @($rows | Where-Object { $_.'Device Serial Number' -eq $suppliedSerialNumber })
Write-Host "Devices found with serial number '$suppliedSerialNumber': $($possibleDevices.Count)" -ForegroundColor Green
foreach ($possibleDevice in $possibleDevices) {
    $results.Add([PSCustomObject][ordered]@{
            'Device Serial Number' = $possibleDevice.'Device Serial Number'
            'Windows Product ID'   = $possibleDevice.'Windows Product ID'
            'Hardware Hash'        = $possibleDevice.'Hardware Hash'
            'Group Tag'            = if ($possibleDevice.'Group Tag') { $possibleDevice.'Group Tag' } else { 'MSB01' }
        })
}

Write-Host "Total devices processed: $($results.Count)" -ForegroundColor Green

if ($results.Count -eq 0) {
    Write-Host "No devices found with the provided serial number: $suppliedSerialNumber" -ForegroundColor Yellow
    exit 0
}
$choice = Read-Host -Prompt "Press E to export to a file, press to display device information,`n press any other key to continue"
if ($choice -notin @('E', 'e', 'D', 'd')) {
    Write-Host "Exiting without exporting." -ForegroundColor Yellow
    exit 0
}
elseif ($choice -in @('D', 'd')) {
    $results | Format-Table -AutoSize
    exit 0
}
else {
    $results | Export-Csv -Path $outputPath -NoTypeInformation -Force
    Write-Host "Exported to: $outputPath" -ForegroundColor Green
}
