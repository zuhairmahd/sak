[cmdletBinding()]
param(
    [string]$groupTag = 'MSB01',
    [string]$CSVPath = $pwd,
    [string]$outputPath = (Join-Path $PSScriptRoot 'output_hardware_hashes.csv'),
    [switch]$sanetize
)

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

$outputFileName = Split-Path $outputPath -Leaf
$outputSanetizedFileName = [System.IO.Path]::ChangeExtension($outputFileName, 'sanetized.csv')
$csvFiles = Get-ChildItem -Path $CSVPath -Filter *.csv |
Where-Object { $_.Name -ne $outputFileName -and $_.Name -ne $outputSanetizedFileName }

Write-Host "Found $($csvFiles.Count) CSV files in $($CSVPath)." -ForegroundColor Green
foreach ($file in $csvFiles) {
    $rows = Import-Csv -Path $file.FullName
    if ($rows.Count -eq 0) {
        continue
    }

    $columns = $rows[0].PSObject.Properties.Name
    $isStandardFormat = $columns -contains 'Device Serial Number'
    $isAlternateFormat = $columns -contains 'ServiceTag'

    if (-not $isStandardFormat -and -not $isAlternateFormat) {
        Write-Warning "Unrecognized format in file: $($file.Name)... skipping."
        continue
    }

    foreach ($row in $rows) {
        if ($isStandardFormat) {
            $serial = $row.'Device Serial Number'
            $productId = $row.'Windows Product ID'
            $hash = $row.'Hardware Hash'
        }
        else {
            $serial = $row.ServiceTag
            $productId = $row.WINDOWSPRODUCTID
            $hash = $row.SysdatarefMax01
        }

        if ([string]::IsNullOrWhiteSpace($serial)) {
            continue
        }

        $results.Add([PSCustomObject][ordered]@{
                'Device Serial Number' = $serial
                'Windows Product ID'   = $productId
                'Hardware Hash'        = $hash
                'Group Tag'            = $groupTag
            })
    }
}

Write-Host "Total devices processed: $($results.Count)" -ForegroundColor Green

$unique = $results | Sort-Object 'Device Serial Number' -Unique
Write-Host "Unique devices after deduplication: $($unique.Count)" -ForegroundColor Green
#If sanetize is true, put only the serial number.
if ($sanetize) {
    Write-host "Sanetizing output to include only serial numbers." -ForegroundColor Yellow
    foreach ($item in $unique) {
        $item.PSObject.Properties.Remove('Windows Product ID')
        $item.PSObject.Properties.Remove('Hardware Hash')
        $item.PSObject.Properties.Remove('Group Tag')
    }                               
    $outputPath = [System.IO.Path]::ChangeExtension($outputPath, 'sanetized.csv')
    Write-Host "Output path changed to: $outputPath" -ForegroundColor Yellow
}               


$unique | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "Exported to: $outputPath" -ForegroundColor Green