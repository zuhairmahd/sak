[CmdletBinding()]
param(
    [string]$inputFile,
    [string]$outputDirectory,
    [int]$splitValue = 500
)

[string]$resolvedInput = $null

if (-not $inputFile)
{
    throw "An input CSV file path is required."
}

try
{
    $resolvedInput = (Resolve-Path -Path $inputFile).Path
}
catch
{
    throw "Input file '$inputFile' was not found."
}

if (-not (Test-Path -Path $resolvedInput -PathType Leaf))
{
    throw "Input file '$inputFile' was not found."
}

if (-not $outputDirectory)
{
    $outputDirectory = Split-Path -Path $resolvedInput -Parent
}

if (-not (Test-Path -Path $outputDirectory))
{
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

if ($splitValue -lt 1)
{
    throw "splitValue must be greater than zero."
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)
$extension = [System.IO.Path]::GetExtension($resolvedInput)

# Use streaming approach with Import-Csv to handle large files efficiently
$chunk = New-Object System.Collections.Generic.List[object]
$partNumber = 1
$totalRowsProcessed = 0

Import-Csv -Path $resolvedInput | ForEach-Object {
    $chunk.Add($_)
    $totalRowsProcessed++

    if ($chunk.Count -ge $splitValue)
    {
        $outputFile = Join-Path $outputDirectory ("{0}_part{1:000}{2}" -f $baseName, $partNumber, $extension)
        $chunk | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        Write-Host "Created: $outputFile ($($chunk.Count) row(s) plus header)"
        $chunk.Clear()
        $partNumber++
    }
}

# Write any remaining rows in the final chunk
if ($chunk.Count -gt 0)
{
    $outputFile = Join-Path $outputDirectory ("{0}_part{1:000}{2}" -f $baseName, $partNumber, $extension)
    $chunk | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Created: $outputFile ($($chunk.Count) row(s) plus header)"
}

if ($totalRowsProcessed -eq 0)
{
    Write-Warning "Input CSV is empty or contains only a header. No files were created."
}

