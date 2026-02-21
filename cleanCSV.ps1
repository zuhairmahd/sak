[CmdletBinding()        ]
param(
    [Parameter(Mandatory = $true)]
    [string]$CSVFileName,
    [string]$outputFile = "Processed_CSVFile.csv",
    [Parameter(Mandatory = $true)]
    [ValidateSet('RemoveColumns', 'extractGroupAssignments')]
    [string]$Operation
)

# Function to parse group assignments with logical operators
function Parse-GroupAssignments()
{
    [CmdletBinding()]
    param(
        [string]$AppliesTo,
        [string]$ExcludesGroups
    )

    $result = @{
        IncludeAND = @()
        IncludeOR  = @()
        ExcludeAND = @()
        ExcludeOR  = @()
    }

    # Parse AppliesTo field for inclusion groups
    if ($AppliesTo -and $AppliesTo -ne 'All Users' -and $AppliesTo -ne 'None')
    {
        # Split by semicolon to get individual group entries
        $groups = $AppliesTo -split ';' | ForEach-Object { $_.Trim() }

        foreach ($group in $groups)
        {
            if ($group)
            {
                # Check if it's an AND group
                if ($group -match '^\[AND\]\s*(.+)$')
                {
                    $groupName = $matches[1].Trim()
                    $result.IncludeAND += $groupName
                }
                # Check if it's an OR group
                elseif ($group -match '^\[OR\]\s*(.+)$')
                {
                    $groupName = $matches[1].Trim()
                    $result.IncludeOR += $groupName
                }
                # No operator means it's a continuation of the previous operator
                else
                {
                    # Determine context based on previous entries
                    if ($result.IncludeAND.Count -gt 0 -and $result.IncludeOR.Count -eq 0)
                    {
                        $result.IncludeAND += $group
                    }
                    else
                    {
                        $result.IncludeOR += $group
                    }
                }
            }
        }
    }

    # Parse ExcludesGroups field for exclusion groups
    if ($ExcludesGroups -and $ExcludesGroups -ne 'None')
    {
        # Split by semicolon to get individual group entries
        $groups = $ExcludesGroups -split ';' | ForEach-Object { $_.Trim() }

        foreach ($group in $groups)
        {
            if ($group)
            {
                # Check if it's an AND NOT group
                if ($group -match '^\[AND\]\s*NOT\s+(.+)$')
                {
                    $groupName = $matches[1].Trim()
                    $result.ExcludeAND += $groupName
                }
                # Check if it's an OR NOT group
                elseif ($group -match '^\[OR\]\s*NOT\s+(.+)$')
                {
                    $groupName = $matches[1].Trim()
                    $result.ExcludeOR += $groupName
                }
                # Check for NOT without explicit operator
                elseif ($group -match '^NOT\s+(.+)$')
                {
                    $groupName = $matches[1].Trim()
                    # Default to AND for exclusions
                    $result.ExcludeAND += $groupName
                }
                # No operator means it's a continuation
                else
                {
                    # Determine context based on previous entries
                    if ($result.ExcludeAND.Count -gt 0 -and $result.ExcludeOR.Count -eq 0)
                    {
                        $result.ExcludeAND += $group
                    }
                    else
                    {
                        $result.ExcludeOR += $group
                    }
                }
            }
        }
    }

    return $result
}

$CSVFile = Join-Path -Path (Get-Location) -ChildPath $CSVFileName
$removedColumns = 0
if (Test-Path -Path $CSVFile)
{
    $csvData = Import-Csv -Path $CSVFile
    Write-Host "Imported $($csvData.Count) rows from $CSVFile"

    if ($Operation -eq 'RemoveColumns'  )
    {
        $columnsToRemove = @(
            'Action',
            'AppliesTo', 'ExcludesGroups',
            'HasInclusionFilter',
            'HasExclusionFilter'
        )
        # Get all column names from the first object
        $allColumns = $csvData[0].PSObject.Properties.Name
        # Determine which columns to keep (exclude the ones in $columnsToRemove)
        $columnsToKeep = $allColumns | Where-Object { $columnsToRemove -notcontains $_ }
        # Count removed columns
        $removedColumns = ($allColumns | Where-Object { $columnsToRemove -contains $_ }).Count
        Write-Host "Removing $removedColumns column(s): $($columnsToRemove -join ', ')"
        # Select only the columns we want to keep
        $csvData = $csvData | Select-Object -Property $columnsToKeep
    }
    elseif ($Operation -eq 'extractGroupAssignments')
    {
        Write-Host "Extracting group assignments with logical operators..."

        # Process each row and extract group assignments
        $results = @()
        foreach ($row in $csvData)
        {
            $groupAssignment = Parse-GroupAssignments -AppliesTo $row.AppliesTo -ExcludesGroups $row.ExcludesGroups
            # Create output object with registry info and group assignments
            $result = [PSCustomObject]@{
                RegistryPath     = $row.RegistryPath
                ValueName        = $row.ValueName
                ValueType        = $row.ValueType
                ValueData        = $row.ValueData
                IncludeGroupsAND = ($groupAssignment.IncludeAND -join '; ')
                IncludeGroupsOR  = ($groupAssignment.IncludeOR -join '; ')
                ExcludeGroupsAND = ($groupAssignment.ExcludeAND -join '; ')
                ExcludeGroupsOR  = ($groupAssignment.ExcludeOR -join '; ')
            }
            $results += $result
        }

        $csvData = $results
        Write-Host "Processed $($results.Count) registry entries with group assignments"
    }
    $csvData | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Columns removed and CSV file updated successfully. Output: $outputFile"
}
else
{
    Write-Host "CSV file not found: $CSVFile"
}
