# Parse Group Policy Report XML and Extract Registry Settings
# This script reads the gpreport.xml file and extracts registry settings with their associated groups

param(
    [string]$XmlPath = ".\gpreport.xml",
    [switch]$MatchGroups
)

# Load the XML file
try
{
    [xml]$gpoXml = Get-Content -Path $XmlPath -ErrorAction Stop
    Write-Host "Successfully loaded XML file: $XmlPath" -ForegroundColor Green
}
catch
{
    Write-Host "Error loading XML file: $_" -ForegroundColor Red
    exit 1
}

# Define namespaces
$nsmgr = New-Object System.Xml.XmlNamespaceManager($gpoXml.NameTable)
$nsmgr.AddNamespace("q3", "http://www.microsoft.com/GroupPolicy/Settings/Windows/Registry")

# Get all Registry nodes
$registryNodes = $gpoXml.SelectNodes("//q3:Registry", $nsmgr)

Write-Host "Found $($registryNodes.Count) registry settings" -ForegroundColor Green

# Create an array to store the results
$registrySettings = @()

foreach ($regNode in $registryNodes)
{
    $properties = $regNode.Properties
    $filters = $regNode.Filters
    # Build registry path
    $hive = $properties.hive
    $key = $properties.key
    $name = $properties.name
    $type = $properties.type
    $value = $properties.value
    $action = $properties.action

    # Construct full registry path
    $fullPath = "$hive\$key"

    # Get associated groups
    $includedGroups = @()
    $excludedGroups = @()
    $allGroups = @()

    if ($filters -and $filters.FilterGroup)
    {
        foreach ($filterGroup in $filters.FilterGroup)
        {
            $groupName = $filterGroup.name
            $sid = $filterGroup.sid
            $boolOperator = $filterGroup.bool
            $isExcluded = ($filterGroup.not -eq "1")
            $notOperator = if ($isExcluded)
            {
                "NOT"
            }
            else
            {
                ""
            }

            $groupObj = [PSCustomObject]@{
                GroupName    = $groupName
                SID          = $sid
                BoolOperator = $boolOperator
                NotOperator  = $notOperator
                IsExcluded   = $isExcluded
            }

            $allGroups += $groupObj

            if ($isExcluded)
            {
                $excludedGroups += $groupObj
            }
            else
            {
                $includedGroups += $groupObj
            }
        }
    }

    # Create registry setting object
    $setting = [PSCustomObject]@{
        RegistryPath   = $fullPath
        ValueName      = $name
        ValueType      = $type
        ValueData      = $value
        Action         = $action
        Groups         = $allGroups
        IncludedGroups = $includedGroups
        ExcludedGroups = $excludedGroups
    }

    $registrySettings += $setting
}

# Display results
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Registry Settings from GPO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

foreach ($setting in $registrySettings)
{
    Write-Host "Registry Path: " -NoNewline -ForegroundColor Yellow
    Write-Host $setting.RegistryPath
    Write-Host "  Value Name: " -NoNewline -ForegroundColor Gray
    Write-Host $setting.ValueName
    Write-Host "  Value Type: " -NoNewline -ForegroundColor Gray
    Write-Host $setting.ValueType
    Write-Host "  Value Data: " -NoNewline -ForegroundColor Gray
    Write-Host $setting.ValueData
    Write-Host "  Action: " -NoNewline -ForegroundColor Gray
    $actionText = switch ($setting.Action)
    {
        "U"
        {
            "Update"
        }
        "D"
        {
            "Delete"
        }
        "C"
        {
            "Create"
        }
        "R"
        {
            "Replace"
        }
        default
        {
            $setting.Action
        }
    }
    Write-Host $actionText

    if ($setting.Groups.Count -gt 0)
    {
        Write-Host "  Applied to Groups:" -ForegroundColor Magenta
        foreach ($group in $setting.Groups)
        {
            $notText = if ($group.NotOperator)
            {
                "NOT "
            }
            else
            {
                ""
            }
            Write-Host "    - [$($group.BoolOperator)] $notText$($group.GroupName) (SID: $($group.SID))" -ForegroundColor DarkCyan
        }
    }
    else
    {
        Write-Host "  Applied to: All Users" -ForegroundColor Magenta
    }

    Write-Host ""
}

# Export to CSV for easy analysis
if ($MatchGroups)
{
    # Enhanced mode: separate included and excluded groups
    $exportPath = Join-Path (Split-Path $XmlPath -Parent) "GPO-Registry-Settings-GroupMatching.csv"
    $csvData = @()
    foreach ($setting in $registrySettings)
    {
        $includedGroupNames = ($setting.IncludedGroups | ForEach-Object {
                "[$($_.BoolOperator)] $($_.GroupName)"
            }) -join "; "

        $excludedGroupNames = ($setting.ExcludedGroups | ForEach-Object {
                "[$($_.BoolOperator)] NOT $($_.GroupName)"
            }) -join "; "

        $csvData += [PSCustomObject]@{
            RegistryPath       = $setting.RegistryPath
            ValueName          = $setting.ValueName
            ValueType          = $setting.ValueType
            ValueData          = $setting.ValueData
            Action             = switch ($setting.Action)
            {
                "U"
                {
                    "Update"
                }
                "D"
                {
                    "Delete"
                }
                "C"
                {
                    "Create"
                }
                "R"
                {
                    "Replace"
                }
                default
                {
                    $setting.Action
                }
            }
            AppliesTo          = if ($includedGroupNames)
            {
                $includedGroupNames
            }
            else
            {
                "All Users"
            }
            ExcludesGroups     = if ($excludedGroupNames)
            {
                $excludedGroupNames
            }
            else
            {
                "None"
            }
            HasInclusionFilter = $setting.IncludedGroups.Count -gt 0
            HasExclusionFilter = $setting.ExcludedGroups.Count -gt 0
        }
    }

    $csvData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported group-matching registry settings to: $exportPath" -ForegroundColor Green

    # Additional summary for MatchGroups mode
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Group Matching Analysis" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Settings with Inclusion Filters: $(($registrySettings | Where-Object { $_.IncludedGroups.Count -gt 0 }).Count)" -ForegroundColor White
    Write-Host "Settings with Exclusion Filters: $(($registrySettings | Where-Object { $_.ExcludedGroups.Count -gt 0 }).Count)" -ForegroundColor White
    Write-Host "Settings with Both Types: $(($registrySettings | Where-Object { $_.IncludedGroups.Count -gt 0 -and $_.ExcludedGroups.Count -gt 0 }).Count)" -ForegroundColor White
}
else
{
    # Original mode: combined groups display
    $exportPath = Join-Path (Split-Path $XmlPath -Parent) "GPO-Registry-Settings.csv"
    $csvData = @()
    foreach ($setting in $registrySettings)
    {
        $groupNames = ($setting.Groups | ForEach-Object {
                $notText = if ($_.NotOperator)
                {
                    "NOT "
                }
                else
                {
                    ""
                }
                "[$($_.BoolOperator)] $notText$($_.GroupName)"
            }) -join "; "

        $csvData += [PSCustomObject]@{
            RegistryPath    = $setting.RegistryPath
            ValueName       = $setting.ValueName
            ValueType       = $setting.ValueType
            ValueData       = $setting.ValueData
            Action          = switch ($setting.Action)
            {
                "U"
                {
                    "Update"
                }
                "D"
                {
                    "Delete"
                }
                "C"
                {
                    "Create"
                }
                "R"
                {
                    "Replace"
                }
                default
                {
                    $setting.Action
                }
            }
            AppliedToGroups = if ($groupNames)
            {
                $groupNames
            }
            else
            {
                "All Users"
            }
        }
    }

    $csvData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported registry settings to: $exportPath" -ForegroundColor Green
}

# Summary statistics
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Registry Settings: $($registrySettings.Count)" -ForegroundColor White
Write-Host "Settings with Group Filters: $(($registrySettings | Where-Object { $_.Groups.Count -gt 0 }).Count)" -ForegroundColor White
Write-Host "Settings Applied to All: $(($registrySettings | Where-Object { $_.Groups.Count -eq 0 }).Count)" -ForegroundColor White

# Group by registry key for easier understanding
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Settings Grouped by Registry Key" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$groupedByKey = $registrySettings | Group-Object -Property { $_.RegistryPath -replace '\\[^\\]+$', '' }
foreach ($keyGroup in $groupedByKey | Sort-Object Name)
{
    Write-Host "$($keyGroup.Name) - $($keyGroup.Count) value(s)" -ForegroundColor Yellow
}
