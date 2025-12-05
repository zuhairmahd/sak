function Set-RegKeys()
{
    <#
.SYNOPSIS
Sets, updates, or validates registry key values in HKLM or HKCU.

.DESCRIPTION
Set-RegKeys processes an array of hashtable registry entry definitions and ensures that each
specified key/value exists with the desired data. In normal mode it will create missing keys,
create missing values, and update existing values whose data differs. In -CheckOnly mode it
performs a dry run: no changes are written, but a summary of what would be created or updated
is returned and written to host/log.

Each registry entry has the following expected fields (case-insensitive):
    Path  (string, required)   Registry path under the hive (e.g. 'SOFTWARE\MyApp')
    Name  (string, required)   Value name; use '(Default)' or '@' for the unnamed default value
    Value (any,    optional)   Desired data. If omitted the value may be created empty depending on type
    Type  (string or RegistryValueKind, optional) Defaults to 'String'. Examples: String, DWord, QWord, ExpandString, MultiString, Binary
    Hive  (string, optional)   'LocalMachine' for HKLM; any other value (or omitted) selects HKCU

Default value handling: Passing Name '(Default)' or '@' maps to the empty string value name "".

Return object properties:
    checkOnly             [bool]    Indicates dry run mode
    hasCorrectValues      [bool]    True if all existing values already matched (only meaningful in -CheckOnly)
    missingKeys           [hashtable[]] Entries whose key or value was missing
    missingKeysCount      [int]
    allEntriesProcessed   [bool]    False if any failures occurred
    totalEntries          [int]
    entriesCreated        [hashtable[]]
    entriesCreatedCount   [int]
    entriesUpdated        [hashtable[]]
    entriesUpdatedCount   [int]
    entriesUnchanged      [hashtable[]]
    entriesUnchangedCount [int]
    entriesFailed         [hashtable[]]
    entriesFailedCount    [int]
    message               [string[]] Summary / status messages

Logging: Uses external Write-Log (and $logFile variable) expected to exist in caller scope.

.PARAMETER RegistryEntries
Array of hashtables describing registry modifications. Mandatory.

.PARAMETER CheckOnly
Switch. When supplied, performs a dry run reporting intended changes without applying them.

.OUTPUTS
Hashtable summarizing processing results (see Return object properties above).

.EXAMPLE
$entries = @(
    @{ Path = 'SOFTWARE\\Contoso'; Name = 'InstallPath'; Value = 'C:\\Program Files\\Contoso'; Type = 'String'; Hive = 'LocalMachine' },
    @{ Path = 'SOFTWARE\\Contoso'; Name = 'Enabled'; Value = 1; Type = 'DWord'; Hive = 'CurrentUser' }
)
Set-RegKeys -RegistryEntries $entries

.EXAMPLE
# Dry-run to see what would change
Set-RegKeys -RegistryEntries $entries -CheckOnly | Format-List

.NOTES
    This function directly uses Microsoft.Win32.Registry APIs instead of Set-ItemProperty for finer control
    over default values and type mapping.
    Any failures are captured per entry; successful operations close opened registry keys promptly.

.LINK
None
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable[]]$RegistryEntries,
        [switch]$CheckOnly
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    
    Write-Verbose "[$functionName] Starting to process $($RegistryEntries.Count) registry entries. CheckOnly: $CheckOnly"
    Write-Log -LogFile $logFile -Module $functionName -Message "Starting to process $($RegistryEntries.Count) registry entries. CheckOnly: $CheckOnly"
    
    $returnObject = @{
        checkOnly             = $CheckOnly.IsPresent
        hasCorrectValues      = $true
        missingKeys           = @()
        missingKeysCount      = 0
        allEntriesProcessed   = $false
        totalEntries          = $RegistryEntries.Count
        entriesCreated        = @()
        entriesCreatedCount   = 0
        entriesUpdated        = @()
        entriesUpdatedCount   = 0
        entriesUnchanged      = @()
        entriesUnchangedCount = 0
        entriesFailed         = @()
        entriesFailedCount    = 0
        message               = @()
    }
    
    foreach ($entry in $RegistryEntries)
    {
        # Validate required parameters
        if (-not $entry.Path -or -not $entry.Name)
        {
            $errorMsg = "Invalid entry: Path and Name are required. Entry: $($entry | ConvertTo-Json -Compress)"
            Write-Warning $errorMsg
            Write-Log -LogFile $logFile -Module $functionName -Message $errorMsg -LogLevel "Error"
            $returnObject.entriesFailed += $entry
            $returnObject.entriesFailedCount++
            $returnObject.message += $errorMsg
            continue
        }
        
        # Set defaults
        $path = $entry.Path
        $name = $entry.Name
        $value = $entry.Value
        
        # Handle registry value type - convert string to RegistryValueKind
        try
        {
            $type = if ($entry.Type)
            { 
                if ($entry.Type -is [string])
                {
                    [Microsoft.Win32.RegistryValueKind]::$($entry.Type)
                }
                else
                {
                    $entry.Type
                }
            }
            else
            { 
                [Microsoft.Win32.RegistryValueKind]::String 
            }
        }
        catch
        {
            Write-Warning "[$functionName] Invalid registry type '$($entry.Type)' specified. Defaulting to String."
            Write-Log -LogFile $logFile -Module $functionName -Message "Invalid registry type '$($entry.Type)' specified. Defaulting to String." -LogLevel "Warning"
            $type = [Microsoft.Win32.RegistryValueKind]::String
        }
        
        $hive = if ($entry.Hive -eq "LocalMachine") { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
        $hiveName = if ($entry.Hive -eq "LocalMachine") { "HKLM" } else { "HKCU" }
        
        # Handle default value name (@ in .reg files becomes null or empty in registry API)
        $isDefaultValue = ($name -eq "(Default)" -or $name -eq "@")
        $registryValueName = if ($isDefaultValue) { "" } else { $name }
        $displayName = if ($isDefaultValue) { "(Default)" } else { $name }
        $fullPath = "$hiveName`:$path\$displayName"
        
        Write-Verbose "[$functionName] Entry details - Path: $path, Name: $displayName, Value: $value, Type: $type, Hive: $hiveName"
        Write-Log -LogFile $logFile -Module $functionName -Message "Entry details - Path: $path, Name: $displayName, Value: $value, Type: $type, Hive: $hiveName" -LogLevel "Information"
        Write-Verbose "[$functionName] Processing registry entry: $fullPath"
        Write-Log -LogFile $logFile -Module $functionName -Message "Processing registry entry: $fullPath"
        try
        {
            # Check if the key exists and get current value
            $regKey = $hive.OpenSubKey($path, $false)
            $keyExists = $null -ne $regKey
            $currentValue = $null
            $valueExists = $false
            
            if ($keyExists)
            {
                try
                {
                    $currentValue = $regKey.GetValue($registryValueName)
                    # For default values, check if it's truly set (not empty string default)
                    if ($isDefaultValue)
                    {
                        $valueExists = $null -ne $currentValue -and $currentValue -ne ""
                        Write-Verbose "[$functionName] Default value check - Current: '$currentValue', Exists: $valueExists"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Default value check - Current: '$currentValue', Exists: $valueExists" -LogLevel "Verbose"
                    }
                    else
                    {
                        $valueExists = $null -ne $currentValue
                    }
                    Write-Verbose "[$functionName] Key exists. Current value: $currentValue, Value exists: $valueExists"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Key exists. Current value: $currentValue, Value exists: $valueExists" -LogLevel "Verbose"
                    $regKey.Close()
                }
                catch
                {
                    Write-Verbose "[$functionName] Error reading current value: $_"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Error reading current value: $_" -LogLevel "Warning"
                    $regKey.Close()
                }
            }
            else
            {
                Write-Verbose "[$functionName] Registry key does not exist: $hiveName`:$path"
                Write-Log -LogFile $logFile -Module $functionName -Message "Registry key does not exist: $hiveName`:$path" -LogLevel "Information"
            }
            
            # Determine action needed
            if (-not $keyExists)
            {
                if ($CheckOnly.IsPresent)
                {
                    # CheckOnly mode - report what would be created
                    $returnObject.hasCorrectValues = $false
                    Write-Host "[CHECK ONLY] Would create registry entry: $fullPath = $value (Type: $type)" -ForegroundColor Yellow
                    Write-Log -LogFile $logFile -Module $functionName -Message "[CHECK ONLY] Would create registry entry: $fullPath = $value (Type: $type)" -LogLevel "Information"
                    $returnObject.entriesCreated += $entry
                    $returnObject.entriesCreatedCount++
                    $returnObject.missingKeys += $entry
                    $returnObject.missingKeysCount++
                }
                else
                {
                    # Create the key path
                    Write-Verbose "[$functionName] Creating registry path: $hiveName`:$path"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Creating registry path: $hiveName`:$path" -LogLevel "Information"
                    $regKey = $hive.CreateSubKey($path)
                    Write-Verbose "[$functionName] Setting value: Name='$registryValueName', Value='$value', Type=$type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Setting value: Name='$registryValueName', Value='$value', Type=$type" -LogLevel "Information"
                    $regKey.SetValue($registryValueName, $value, $type)
                    $regKey.Close()
                    
                    Write-Host "Created registry entry: $fullPath = $value (Type: $type)"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Created registry entry: $fullPath = $value (Type: $type)" -LogLevel "Information"
                    $returnObject.entriesCreated += $entry
                    $returnObject.entriesCreatedCount++
                }
            }
            elseif (-not $valueExists)
            {
                if ($CheckOnly.IsPresent)
                {
                    # CheckOnly mode - report what would be created
                    $returnObject.hasCorrectValues = $false
                    Write-Host "[CHECK ONLY] Would create registry value: $fullPath = $value (Type: $type)" -ForegroundColor Yellow
                    Write-Log -LogFile $logFile -Module $functionName -Message "[CHECK ONLY] Would create registry value: $fullPath = $value (Type: $type)" -LogLevel "Information"
                    $returnObject.entriesCreated += $entry
                    $returnObject.entriesCreatedCount++
                    $returnObject.missingKeys += $entry
                    $returnObject.missingKeysCount++
                }
                else
                {
                    # Key exists but value doesn't - create value
                    Write-Verbose "[$functionName] Creating registry value: $fullPath"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Creating registry value: $fullPath" -LogLevel "Information"
                    $regKey = $hive.OpenSubKey($path, $true)
                    Write-Verbose "[$functionName] Setting value: Name='$registryValueName', Value='$value', Type=$type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Setting value: Name='$registryValueName', Value='$value', Type=$type" -LogLevel "Information"
                    $regKey.SetValue($registryValueName, $value, $type)
                    $regKey.Close()
                    Write-Host "Created registry value: $fullPath = $value (Type: $type)"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Created registry value: $fullPath = $value (Type: $type)" -LogLevel "Information"
                    $returnObject.entriesCreated += $entry
                    $returnObject.entriesCreatedCount++
                }
            }
            elseif ($currentValue -ne $value)
            {
                # Special handling for array comparison (MultiString values)
                $valuesAreDifferent = $false
                if ($value -is [array] -and $currentValue -is [array])
                {
                    # Compare arrays element by element
                    if ($value.Count -ne $currentValue.Count)
                    {
                        $valuesAreDifferent = $true
                    }
                    else
                    {
                        for ($idx = 0; $idx -lt $value.Count; $idx++)
                        {
                            if ($value[$idx] -ne $currentValue[$idx])
                            {
                                $valuesAreDifferent = $true
                                break
                            }
                        }
                    }
                }
                elseif ($value -is [byte[]] -and $currentValue -is [byte[]])
                {
                    # Compare byte arrays for Binary values
                    if ($value.Count -ne $currentValue.Count)
                    {
                        $valuesAreDifferent = $true
                    }
                    else
                    {
                        for ($idx = 0; $idx -lt $value.Count; $idx++)
                        {
                            if ($value[$idx] -ne $currentValue[$idx])
                            {
                                $valuesAreDifferent = $true
                                break
                            }
                        }
                    }
                }
                else
                {
                    # Simple comparison for scalars
                    $valuesAreDifferent = ($currentValue -ne $value)
                }
                
                if ($valuesAreDifferent)
                {
                    if ($CheckOnly.IsPresent)
                    {
                        # CheckOnly mode - report what would be updated
                        $returnObject.hasCorrectValues = $false
                        Write-Host "[CHECK ONLY] Would update registry value: $fullPath = $value (currently: $currentValue) (Type: $type)" -ForegroundColor Yellow
                        Write-Log -LogFile $logFile -Module $functionName -Message "[CHECK ONLY] Would update registry value: $fullPath = $value (currently: $currentValue) (Type: $type)" -LogLevel "Information"
                        $returnObject.entriesUpdated += $entry
                        $returnObject.entriesUpdatedCount++
                    }
                    else
                    {
                        # Value exists but is different - update it
                        Write-Verbose "[$functionName] Updating registry value: $fullPath (Old: $currentValue, New: $value)"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Updating registry value: $fullPath (Old: $currentValue, New: $value)" -LogLevel "Information"
                        $regKey = $hive.OpenSubKey($path, $true)
                        Write-Verbose "[$functionName] Setting value: Name='$registryValueName', Value='$value', Type=$type"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Setting value: Name='$registryValueName', Value='$value', Type=$type" -LogLevel "Information"
                        $regKey.SetValue($registryValueName, $value, $type)
                        $regKey.Close()
                        
                        Write-Host "Updated registry value: $fullPath = $value (was: $currentValue) (Type: $type)"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Updated registry value: $fullPath = $value (was: $currentValue) (Type: $type)" -LogLevel "Information"
                        $returnObject.entriesUpdated += $entry
                        $returnObject.entriesUpdatedCount++
                    }
                }
            }
            else
            {
                # Value already correct
                Write-Verbose "[$functionName] Registry value unchanged: $fullPath = $value"
                Write-Log -LogFile $logFile -Module $functionName -Message "Registry value unchanged: $fullPath = $value" -LogLevel "Information"
                Write-Host "Registry value already correct: $fullPath = $value"
                $returnObject.entriesUnchanged += $entry
                $returnObject.entriesUnchangedCount++
            }
        }
        catch
        {
            $errorMsg = "Failed to process registry entry: $fullPath. Error: $($_.Exception.Message)"
            $errorDetails = "Error details - Path: $path, Name: $displayName, Value: $value, Type: $type, Exception: $($_.Exception.GetType().FullName), StackTrace: $($_.ScriptStackTrace)"
            Write-Warning $errorMsg
            Write-Log -LogFile $logFile -Module $functionName -Message $errorMsg -LogLevel "Error"
            Write-Log -LogFile $logFile -Module $functionName -Message $errorDetails -LogLevel "Error"
            $returnObject.entriesFailed += $entry
            $returnObject.entriesFailedCount++
            $returnObject.message += $errorMsg
        }
    }
    
    # Determine overall success
    $checkOnlyPrefix = if ($CheckOnly.IsPresent) { "[CHECK ONLY] " } else { "" }
    $summaryMsg = "${checkOnlyPrefix}Registry processing summary: Total=$($returnObject.totalEntries), Created=$($returnObject.entriesCreatedCount), Updated=$($returnObject.entriesUpdatedCount), Unchanged=$($returnObject.entriesUnchangedCount), Failed=$($returnObject.entriesFailedCount)"
    Write-Verbose "[$functionName] $summaryMsg"
    Write-Log -LogFile $logFile -Module $functionName -Message $summaryMsg -LogLevel "Information"
    
    if ($returnObject.entriesFailedCount -eq 0)
    {
        $returnObject.allEntriesProcessed = $true
        if ($CheckOnly.IsPresent)
        {
            if ($returnObject.hasCorrectValues)
            {
                $successMsg = "[CHECK ONLY] All $($returnObject.totalEntries) registry entries already have correct values."
                Write-Host $successMsg -ForegroundColor Green
            }
            else
            {
                $changesNeeded = $returnObject.entriesCreatedCount + $returnObject.entriesUpdatedCount
                $successMsg = "[CHECK ONLY] $changesNeeded of $($returnObject.totalEntries) registry entries need changes."
                Write-Host $successMsg -ForegroundColor Yellow
            }
        }
        else
        {
            $successMsg = "All $($returnObject.totalEntries) registry entries processed successfully."
            Write-Host $successMsg -ForegroundColor Green
        }
        $returnObject.message += $successMsg
        Write-Log -LogFile $logFile -Module $functionName -Message $successMsg -LogLevel "Information"
    }
    else
    {
        $failureMsg = "${checkOnlyPrefix}Processed $($returnObject.totalEntries) registry entries with $($returnObject.entriesFailedCount) failures."
        $returnObject.message += $failureMsg
        Write-Host $failureMsg -ForegroundColor Yellow
        Write-Log -LogFile $logFile -Module $functionName -Message $failureMsg -LogLevel "Warning"
    }
    
    return $returnObject
}