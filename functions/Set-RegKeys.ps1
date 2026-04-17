function Set-RegKeys()
{
    <#
.SYNOPSIS
Sets, updates, or validates registry key values in HKLM or HKCU.

.DESCRIPTION
Set-RegKeys processes an array of registry entry definitions (hashtables or PSCustomObjects)
and ensures that each specified key/value exists with the desired data. In normal mode it will
create missing keys, create missing values, and update existing values whose data differs.
In -CheckOnly mode it performs a dry run: no changes are written, but a summary of what would
be created or updated is returned and written to host/log.

Each registry entry has the following expected fields (case-insensitive):
    Path  (string, required)   Registry path under the hive (e.g. 'SOFTWARE\MyApp')
    Name  (string, required)   Value name; use '(Default)' or '@' for the unnamed default value
    Value (any,    optional)   Desired data. If omitted the value may be created empty depending on type
    Type  (string or RegistryValueKind, optional) Defaults to 'String'. Examples: String, DWord, QWord, ExpandString, MultiString, Binary
    Hive  (string, optional)   'LocalMachine' for HKLM; any other value (or omitted) selects HKCU

Default value handling: Passing Name '(Default)' or '@' maps to the empty string value name "".

Return object properties:
    success               [bool]    True if all entries were processed without failures (entriesFailedCount -eq 0)
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
Array of hashtables or PSCustomObjects describing registry modifications. Accepts output from
Import-CSVRegKeys or manually constructed hashtable arrays. Mandatory.

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

.EXAMPLE
# Import registry keys from CSV and apply them
$importResult = Import-CSVRegKeys -FilePath 'C:\config\RegKeys.csv'
if ($importResult.success) {
    Set-RegKeys -RegistryEntries $importResult.regKeys
}

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
        [object[]]$RegistryEntries,
        [switch]$CheckOnly
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting to process $($RegistryEntries.Count) registry entries. CheckOnly: $CheckOnly"
    Write-Log -LogFile $logFile -Module $functionName -Message "Starting to process $($RegistryEntries.Count) registry entries. CheckOnly: $CheckOnly"

    $returnObject = @{
        success               = $false
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
                    # Attempt to convert string to RegistryValueKind
                    $typeValue = [Microsoft.Win32.RegistryValueKind]::$($entry.Type)
                    $typeValue
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
            $validTypes = [System.Enum]::GetNames([Microsoft.Win32.RegistryValueKind]) -join ', '
            Write-Warning "[$functionName] Invalid registry type '$($entry.Type)' specified for entry at path '$($entry.Path)'. Valid types are: $validTypes. Defaulting to String."
            Write-Log -LogFile $logFile -Module $functionName -Message "Invalid registry type '$($entry.Type)' specified for entry at path '$($entry.Path)'. Valid types are: $validTypes. Defaulting to String." -LogLevel "Warning"
            $type = [Microsoft.Win32.RegistryValueKind]::String
        }

        # Convert hexadecimal strings to integers for DWord and QWord types
        if (($type -eq [Microsoft.Win32.RegistryValueKind]::DWord -or $type -eq [Microsoft.Win32.RegistryValueKind]::QWord) -and $value -is [string])
        {
            try
            {
                # Check if the string looks like a hexadecimal value (all chars are 0-9, A-F)
                if ($value -match '^[0-9A-Fa-f]+$')
                {
                    $originalValue = $value
                    # Convert hex string to integer
                    if ($type -eq [Microsoft.Win32.RegistryValueKind]::DWord)
                    {
                        # DWord values are unsigned 32-bit (0 to 0xFFFFFFFF)
                        $value = [Convert]::ToUInt32($value, 16)
                    }
                    else
                    {
                        # QWord values are unsigned 64-bit (0 to 0xFFFFFFFFFFFFFFFF)
                        $value = [Convert]::ToUInt64($value, 16)
                    }
                    Write-Verbose "[$functionName] Converted hexadecimal string '$originalValue' to integer $value for $type type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Converted hexadecimal string '$originalValue' to integer $value for $type type" -LogLevel "Verbose"
                }
                elseif ($value -match '^\d+$')
                {
                    # It's a decimal string, convert it to integer
                    $originalValue = $value
                    if ($type -eq [Microsoft.Win32.RegistryValueKind]::DWord)
                    {
                        # DWord values are unsigned 32-bit (0 to 4294967295)
                        $value = [Convert]::ToUInt32($value, 10)
                    }
                    else
                    {
                        # QWord values are unsigned 64-bit
                        $value = [Convert]::ToUInt64($value, 10)
                    }
                    Write-Verbose "[$functionName] Converted decimal string '$originalValue' to integer $value for $type type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Converted decimal string '$originalValue' to integer $value for $type type" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-Warning "[$functionName] Failed to convert value '$value' to integer for $type type: $_"
                Write-Log -LogFile $logFile -Module $functionName -Message "Failed to convert value '$value' to integer for $type type: $_" -LogLevel "Warning"
            }
        }

        # Convert strings/integers to byte arrays for Binary type
        if ($type -eq [Microsoft.Win32.RegistryValueKind]::Binary -and $value -isnot [byte[]])
        {
            try
            {
                $originalValue = $value
                if ($value -is [string])
                {
                    # Check if it's a hex string (even number of hex digits, optionally with spaces or commas)
                    $cleanHex = $value -replace '[\s,]', ''
                    if ($cleanHex -match '^([0-9A-Fa-f]{2})+$')
                    {
                        # Convert hex string to byte array (e.g., "010203" -> [byte[]]@(1,2,3))
                        $bytes = for ($i = 0; $i -lt $cleanHex.Length; $i += 2)
                        {
                            [Convert]::ToByte($cleanHex.Substring($i, 2), 16)
                        }
                        $value = [byte[]]$bytes
                        Write-Verbose "[$functionName] Converted hex string '$originalValue' to byte array (length: $($value.Length)) for Binary type"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Converted hex string '$originalValue' to byte array (length: $($value.Length)) for Binary type" -LogLevel "Verbose"
                    }
                    elseif ($cleanHex -match '^\d+$')
                    {
                        # It's a decimal string, convert to single byte or multi-byte based on value
                        $intValue = [int]$cleanHex
                        if ($intValue -le 255)
                        {
                            $value = [byte[]]@([byte]$intValue)
                        }
                        else
                        {
                            # Convert integer to byte array (little-endian)
                            $value = [System.BitConverter]::GetBytes($intValue)
                        }
                        Write-Verbose "[$functionName] Converted decimal string '$originalValue' to byte array (length: $($value.Length)) for Binary type"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Converted decimal string '$originalValue' to byte array (length: $($value.Length)) for Binary type" -LogLevel "Verbose"
                    }
                    else
                    {
                        # Treat as UTF8 string and convert to bytes
                        $value = [System.Text.Encoding]::UTF8.GetBytes($value)
                        Write-Verbose "[$functionName] Converted string '$originalValue' to UTF8 byte array (length: $($value.Length)) for Binary type"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Converted string '$originalValue' to UTF8 byte array (length: $($value.Length)) for Binary type" -LogLevel "Verbose"
                    }
                }
                elseif ($value -is [int] -or $value -is [long])
                {
                    # Convert integer to byte array
                    if ($value -le 255 -and $value -ge 0)
                    {
                        $value = [byte[]]@([byte]$value)
                    }
                    else
                    {
                        $value = [System.BitConverter]::GetBytes($value)
                    }
                    Write-Verbose "[$functionName] Converted integer '$originalValue' to byte array (length: $($value.Length)) for Binary type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Converted integer '$originalValue' to byte array (length: $($value.Length)) for Binary type" -LogLevel "Verbose"
                }
                else
                {
                    # Attempt generic conversion
                    $value = [byte[]]$value
                    Write-Verbose "[$functionName] Converted value '$originalValue' to byte array (length: $($value.Length)) for Binary type"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Converted value '$originalValue' to byte array (length: $($value.Length)) for Binary type" -LogLevel "Verbose"
                }
            }
            catch
            {
                Write-Warning "[$functionName] Failed to convert value '$originalValue' to byte array for Binary type: $_"
                Write-Log -LogFile $logFile -Module $functionName -Message "Failed to convert value '$originalValue' to byte array for Binary type: $_" -LogLevel "Warning"
            }
        }

        $hive = if ($entry.Hive -eq "LocalMachine")
        {
            [Microsoft.Win32.Registry]::LocalMachine
        }
        else
        {
            [Microsoft.Win32.Registry]::CurrentUser
        }
        $hiveName = if ($entry.Hive -eq "LocalMachine")
        {
            "HKLM"
        }
        else
        {
            "HKCU"
        }

        # Handle default value name (@ in .reg files becomes null or empty in registry API)
        $isDefaultValue = ($name -eq "(Default)" -or $name -eq "@")
        $registryValueName = if ($isDefaultValue)
        {
            ""
        }
        else
        {
            $name
        }
        $displayName = if ($isDefaultValue)
        {
            "(Default)"
        }
        else
        {
            $name
        }
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

                    Write-Verbose "[$functionName] Created registry entry: $fullPath = $value (Type: $type)"
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
                    Write-Verbose "[$functionName] Created registry value: $fullPath = $value (Type: $type)"
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

                        Write-Verbose "[$functionName] Updated registry value: $fullPath = $value (was: $currentValue) (Type: $type)"
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
                Write-Verbose "[$functionName] No action needed for registry entry: $fullPath = $value"
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
    $checkOnlyPrefix = if ($CheckOnly.IsPresent)
    {
        "[CHECK ONLY] "
    }
    else
    {
        ""
    }
    $summaryMsg = "${checkOnlyPrefix}Registry processing summary: Total=$($returnObject.totalEntries), Created=$($returnObject.entriesCreatedCount), Updated=$($returnObject.entriesUpdatedCount), Unchanged=$($returnObject.entriesUnchangedCount), Failed=$($returnObject.entriesFailedCount)"
    Write-Verbose "[$functionName] $summaryMsg"
    Write-Log -LogFile $logFile -Module $functionName -Message $summaryMsg -LogLevel "Information"

    if ($returnObject.entriesFailedCount -eq 0)
    {
        $returnObject.success = $true
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