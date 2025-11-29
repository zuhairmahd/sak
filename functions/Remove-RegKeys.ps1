function Remove-RegKeys()
{
    <#
.SYNOPSIS
    Removes one or more registry values based on an array of registry configurations.

.DESCRIPTION
    The Remove-RegKeys function processes an array of hashtables, each containing registry path 
    and key name. It checks if the registry path and value exist, and removes them if found.
    The function returns detailed information about which registry entries were removed, already 
    removed, or failed to process. Supports -WhatIf and -Confirm for safe operations.

.PARAMETER RegistryEntries
    An array of hashtables containing registry entries to remove. Each hashtable should have:
    - Path: The registry path (e.g., "Software\MyApp\Settings")
    - Name: The registry key name (use "(Default)" or "@" for default value)
    - Hive: (Optional) The registry hive (default: CurrentUser). Options: CurrentUser, LocalMachine

.EXAMPLE
    $registryEntries = @(
        @{Path = "Software\MyApp\Settings"; Name = "EnableFeature"},
        @{Path = "Software\MyApp\Config"; Name = "ApiUrl"}
    )
    Remove-RegKeys -RegistryEntries $registryEntries
    This example removes the specified registry values.

.EXAMPLE
    $registryEntries = @(
        @{Path = "Software\MyApp\Settings"; Name = "(Default)"; Hive = "LocalMachine"}
    )
    Remove-RegKeys -RegistryEntries $registryEntries -WhatIf
    This example shows what would be removed without actually removing it.

.OUTPUTS
    Returns a hashtable containing:
    - allEntriesProcessed: Boolean indicating if all entries were successfully processed
    - allRemoved: Boolean indicating if all entries were successfully removed
    - removedKeys: Array of entries that were removed
    - removedKeysCount: Count of removed entries
    - entriesUnchanged: Array of entries that were already removed or didn't exist
    - entriesUnchangedCount: Count of entries that were already removed
    - entriesFailed: Array of entries that failed to process
    - entriesFailedCount: Count of failed entries
    - totalEntries: Total number of entries processed
    - message: Array of status messages

    This function uses Write-Log for logging. Ensure the $logFile variable is defined in the calling scope.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'ShouldProcess is invoked within the registry value removal logic')]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [hashtable[]]$RegistryEntries
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting registry removal operation."
    Write-Log -LogFile $logFile -Module $functionName -Message "Starting registry removal operation for $($RegistryEntries.Count) entries." -LogLevel "Information"
    
    # Initialize return object with only relevant properties for removal
    $returnObject = @{
        allRemoved            = $false
        allEntriesProcessed   = $false
        removedKeys           = @()        
        removedKeysCount      = 0    
        entriesUnchanged      = @()
        entriesUnchangedCount = 0
        entriesFailed         = @()
        entriesFailedCount    = 0
        totalEntries          = $RegistryEntries.Count
        message               = @()
    }
    
    Write-Verbose "[$functionName] Processing $($RegistryEntries.Count) registry entries for removal."
    Write-Log -LogFile $logFile -Module $functionName -Message "Processing $($RegistryEntries.Count) registry entries for removal." -LogLevel "Information"
    
    foreach ($entry in $RegistryEntries)
    {
        # Validate required parameters
        if (-not $entry.Path -or -not $entry.Name)
        {
            $errorMsg = "Invalid entry: Path and Name are required. Entry: $($entry | ConvertTo-Json -Compress)"
            Write-Warning "[$functionName] $errorMsg"
            Write-Log -LogFile $logFile -Module $functionName -Message $errorMsg -LogLevel "Error"
            $returnObject.entriesFailed += $entry
            $returnObject.entriesFailedCount++
            $returnObject.message += $errorMsg
            continue
        }
        
        $path = $entry.Path
        $name = $entry.Name
        
        # Default to CurrentUser if not specified
        $hiveValue = if ($entry.Hive) { $entry.Hive } else { "CurrentUser" }
        $hive = if ($hiveValue -eq "LocalMachine") { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
        $hiveName = if ($hiveValue -eq "LocalMachine") { "HKLM" } else { "HKCU" }
        
        # Handle default value name
        $isDefaultValue = ($name -eq "(Default)" -or $name -eq "@")
        $registryValueName = if ($isDefaultValue) { "" } else { $name }
        $displayName = if ($isDefaultValue) { "(Default)" } else { $name }
        $fullPath = "${hiveName}:\$path\$displayName"
        
        Write-Verbose "[$functionName] Processing removal of registry entry: $fullPath"
        Write-Log -LogFile $logFile -Module $functionName -Message "Processing removal of registry entry: $fullPath" -LogLevel "Information"
        
        try
        {
            # Check if the key exists
            $regKey = $hive.OpenSubKey($path, $true)
            
            if ($null -eq $regKey)
            {
                Write-Verbose "[$functionName] Registry key does not exist: ${hiveName}:\$path"
                Write-Log -LogFile $logFile -Module $functionName -Message "Registry key does not exist: ${hiveName}:\$path" -LogLevel "Information"
                Write-Host "Registry key does not exist: $fullPath" -ForegroundColor Yellow
                $returnObject.entriesUnchanged += $entry
                $returnObject.entriesUnchangedCount++
                continue
            }
            
            try
            {
                # Check if the value exists
                $valueExists = $false
                try
                {
                    $currentValue = $regKey.GetValue($registryValueName)
                    
                    # Special handling for default values
                    if ($isDefaultValue)
                    {
                        $valueExists = ($null -ne $currentValue -and $currentValue -ne "")
                        Write-Verbose "[$functionName] Default value check - Current: '$currentValue', Exists: $valueExists"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Default value check - Current: '$currentValue', Exists: $valueExists" -LogLevel "Verbose"
                    }
                    else
                    {
                        $valueExists = ($null -ne $currentValue)
                        Write-Verbose "[$functionName] Value check - Current: '$currentValue', Exists: $valueExists"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Value check - Current: '$currentValue', Exists: $valueExists" -LogLevel "Verbose"
                    }
                }
                catch
                {
                    Write-Verbose "[$functionName] Value does not exist (caught exception): $($_.Exception.Message)"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Value does not exist: $fullPath" -LogLevel "Verbose"
                }
                
                if ($valueExists)
                {
                    # Apply ShouldProcess confirmation
                    $target = $fullPath
                    $action = "Remove registry value"
                    
                    if ($PSCmdlet.ShouldProcess($target, $action))
                    {
                        # Remove the value
                        Write-Verbose "[$functionName] Removing registry value: $fullPath"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Removing registry value: $fullPath" -LogLevel "Information"
                        
                        $regKey.DeleteValue($registryValueName)
                        
                        Write-Host "Removed registry value: $fullPath" -ForegroundColor Green
                        Write-Log -LogFile $logFile -Module $functionName -Message "Successfully removed registry value: $fullPath" -LogLevel "Information"
                        $returnObject.removedKeys += $entry
                        $returnObject.removedKeysCount++
                    }
                    else
                    {
                        Write-Verbose "[$functionName] Removal cancelled by user (ShouldProcess): $fullPath"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Removal cancelled by user: $fullPath" -LogLevel "Information"
                        $returnObject.entriesUnchanged += $entry
                        $returnObject.entriesUnchangedCount++
                    }
                }
                else
                {
                    Write-Verbose "[$functionName] Registry value does not exist: $fullPath"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Registry value does not exist: $fullPath" -LogLevel "Information"
                    Write-Host "Registry value already removed or does not exist: $fullPath" -ForegroundColor Yellow
                    $returnObject.entriesUnchanged += $entry
                    $returnObject.entriesUnchangedCount++
                }
            }
            finally
            {
                if ($null -ne $regKey)
                {
                    $regKey.Close()
                    Write-Verbose "[$functionName] Closed registry key: ${hiveName}:\$path"
                }
            }
        }
        catch
        {
            $errorMsg = "Failed to remove registry entry: $fullPath. Error: $($_.Exception.Message)"
            $errorDetails = "Error details - Path: $path, Name: $displayName, Hive: $hiveName, Exception: $($_.Exception.GetType().FullName), StackTrace: $($_.ScriptStackTrace)"
            Write-Warning "[$functionName] $errorMsg"
            Write-Log -LogFile $logFile -Module $functionName -Message $errorMsg -LogLevel "Error"
            Write-Log -LogFile $logFile -Module $functionName -Message $errorDetails -LogLevel "Error"
            $returnObject.entriesFailed += $entry
            $returnObject.entriesFailedCount++
            $returnObject.message += $errorMsg
        }
    }
    
    # Generate summary and determine overall success
    $summaryMsg = "Registry removal summary: Total=$($returnObject.totalEntries), Removed=$($returnObject.removedKeysCount), AlreadyRemoved=$($returnObject.entriesUnchangedCount), Failed=$($returnObject.entriesFailedCount)"
    Write-Verbose "[$functionName] $summaryMsg"
    Write-Log -LogFile $logFile -Module $functionName -Message $summaryMsg -LogLevel "Information"
    
    if ($returnObject.entriesFailedCount -eq 0)
    {
        $returnObject.allEntriesProcessed = $true
        $returnObject.allRemoved = $true
        
        if ($returnObject.removedKeysCount -gt 0)
        {
            $successMsg = "Successfully removed $($returnObject.removedKeysCount) of $($returnObject.totalEntries) registry entries."
            Write-Host $successMsg -ForegroundColor Green
        }
        else
        {
            $successMsg = "All $($returnObject.totalEntries) registry entries were already removed or do not exist."
            Write-Host $successMsg -ForegroundColor Green
        }
        
        $returnObject.message += $successMsg
        Write-Log -LogFile $logFile -Module $functionName -Message $successMsg -LogLevel "Information"
    }
    else
    {
        $returnObject.allEntriesProcessed = $false
        $returnObject.allRemoved = $false
        $failureMsg = "Processed $($returnObject.totalEntries) registry entries for removal with $($returnObject.entriesFailedCount) failures."
        $returnObject.message += $failureMsg
        Write-Host $failureMsg -ForegroundColor Yellow
        Write-Log -LogFile $logFile -Module $functionName -Message $failureMsg -LogLevel "Warning"
    }
    
    Write-Verbose "[$functionName] Registry removal operation completed."
    Write-Log -LogFile $logFile -Module $functionName -Message "Registry removal operation completed. Result: AllProcessed=$($returnObject.allEntriesProcessed), AllRemoved=$($returnObject.allRemoved)" -LogLevel "Information"
    
    return $returnObject
}
