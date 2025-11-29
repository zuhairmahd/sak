function Import-RegKeysFromFile()
{
    <#
.SYNOPSIS
    Reads a .reg file and converts it to a format compatible with Set-RegKeys function.

.DESCRIPTION
    Parses a Windows Registry Editor (.reg) file and extracts registry paths, names, values, 
    and types. Returns an array of hashtables that can be passed to Set-RegKeys function.
    
    Supports:
    - HKEY_LOCAL_MACHINE (HKLM) and HKEY_CURRENT_USER (HKCU) hives
    - String, DWord, Binary, ExpandString, MultiString, and QWord value types
    - Default values (@)
    - Hex-encoded strings and data

.PARAMETER FilePath
    Path to the .reg file to import

.EXAMPLE
    $regEntries = Import-RegKeysFromFile -FilePath "C:\path\to\file.reg"
    Set-RegKeys -RegistryEntries $regEntries

.OUTPUTS
    Returns an array of hashtables with Path, Name, Value, Type, and Hive properties

.NOTES
    This function uses Write-Log for logging. Ensure the $logFile variable is defined in the calling scope.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    
    if (-not (Test-Path -Path $FilePath))
    {
        $errorMsg = "The specified file path '$FilePath' does not exist."
        Write-Warning "[$functionName] $errorMsg"
        Write-Log -LogFile $logFile -Module $functionName -Message $errorMsg -LogLevel "Error"
        throw $errorMsg
    }
    
    Write-Verbose "[$functionName] Reading registry file: $FilePath"
    Write-Log -LogFile $logFile -Module $functionName -Message "Reading registry file: $FilePath" -LogLevel "Information"
    
    $registryEntries = @()
    $content = Get-Content -Path $FilePath -Raw
    
    # Remove BOM if present and normalize line endings
    $content = $content -replace '^\xEF\xBB\xBF', ''
    $lines = $content -split "`r?`n"
    
    $currentPath = $null
    $currentHive = $null
    $i = 0
    
    while ($i -lt $lines.Count)
    {
        $line = $lines[$i].Trim()
        
        # Skip empty lines and comments
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(';'))
        {
            $i++
            continue
        }
        
        # Check for registry path line [HKEY_...]
        if ($line -match '^\[(.+)\]$')
        {
            $fullPath = $Matches[1]
            Write-Verbose "[$functionName] Found registry path: $fullPath"
            
            # Parse hive and path
            if ($fullPath -match '^HKEY_LOCAL_MACHINE\\(.+)$' -or $fullPath -match '^HKLM\\(.+)$')
            {
                $currentHive = "LocalMachine"
                $currentPath = $Matches[1]
            }
            elseif ($fullPath -match '^HKEY_CURRENT_USER\\(.+)$' -or $fullPath -match '^HKCU\\(.+)$')
            {
                $currentHive = "CurrentUser"
                $currentPath = $Matches[1]
            }
            else
            {
                Write-Warning "[$functionName] Unsupported registry hive in path: $fullPath"
                Write-Log -LogFile $logFile -Module $functionName -Message "Unsupported registry hive in path: $fullPath" -LogLevel "Warning"
                $currentPath = $null
                $currentHive = $null
            }
            
            $i++
            continue
        }
        
        # Parse registry value lines
        if ($currentPath -and $line -match '^"?([^"=]+)"?\s*=\s*(.+)$')
        {
            $valueName = $Matches[1]
            $valueData = $Matches[2].Trim()
            
            # Handle default value
            if ($valueName -eq '@')
            {
                $valueName = '(Default)'
            }
            else
            {
                # Remove quotes from value name
                $valueName = $valueName -replace '^"(.*)"$', '$1'
            }
            
            $valueType = "String"
            $finalValue = $null
            
            # Parse value type and data
            if ($valueData -match '^"(.*)"$')
            {
                # String value
                $valueType = "String"
                $finalValue = $Matches[1] -replace '\\\\', '\' -replace '\\"', '"'
            }
            elseif ($valueData -match '^dword:([0-9a-fA-F]{8})$')
            {
                # DWORD value
                $valueType = "DWord"
                $finalValue = [Convert]::ToInt32($Matches[1], 16)
            }
            elseif ($valueData -match '^hex:(.+)$')
            {
                # Binary data
                $valueType = "Binary"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1].Trim() -match '^\s+(.+)$')
                {
                    $i++
                    $hexString += $Matches[1]
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ } | ForEach-Object { [Convert]::ToByte($_, 16) }
                $finalValue = $hexBytes
            }
            elseif ($valueData -match '^hex\(2\):(.+)$')
            {
                # REG_EXPAND_SZ (ExpandString)
                $valueType = "ExpandString"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1].Trim() -match '^\s+(.+)$')
                {
                    $i++
                    $hexString += $Matches[1]
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ } | ForEach-Object { [Convert]::ToByte($_, 16) }
                $finalValue = [System.Text.Encoding]::Unicode.GetString($hexBytes) -replace '\x00+$', ''
            }
            elseif ($valueData -match '^hex\(7\):(.+)$')
            {
                # REG_MULTI_SZ (MultiString)
                $valueType = "MultiString"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1].Trim() -match '^\s+(.+)$')
                {
                    $i++
                    $hexString += $Matches[1]
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ } | ForEach-Object { [Convert]::ToByte($_, 16) }
                $finalValue = ([System.Text.Encoding]::Unicode.GetString($hexBytes) -replace '\x00+$', '') -split '\x00'
            }
            elseif ($valueData -match '^hex\(b\):(.+)$')
            {
                # REG_QWORD (QWord)
                $valueType = "QWord"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1].Trim() -match '^\s+(.+)$')
                {
                    $i++
                    $hexString += $Matches[1]
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ } | ForEach-Object { $_ }
                $finalValue = [Convert]::ToInt64(($hexBytes[0..7] -join ''), 16)
            }
            else
            {
                Write-Warning "[$functionName] Unable to parse value type for: $valueName = $valueData"
                Write-Log -LogFile $logFile -Module $functionName -Message "Unable to parse value type for: $valueName = $valueData" -LogLevel "Warning"
                $i++
                continue
            }
            
            $entry = @{
                Path  = $currentPath
                Name  = $valueName
                Value = $finalValue
                Type  = $valueType
                Hive  = $currentHive
            }
            
            Write-Verbose "[$functionName] Parsed entry: Path=$currentPath, Name=$valueName, Value=$finalValue, Type=$valueType, Hive=$currentHive"
            Write-Log -LogFile $logFile -Module $functionName -Message "Parsed entry: Path=$currentPath, Name=$valueName, Value=$finalValue, Type=$valueType, Hive=$currentHive" -LogLevel "Verbose"
            
            $registryEntries += $entry
        }
        
        $i++
    }
    
    Write-Verbose "[$functionName] Successfully parsed $($registryEntries.Count) registry entries from $FilePath"
    Write-Log -LogFile $logFile -Module $functionName -Message "Successfully parsed $($registryEntries.Count) registry entries from $FilePath" -LogLevel "Information"
    
    return $registryEntries
} 