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
            Write-Log -LogFile $logFile -Module $functionName -Message "Found registry path: $fullPath" -LogLevel "Verbose"
            
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
        # Match either: @=value, "quoted name"=value (handles spaces and special chars)
        if ($currentPath -and $line -match '^(@|"([^"]+)")\s*=\s*(.*)$')
        {
            # $Matches[1] = entire name part (@, or "quoted")
            # $Matches[2] = content inside quotes (if quoted, empty if @)
            # $Matches[3] = value data (everything after =)
            
            if ($Matches[1] -eq '@')
            {
                $valueName = '(Default)'
            }
            else
            {
                # Quoted value name (handles spaces, special characters)
                $valueName = $Matches[2]
            }
            
            $valueData = $Matches[3].Trim()
            
            $valueType = "String"
            $finalValue = $null
            
            # Parse value type and data
            if ($valueData -match '^"(.*)"$')
            {
                # String value (including empty strings)
                $valueType = "String"
                # Handle escape sequences: \\ -> \, \" -> ", and preserve the value
                $finalValue = $Matches[1] -replace '\\"', '"' -replace '\\\\', '\'
            }
            elseif ($valueData -match '^dword:([0-9a-fA-F]{1,8})$')
            {
                # DWORD value (handles both full 8 digits and shorter representations)
                $valueType = "DWord"
                # Pad to 8 digits if needed
                $hexValue = $Matches[1].PadLeft(8, '0')
                $finalValue = [Convert]::ToInt32($hexValue, 16)
            }
            elseif ($valueData -match '^hex:(.+)$')
            {
                # Binary data (REG_BINARY)
                $valueType = "Binary"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values (continuation lines start with spaces and end with backslash)
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    # Remove trailing backslash if present
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                # Remove all whitespace and backslashes, then split by comma
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' } | ForEach-Object { 
                    try { [Convert]::ToByte($_, 16) } 
                    catch { Write-Warning "[$functionName] Invalid hex byte: $_"; $null }
                } | Where-Object { $null -ne $_ }
                $finalValue = $hexBytes
            }
            elseif ($valueData -match '^hex\(0\):(.+)$')
            {
                # REG_NONE (hex(0)) - treat as binary
                $valueType = "Binary"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' } | ForEach-Object { 
                    try { [Convert]::ToByte($_, 16) } 
                    catch { Write-Warning "[$functionName] Invalid hex byte: $_"; $null }
                } | Where-Object { $null -ne $_ }
                $finalValue = $hexBytes
            }
            elseif ($valueData -match '^hex\(1\):(.+)$')
            {
                # REG_SZ as hex (hex(1)) - convert to string
                $valueType = "String"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' } | ForEach-Object { 
                    try { [Convert]::ToByte($_, 16) } 
                    catch { Write-Warning "[$functionName] Invalid hex byte: $_"; $null }
                } | Where-Object { $null -ne $_ }
                
                if ($hexBytes.Count -gt 0)
                {
                    $finalValue = [System.Text.Encoding]::Unicode.GetString($hexBytes) -replace '\x00+$', ''
                }
                else
                {
                    $finalValue = ''
                }
            }
            elseif ($valueData -match '^hex\(2\):(.+)$')
            {
                # REG_EXPAND_SZ (ExpandString)
                $valueType = "ExpandString"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values (continuation lines start with spaces and end with backslash)
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    # Remove trailing backslash if present
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                # Remove all whitespace and backslashes, then split by comma
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' } | ForEach-Object { 
                    try { [Convert]::ToByte($_, 16) } 
                    catch { Write-Warning "[$functionName] Invalid hex byte: $_"; $null }
                } | Where-Object { $null -ne $_ }
                
                if ($hexBytes.Count -gt 0)
                {
                    $finalValue = [System.Text.Encoding]::Unicode.GetString($hexBytes) -replace '\x00+$', ''
                }
                else
                {
                    $finalValue = ''
                }
            }
            elseif ($valueData -match '^hex\(7\):(.+)$')
            {
                # REG_MULTI_SZ (MultiString)
                $valueType = "MultiString"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values (continuation lines start with spaces and end with backslash)
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    # Remove trailing backslash if present
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                # Remove all whitespace and backslashes, then split by comma
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' } | ForEach-Object { 
                    try { [Convert]::ToByte($_, 16) } 
                    catch { Write-Warning "[$functionName] Invalid hex byte: $_"; $null }
                } | Where-Object { $null -ne $_ }
                
                if ($hexBytes.Count -gt 0)
                {
                    $unicodeString = [System.Text.Encoding]::Unicode.GetString($hexBytes) -replace '\x00+$', ''
                    $finalValue = $unicodeString -split '\x00' | Where-Object { $_ -ne '' }
                }
                else
                {
                    $finalValue = @()
                }
            }
            elseif ($valueData -match '^hex\(b\):(.+)$')
            {
                # REG_QWORD (QWord)
                $valueType = "QWord"
                $hexString = $Matches[1]
                
                # Handle multi-line hex values (continuation lines start with spaces and end with backslash)
                while ($i + 1 -lt $lines.Count -and $lines[$i + 1] -match '^\s+(.+)$')
                {
                    $i++
                    $continuationLine = $Matches[1].Trim()
                    # Remove trailing backslash if present
                    $continuationLine = $continuationLine -replace '\\$', ''
                    $hexString += $continuationLine
                }
                
                # Remove all whitespace and backslashes, then split by comma
                $hexString = $hexString -replace '\s', '' -replace '\\', ''
                $hexBytes = $hexString -split ',' | Where-Object { $_ -ne '' }
                
                # QWORD is 8 bytes in little-endian format
                if ($hexBytes.Count -ge 8)
                {
                    try
                    {
                        # Convert little-endian hex bytes to QWORD
                        $qwordValue = 0
                        for ($b = 0; $b -lt 8; $b++)
                        {
                            $byteValue = [Convert]::ToByte($hexBytes[$b], 16)
                            $qwordValue += [int64]$byteValue -shl ($b * 8)
                        }
                        $finalValue = $qwordValue
                    }
                    catch
                    {
                        Write-Warning "[$functionName] Failed to convert QWORD bytes: $hexString"
                        Write-Log -LogFile $logFile -Module $functionName -Message "Failed to convert QWORD bytes: $hexString" -LogLevel "Warning"
                        $finalValue = 0
                    }
                }
                else
                {
                    Write-Warning "[$functionName] Insufficient bytes for QWORD (need 8, got $($hexBytes.Count))"
                    Write-Log -LogFile $logFile -Module $functionName -Message "Insufficient bytes for QWORD (need 8, got $($hexBytes.Count))" -LogLevel "Warning"
                    $finalValue = 0
                }
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