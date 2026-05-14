function Import-CSVRegKeys()
{
    <#
    .SYNOPSIS
        Imports registry keys from a CSV file.

    .DESCRIPTION
        Reads a CSV file containing registry key definitions and converts them into PowerShell objects.
        The function processes registry paths, value names, types, and data from the CSV file.
        It supports both HKEY_CURRENT_USER and HKEY_LOCAL_MACHINE hives and normalizes registry type names.

    .PARAMETER FilePath
        The full path to the CSV file containing registry key definitions.
        Required. The CSV file must contain the following columns:
        - RegistryPath: The full registry path (e.g., HKEY_LOCAL_MACHINE\Software\MyApp)
        - ValueName: The name of the registry value (leave empty for default value)
        - ValueType: The registry value type (REG_SZ, REG_DWORD, REG_BINARY, etc.)
        - ValueData: The data to set for the registry value

    .OUTPUTS
        PSCustomObject
        Returns an object with the following properties:
        - success: Boolean indicating if all registry keys were processed successfully
        - regKeys: Array of PSCustomObject containing Path, Name, Type, Value, and Hive properties

    .EXAMPLE
        $result = Import-CSVRegKeys -FilePath "C:\config\registry-settings.csv"
        if ($result.success) {
            foreach ($key in $result.regKeys) {
                Write-Host "Path: $($key.Path), Name: $($key.Name), Type: $($key.Type)"
            }
        }

    .EXAMPLE
        Import-CSVRegKeys -FilePath "C:\config\registry-settings.csv" -Verbose

    .NOTES
        The function includes a helper function ConvertTo-RegistryValueKind to convert Windows registry
        type names (REG_SZ, REG_DWORD, etc.) to .NET RegistryValueKind names.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Importing registry keys from CSV file: $FilePath"
    write-log -logFile $logFile -module $functionName -message "Importing registry keys from CSV file: $FilePath" -logLevel "INFORMATION"

    # Helper function to convert Windows registry type names to RegistryValueKind names
    function ConvertTo-RegistryValueKind()
    {
        [CmdletBinding()]
        param([string]$RegistryType)

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Converting registry type: $RegistryType"
        write-log -logFile $logFile -module $functionName -message "Converting registry type: $RegistryType" -logLevel "Information"
        # Map common Windows registry type names to .NET RegistryValueKind names
        $typeMap = @{
            'REG_SZ'        = 'String'
            'REG_EXPAND_SZ' = 'ExpandString'
            'REG_BINARY'    = 'Binary'
            'REG_DWORD'     = 'DWord'
            'REG_QWORD'     = 'QWord'
            'REG_MULTI_SZ'  = 'MultiString'
            # Also support RegistryValueKind names directly (case-insensitive)
            'String'        = 'String'
            'ExpandString'  = 'ExpandString'
            'Binary'        = 'Binary'
            'DWord'         = 'DWord'
            'QWord'         = 'QWord'
            'MultiString'   = 'MultiString'
        }

        $normalizedType = $typeMap[$RegistryType]
        Write-Verbose "[$functionName] Normalized registry type: $normalizedType"
        write-log -logFile $logFile -module $functionName -message "Normalized registry type: $normalizedType" -logLevel "Information"
        if ($normalizedType)
        {
            Write-Verbose "[$functionName] Returning normalized registry type: $normalizedType"
            write-log -logFile $logFile -module $functionName -message "Returning normalized registry type: $normalizedType" -logLevel "Information"
            return $normalizedType
        }
        else
        {
            Write-Warning "[$functionName] Unknown registry type '$RegistryType', defaulting to 'String'"
            write-log -logFile $logFile -module $functionName -message "Unknown registry type '$RegistryType', defaulting to 'String'" -logLevel "Warning"
            return 'String'
        }
    }

    $returnObject = @{
        success = $false
        regKeys = @()
    }
    $regKeys = @()
    if (Test-Path -Path $FilePath)
    {
        $csvData = Import-Csv -Path $FilePath
        Write-Verbose "Imported $($csvData.Count) rows from CSV file."
        write-log -logFile $logFile -module $functionName -message "Imported $($csvData.Count) rows from CSV file." -logLevel "Information"
        foreach ($row in $csvData)
        {
            $name = if ($null -ne $row.ValueName -and $row.ValueName -ne '')
            {
                $row.ValueName
            }
            else
            {
                '@'
            }
            Write-Verbose "[$functionName] Normalizing registry key name: $name"
            write-log -logFile $logFile -module $functionName -message "Normalizing registry key name: $name" -logLevel "Information"
            $path = $row.RegistryPath
            Write-Verbose "[$functionName] Extracting hive from path: $path"
            write-log -logFile $logFile -module $functionName -message "Extracting hive from path: $path" -logLevel "Information"
            $hive = if ($path -like "HKEY_CURRENT_USER\*")
            {
                "CurrentUser"
            }
            elseif ($path -like "HKEY_LOCAL_MACHINE\*")
            {
                "LocalMachine"
            }
            else
            {
                "UNKNOWN"
            }
            Write-Verbose "[$functionName] Extracted hive: $hive"
            write-log -logFile $logFile -module $functionName -message "Extracted hive: $hive" -logLevel "Information"
            if ($null -ne $hive -and $hive -ne "UNKNOWN")
            {
                #remove the hive from the $row.registryPath
                Write-Verbose "[$functionName] Removing hive from path: $path"
                write-log -logFile $logFile -module $functionName -message "Removing hive from path: $path" -logLevel "Information"
                $path = $path -replace "HKEY_CURRENT_USER\\", "" -replace "HKEY_LOCAL_MACHINE\\", ""
                Write-Verbose "[$functionName] Updated path: $path"
                write-log -logFile $logFile -module $functionName -message "Updated path: $path" -logLevel "Information"
            }
            $regKey = [PSCustomObject]@{
                Path  = $path
                Name  = $name
                Type  = ConvertTo-RegistryValueKind -RegistryType $row.ValueType
                Value = $row.ValueData
                Hive  = if ($null -ne $hive -and $hive -ne "UNKNOWN")
                {
                    $hive
                }
                else
                {
                    $null
                }
            }
            $regKeys += $regKey
        }
        Write-Verbose "Processed $($regKeys.Count) registry keys from CSV file."
        write-log -logFile $logFile -module $functionName -message "Processed $($regKeys.Count) registry keys from CSV file." -logLevel "Information"
        $returnObject.regKeys = $regKeys
        if ($regKeys.count -eq $csvData.count)
        {
            Write-Verbose "[$functionName] Successfully processed all registry keys from CSV file."
            write-log -logFile $logFile -module $functionName -message "Successfully processed all registry keys from CSV file." -logLevel "Information"
            $returnObject.success = $true
        }
        else
        {
            Write-Error "[$functionName] Mismatch in processed registry keys count. Processed: $($regKeys.Count), Expected: $($csvData.Count)"
            write-log -logFile $logFile -module $functionName -message "Mismatch in processed registry keys count. Processed: $($regKeys.Count), Expected: $($csvData.Count)" -logLevel "Error"
            $returnObject.success = $false
        }
    }
    else
    {
        Write-Error "File not found: $FilePath"
        write-log -logFile $logFile -module $functionName -message "File not found: $FilePath" -logLevel "Error"
    }
    return $returnObject
}