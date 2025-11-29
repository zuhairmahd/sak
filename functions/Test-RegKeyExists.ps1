function Test-RegKeyExists()
<#
.SYNOPSIS
    Checks if specified registry keys exist and have the correct values.

.DESCRIPTION
    This function iterates through an array of registry key objects, verifying the existence of each key and whether its value matches the expected value.
    If a key does not exist or its value does not match, the function returns $false. Otherwise, it returns $true.

.PARAMETER keys
    An array of objects, each containing 'path', 'Name', and 'Value' properties representing the registry key path, the value name, and the expected value.

.OUTPUTS
    [bool] Returns $true if all keys exist and have the correct values, otherwise $false.

.EXAMPLE
    $keys = @(
        @{ path = "HKLM:\Software\MyApp"; Name = "Setting"; Value = "Enabled" }
    )
    Test-RegKeyExists -keys $keys

.NOTES
    Author: MahmoudZ
    Date: 2024-06
#>
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [array]$keys
    )
    
    $allKeysExist = $true
    foreach ($key in $keys)
    {
        Write-Host "Checking registry key: $($key.path)"
        if (-not (Test-Path -Path $key.path))
        {
            Write-Host "Registry key $($key.path) does not exist. Creating it now."
            $allKeysExist = $false
        }
        else
        {
            Write-Host "Registry key $($key.path) already exists."
            Write-Host "Checking correct value assignments:"
            $desiredKeyName = $key.Name
            $desiredKeyValue = $key.Value
            Write-Host "Key name: $desiredKeyName"
            Write-Host "Desired key value: $desiredKeyValue"
            $item = Get-ItemProperty -Path $key.path -Name $key.Name -ErrorAction SilentlyContinue
            $currentKeyValue = $item.$($key.Name)
            Write-Host "Current key value: $currentKeyValue"
            if (-not $item)
            {
                Write-Host "Registry key $($key.Name) does not exist."
                $allKeysExist = $false
            }
            else
            {
                #check if the values match, otherwise set them.
                if ($currentKeyValue -ne $desiredKeyValue)
                {
                    Write-Host "Registry key $($key.Name) exists but does not match the expected value."
                    $allKeysExist = $false
                }
                else
                {
                    Write-Host "Registry key $($key.Name) already exists with the correct value and type."
                }
            }                
        }
    }
    return $allKeysExist
}

