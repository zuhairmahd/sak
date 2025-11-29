function Set-RegKey()
<#
.SYNOPSIS
    Ensures a registry key exists and sets its value to the specified value.

.DESCRIPTION
    The Set-RegKey function checks if a registry key exists at the given path. If it does not exist, it creates the key.
    It then checks if the specified value for the key matches the desired value. If not, it updates the key with the desired value.
    All actions are logged using Write-Log and output to the host.

.PARAMETER keyPath
    The path to the registry key.

.PARAMETER Name
    The name of the registry value to set.

.PARAMETER Value
    The value to assign to the registry key.

.OUTPUTS
    [bool] Returns $true if the registry key was set successfully, otherwise $false.

.EXAMPLE
    Set-RegKey -keyPath "HKLM:\Software\MyApp" -Name "Setting" -Value "Enabled"

.NOTES
    Requires administrative privileges to modify registry keys.
    This function is part of a PowerShell module and uses Write-Log for logging.
    Ensure that the Write-Log function is defined in the same module or imported before using this function.
#>
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$keyPath,
        [string]$Name,
        [string]$Value,
        [string]$Type,
        [switch]$CheckOnly
    )
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Message "Executing function $functionName with parameters: keyPath='$keyPath', Name='$Name', Value='$Value'" -LogFile $logFile -Module $functionName -LogLevel "Information"
    Write-Host "Checking registry key $Name at path $keyPath for the value $Value"
    Write-Log -Message "Checking registry key $Name at path $keyPath for the value $Value" -LogFile $logFile -Module $functionName -LogLevel "Information"

    # If -CheckOnly is specified, do not modify anything; just validate and return true/false
    if ($CheckOnly)
    {
        Write-Host "CheckOnly is set. No changes will be made."
        Write-Log -Message "CheckOnly is set. Validation only; no changes will be made." -LogFile $logFile -Module $functionName -LogLevel "Information"

        if (-not (Test-Path -Path $keyPath))
        {
            Write-Host "Registry path $keyPath does not exist."
            Write-Log -Message "Registry path $keyPath does not exist." -LogFile $logFile -Module $functionName -LogLevel "Information"
            return $false
        }

        $item = Get-ItemProperty -Path $keyPath -Name $Name -ErrorAction SilentlyContinue
        if (-not $item)
        {
            Write-Host "Registry value $Name does not exist at $keyPath."
            Write-Log -Message "Registry value $Name does not exist at $keyPath." -LogFile $logFile -Module $functionName -LogLevel "Information"
            return $false
        }

        $currentKeyValue = $item.$Name
        # Normalize to string for a predictable comparison; this avoids array semantics surprises
        $currentString = if ($null -ne $currentKeyValue) { -join ([string[]]$currentKeyValue) } else { $null }
        $desiredString = if ($null -ne $Value) { -join ([string[]]$Value) } else { $null }

        $isMatch = ($currentString -eq $desiredString)
        Write-Log -Message "CheckOnly comparison result for $Name at $keyPath. Current='$currentString' Desired='$desiredString' Match=$isMatch" -LogFile $logFile -Module $functionName -LogLevel "Information"
        return [bool]$isMatch
    }
    #build a splat for the New-ItemProperty command
    $splat = @{
        Path  = $keyPath
        Name  = $Name
        Value = $Value
        Force = $true
    }
    if ($Type)
    {
        #Parse the type and make sure it is one of "String, ExpandString, Binary, DWord, MultiString, QWord, Unknown".
        #Map commonly used registry types to the expected values.
        switch ($Type.ToLower())
        {
            "reg_sz" { $splatType = "String" }
            "reg_expandstring" { $splatType = "ExpandString" }
            "reg_binary" { $splatType = "Binary" }
            "reg_dword" { $splatType = "DWord" }
            "reg_multistring" { $splatType = "MultiString" }
            "reg_qword" { $splatType = "QWord" }
            default { Write-Host "Unknown registry type: $Type. Passing as is." -ForegroundColor Yellow; $splatType = $Type }
        }
        $splat += @{
            Type = $splatType
        }
    }
    if (-not (Test-Path -Path $keyPath))
    {
        Write-Host "Registry key $keyPath does not exist. Creating it now."
        Write-Log -Message "Registry key $keyPath does not exist. Creating it now." -LogFile $logFile -Module $functionName -LogLevel "Information"
        try
        {
            New-Item -Path $keyPath -Force | Out-Null
            Write-Host "Registry key $keyPath created successfully."
            Write-Log -Message "Registry key $keyPath created successfully." -LogFile $logFile -Module $functionName -LogLevel "Information"
        }
        catch
        {
            Write-Host "Failed to create registry key $keyPath. Error: $_"
            Write-Log -Message "Failed to create registry key $keyPath. Error: $_" -LogFile $logFile -Module $functionName -LogLevel "Error"
            return $false
        }
    }
    Write-Host "Checking correct value assignment:"
    Write-Log -Message "Checking correct value assignment for key $Name at path $keyPath" -LogFile $logFile -Module $functionName -LogLevel "Information"
    $desiredKeyName = $Name
    $desiredKeyValue = $Value
    Write-Host "Desired key name: $desiredKeyName"
    Write-Log -Message "Desired key name: $desiredKeyName" -LogFile $logFile -Module $functionName -LogLevel "Information"
    Write-Host "Desired key value: $desiredKeyValue"
    Write-Log -Message "Desired key value: $desiredKeyValue" -LogFile $logFile -Module $functionName -LogLevel "Information"
    Write-Host "Getting current key value for $Name at path $keyPath"
    $item = Get-ItemProperty -Path $keyPath -Name $Name -ErrorAction SilentlyContinue
    if (-not $item)
    {
        Write-Host "Registry key $Name does not exist."
        Write-Host "Creating registry key $Name with value $desiredKeyValue at path $keyPath"
        Write-Log -Message "Registry key $Name does not exist. Creating it with value $desiredKeyValue at path $keyPath" -LogFile $logFile -Module $functionName -LogLevel "Information"
        try
        {
            New-ItemProperty @splat | Out-Null
            Write-Host "Registry key $Name created successfully with value $desiredKeyValue."
            Write-Log -Message "Registry key $Name created successfully with value $desiredKeyValue." -LogFile $logFile -Module $functionName -LogLevel "Information"
        }
        catch
        {
            Write-Host "Failed to create registry key $Name. Error: $_"
            Write-Log -Message "Failed to create registry key $Name. Error: $_" -LogFile $logFile -Module $functionName -LogLevel "Error"
            return $false
        }
    }
    else
    {
        #check if the values match, otherwise set them.
        $currentKeyValue = $item.$Name
        $currentKeyType = $item.PSObject.Properties[$Name].TypeNameOfValue
        Write-Log -Message "Current key value for $Name at path $($keyPath): $currentKeyValue" -LogFile $logFile -Module $functionName -LogLevel "Information"
        Write-Log -Message "Current key type for $Name at path $($keyPath): $currentKeyType" -LogFile $logFile -Module $functionName -LogLevel "Information"
        Write-Host "Current key value: $currentKeyValue"
        Write-Host "Current key type: $currentKeyType"
        Write-Log -Message "Current key value: $currentKeyValue" -LogFile $logFile -Module $functionName -LogLevel "Information"
        Write-Log -Message "Checking if the registry key $Name at path $keyPath matches the desired value $desiredKeyValue" -LogFile $logFile -Module $functionName -LogLevel "Information"
        if ($currentKeyValue -ne $desiredKeyValue)
        {
            Write-Host "Registry key $Name exists but does not match the expected value."
            Write-Host "Updating registry key $Name with value $desiredKeyValue at path $keyPath"
            Write-Log -Message "Registry key $Name exists but does not match the expected value. Updating it with value $desiredKeyValue at path $keyPath" -LogFile $logFile -Module $functionName -LogLevel "Information"
            try
            {
                Set-ItemProperty @splat | Out-Null
                Write-Host "Registry key $Name updated successfully with value $desiredKeyValue."
                Write-Log -Message "Registry key $Name updated successfully with value $desiredKeyValue." -LogFile $logFile -Module $functionName -LogLevel "Information"
            }
            catch
            {
                Write-Host "Failed to update registry key $Name. Error: $_"
                Write-Log -Message "Failed to update registry key $Name. Error: $_" -LogFile $logFile -Module $functionName -LogLevel "Error"
                return $false
            }
        }
        else
        {
            Write-Host "Registry key $Name already exists with the correct value and type."
            Write-Log -Message "Registry key $Name already exists with the correct value and type." -LogFile $logFile -Module $functionName -LogLevel "Information"
            return $true
        }
    }
    return $true
}