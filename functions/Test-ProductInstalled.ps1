function Test-ProductInstalled()
{
    <#
    .SYNOPSIS
        Tests whether a product is installed on the system by checking registry uninstall keys.

    .DESCRIPTION
        This function searches Windows registry uninstall keys (both 64-bit and 32-bit locations)
        to determine if a specified product is installed. It can perform either a loose match
        (by product name OR product code) or a strict match (by product name AND product code AND version).

        The function returns detailed information about the installed product, including uninstall
        commands, size, install date, publisher, and more.

    .PARAMETER product
        A hashtable containing product information to search for. Expected keys:
        - Name: The display name of the product (supports wildcard matching)
        - productCode: The product code GUID (exact match)
        - Version: The product version (used only with -strictMatch)

    .PARAMETER strictMatch
        When specified, requires all three criteria (Name, productCode, and Version) to match.
        Without this switch, the function matches if either the product name OR product code matches.

    .OUTPUTS
        Returns a PSCustomObject with the following properties:
        - installed: Boolean indicating if the product is installed
        - Name: Display name of the product
        - Version: Version of the installed product
        - UninstallCmd: Full uninstall command string
        - QuietUninstall: Quiet uninstall command string (if available)
        - SizeMB: Estimated size in megabytes
        - SizeKB: Estimated size in kilobytes
        - InstallDate: Installation date
        - RegKey: Registry key name (often the product code GUID)
        - Publisher: Product publisher
        - InstallLocation: Installation directory
        - RegistryPath: Full registry path where product was found
        - UninstallFile: Parsed executable path from uninstall command
        - UninstallArguments: Parsed arguments from uninstall command
        - QuietUninstallFile: Parsed executable path from quiet uninstall command
        - QuietUninstallArguments: Parsed arguments from quiet uninstall command

    .EXAMPLE
        $product = @{
            Name = "Google Chrome"
            productCode = "{12345678-1234-1234-1234-123456789012}"
        }
        $result = Test-ProductInstalled -product $product
        if ($result.installed) {
            Write-Host "Chrome is installed, version: $($result.Version)"
        }

    .EXAMPLE
        $product = @{
            Name = "Adobe Reader"
            productCode = "{AC76BA86-7AD7-1033-7B44-AC0F074E4100}"
            Version = "23.001.20093"
        }
        $result = Test-ProductInstalled -product $product -strictMatch
        # Only returns true if name, product code, AND version all match

    .NOTES
        The function searches the following registry locations:
        - HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (64-bit)
        - HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall (32-bit)
        - HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall (per-user)
    #>
    [CmdletBinding()]
    param(
        [hashtable]$product,
        [switch]$strictMatch
    )

    function ConvertFrom-UninstallCommand()
    {
        <#
    .SYNOPSIS
        Parses an uninstall command string into FilePath and Arguments components.
    .PARAMETER cmd
        The uninstall command string to parse.
    .OUTPUTS
        Returns a PSCustomObject with FilePath and Arguments properties.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$cmd
        )

        $functionName = $MyInvocation.MyCommand.Name
        $filePath = $null
        $arguments = $null
        if ($cmd -match '^"([^"]+)"(.*)$')
        {
            # Quoted path (e.g., "C:\Program Files\App\uninstall.exe" /args)
            $filePath = $matches[1]
            $arguments = $matches[2].Trim()
            Write-Verbose "[$functionName] Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
            write-log -logFile $logFile -Module $functionName -Message "Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
        }
        elseif ($cmd -match '^(MsiExec\.exe)\s+(.*)$')
        {
            # Special case for MsiExec.exe (case-insensitive)
            $filePath = "MsiExec.exe"
            $arguments = $matches[2].Trim()
            Write-Verbose "[$functionName] Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"
            write-log -logFile $logFile -Module $functionName -Message "Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"
        }
        elseif ($cmd -match '^([A-Z]:\\.+\.exe)\s+(.*)$')
        {
            # Unquoted full path with .exe extension
            Write-Verbose "[$functionName] Parsing unquoted full path with .exe extension: $cmd"
            write-log -logFile $logFile -Module $functionName -Message "Parsing unquoted full path with .exe extension: $cmd"
            $exeIndex = $cmd.LastIndexOf('.exe')
            if ($exeIndex -ge 0)
            {
                $filePath = $cmd.Substring(0, $exeIndex + 4).Trim()
                $arguments = $cmd.Substring($exeIndex + 4).Trim()
            }
            else
            {
                $filePath = $matches[1]
                $arguments = $matches[2].Trim()
            }
            Write-Verbose "[$functionName] Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"
            write-log -logFile $logFile -Module $functionName -Message "Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"
        }
        elseif ($cmd -match '^(\S+\.exe)\s*(.*)$')
        {
            # Simple executable name without path
            $filePath = $matches[1]
            $arguments = $matches[2].Trim()
            Write-Verbose "[$functionName] Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"
            write-log -logFile $logFile -Module $functionName -Message "Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"
        }
        else
        {
            # Fallback: treat the whole command as filepath
            $filePath = $cmd.Trim()
            $arguments = ""
            Write-Verbose "[$functionName] Fallback parsing: FilePath='$filePath', Arguments='$arguments'"
            write-log -logFile $logFile -Module $functionName -Message "Fallback parsing: FilePath='$filePath', Arguments='$arguments'"
        }

        return [PSCustomObject]@{
            FilePath  = $filePath
            Arguments = $arguments
        }
    }

    $functionName = $MyInvocation.MyCommand.Name
    # Comprehensive list of uninstall registry keys (64-bit and 32-bit)
    $UninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $productObj = [PSCustomObject]@{
        installed = $false
        message   = 'Product Name and ProductCode are required to check installation status.'
    }
    $productName = $product.Name
    $productCode = $product.productCode
    $productVersion = $product.Version
    Write-Verbose "[$functionName] Checking installation status for product: $productName, Version: $productVersion, ProductCode: $productCode                      "
    write-log -logFile $LogFile -Module $scriptName -Message "Checking installation status for product: $productName, Version: $productVersion, ProductCode: $productCode" -LogLevel "Information"
    if (-not ($productName) -and -not ($productCode))
    {
        Write-Error "[$functionName] $productObj.message"
        write-log -logFile $LogFile -Module $scriptName -Message $productObj.message -LogLevel "Error"
        return $productObj
    }

    try
    {
        foreach ($key in $UninstallKeys)
        {
            Write-Verbose "[$functionName] Checking registry key: $key"
            write-log -logFile $LogFile -Module $scriptName -Message "Checking registry key: $key" -LogLevel "Verbose"
            if (-not (Test-Path $key))
            {
                Write-Verbose "[$functionName] Registry key does not exist: $key"
                write-log -logFile $LogFile -Module $scriptName -Message "Registry key does not exist: $key" -LogLevel "Verbose"
                continue
            }
            $registryItems = Get-ChildItem $key -ErrorAction SilentlyContinue
            Write-Verbose "[$functionName] Found $($registryItems.Count) items in $key"
            write-log -logFile $LogFile -Module $scriptName -Message "Found $($registryItems.Count) items in $key" -LogLevel "Verbose"
            foreach ($item in $registryItems)
            {
                try
                {
                    $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                    Write-Verbose "[$functionName] Processing registry item: $($item.PSChildName)"
                    if ($strictMatch)
                    {
                        Write-Verbose "[$functionName] Performing strict match for product $productName with product code $productCode and version $productVersion."
                        Write-Verbose "[$functionName] Display name: $($props.DisplayName)"
                        Write-Verbose "[$functionName] Display version: $($props.DisplayVersion)"
                        Write-Verbose "[$functionName] Product code: $($item.PSChildName)"
                        $productFound = (($productName -and $props.DisplayName -and $props.DisplayName -like "*$productName*") -and ($productCode -and $item.PSChildName -eq $productCode) -and ($productVersion -and $props.DisplayVersion -eq $productVersion)                )
                        Write-Verbose "[$functionName] Strict match result: $productFound"
                    }
                    else
                    {
                        Write-Verbose "[$functionName] Performing loose match for product $productName with product code $productCode."
                        Write-Verbose "[$functionName] Display name: $($props.DisplayName)"
                        Write-Verbose "[$functionName] Display version: $($props.DisplayVersion)"
                        Write-Verbose "[$functionName] Product code: $($item.PSChildName)"
                        $productFound = (($productName -and $props.DisplayName -and $props.DisplayName -like "*$productName*") -or ($productCode -and $item.PSChildName -eq $productCode))
                        Write-Verbose "[$functionName] Loose match result: $productFound"
                    }
                    if ($productFound)
                    {
                        Write-Verbose "[$functionName] Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion))"
                        write-log -logFile $LogFile -Module $scriptName -Message "Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion), Size: $($props.EstimatedSize) KB)" -LogLevel "Information"
                        $processedUninstallCmd = if ($props.UninstallString)
                        {
                            ConvertFrom-UninstallCommand -cmd $props.UninstallString
                        }
                        else
                        {
                            $null
                        }
                        $processedQuietUninstallCmd = if ($props.QuietUninstallString)
                        {
                            ConvertFrom-UninstallCommand -cmd $props.QuietUninstallString
                        }
                        else
                        {
                            $null
                        }
                        # Create product object with all relevant details
                        $productObj = [PSCustomObject]@{
                            installed               = $true
                            Name                    = $props.DisplayName
                            Version                 = $props.DisplayVersion
                            UninstallCmd            = $props.UninstallString
                            QuietUninstall          = $props.QuietUninstallString
                            SizeMB                  = if ($props.EstimatedSize)
                            {
                                [math]::Round($props.EstimatedSize / 1024, 2)
                            }
                            else
                            {
                                0
                            }
                            SizeKB                  = if ($props.EstimatedSize)
                            {
                                $props.EstimatedSize
                            }
                            else
                            {
                                0
                            }
                            InstallDate             = $props.InstallDate
                            RegKey                  = $item.PSChildName # This is often the Product Code GUID
                            Publisher               = $props.Publisher
                            InstallLocation         = $props.InstallLocation
                            RegistryPath            = $key
                            UninstallFile           = if ($processedUninstallCmd)
                            {
                                $processedUninstallCmd.FilePath
                            }
                            else
                            {
                                $null
                            }
                            UninstallArguments      = if ($processedUninstallCmd)
                            {
                                $processedUninstallCmd.Arguments
                            }
                            else
                            {
                                $null
                            }
                            QuietUninstallFile      = if ($processedQuietUninstallCmd)
                            {
                                $processedQuietUninstallCmd.FilePath
                            }
                            else
                            {
                                $null
                            }
                            QuietUninstallArguments = if ($processedQuietUninstallCmd)
                            {
                                $processedQuietUninstallCmd.Arguments
                            }
                            else
                            {
                                $null
                            }
                        }
                    }
                }
                catch
                {
                    Write-Verbose "[$functionName] Error processing registry item $($item.PSPath): $_"
                    write-log -logFile $LogFile -Module $scriptName -Message "Error processing registry item $($item.PSPath): $_" -LogLevel "Warning"
                }
            }
        }
    }
    catch
    {
        Write-Error "[$functionName] Error occurred while searching for product '$productName': $_"
        write-log -logFile $LogFile -Module $scriptName -Message "Error occurred while searching for product '$productName': $_" -LogLevel "Error"
        $productObj = [PSCustomObject]@{
            installed = $false
            message   = "Error occurred while searching for product: $_"
        }
    }

    return $productObj
}
