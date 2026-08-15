function Get-UninstallCommand() {
    <#
.SYNOPSIS
Retrieves uninstall command information for installed applications matching one or more keywords.

.DESCRIPTION
Searches common 32-bit and 64-bit Windows uninstall registry locations under HKLM and HKCU for products whose DisplayName contains any of the provided keywords. Builds a list of matching products, de-duplicates them, computes size information, and identifies the most likely main product (largest EstimatedSize). Returns a summary object including all products and the mostLikelyMatch.

.PARAMETER keywords
One or more keywords to match against the DisplayName of installed products. Matching is a contains match (e.g., "*keyword*") and is case-insensitive. If no keywords are provided, the function returns hasErrors = $true with a message.

.OUTPUTS
PSCustomObject
- hasErrors [bool]            Indicates if any errors occurred or if input was invalid.
- message [string]            Informational or error message.
- products [PSCustomObject[]] Collection of matching products with properties:
    Name, Version, UninstallCmd, QuietUninstall, SizeMB, SizeKB, InstallDate,
    RegKey, Publisher, InstallLocation, RegistryPath, IsMostLikely
- mostLikelyMatch [PSCustomObject] The product considered the main app (largest size), if found.

.EXAMPLE
Get-UninstallCommand -keywords "Google", "Chrome"

.EXAMPLE
Get-UninstallCommand -keywords "Zoom"

.NOTES
Requires read access to the following registry locations:
- HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
- HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall
- HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
Uses external helper write-log and assumes $LogFile and $scriptName are available in scope.

.LINK
https://learn.microsoft.com/windows/win32/msi/uninstall-registry-key
    #>
    [CmdletBinding()]
    param(
        [string[]]$keywords
    )

    $functionName = $MyInvocation.MyCommand.Name
    # Comprehensive list of uninstall registry keys (64-bit and 32-bit)
    $UninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $uninstallationCommands = @{
        hasErrors       = $false
        message         = ''
        products        = @()
        mostLikelyMatch = $null
    }
    Write-Verbose "[$functionName] Get uninstallation commands for applications with keywords: $($keywords -join ', ')"
    write-log -logFile $LogFile -Module $scriptName -Message "Get uninstallation commands for applications with keywords: $($keywords -join ', ')" -LogLevel "Information"
    if ($keywords.count -eq 0) {
        Write-Verbose "[$functionName] No keywords provided. Returning empty list."
        write-log -logFile $LogFile -Module $scriptName -Message "No keywords provided. Returning empty list." -LogLevel "Warning"
        $uninstallationCommands.message = "No keywords provided. Returning empty list."
        $uninstallationCommands.hasErrors = $true
        return $uninstallationCommands
    }
    Write-Verbose "[$functionName] Searching for $($keywords.count) keyword(s) across $($UninstallKeys.count) registry locations."
    write-log -logFile $LogFile -Module $scriptName -Message "Searching for $($keywords.count) keyword(s) across $($UninstallKeys.count) registry locations." -LogLevel "Information"
    # Collect all matching products
    $allProducts = @()
    foreach ($keyword in $keywords) {
        Write-Verbose "[$functionName] Searching for products with keyword: '$keyword'"
        write-log -logFile $LogFile -Module $scriptName -Message "Searching for products with keyword: '$keyword'" -LogLevel "Information"
        try {
            foreach ($key in $UninstallKeys) {
                Write-Verbose "[$functionName] Checking registry key: $key"
                write-log -logFile $LogFile -Module $scriptName -Message "Checking registry key: $key" -LogLevel "Verbose"
                if (-not (Test-Path $key)) {
                    Write-Verbose "[$functionName] Registry key does not exist: $key"
                    write-log -logFile $LogFile -Module $scriptName -Message "Registry key does not exist: $key" -LogLevel "Verbose"
                    continue
                }
                $registryItems = Get-ChildItem $key -ErrorAction SilentlyContinue
                Write-Verbose "[$functionName] Found $($registryItems.Count) items in $key"
                write-log -logFile $LogFile -Module $scriptName -Message "Found $($registryItems.Count) items in $key" -LogLevel "Verbose"
                foreach ($item in $registryItems) {
                    try {
                        $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                        if ($keyword -and $props.DisplayName -and $props.DisplayName -like "*$keyword*") {
                            Write-Verbose "[$functionName] Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion))"
                            write-log -logFile $LogFile -Module $scriptName -Message "Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion), Size: $($props.EstimatedSize) KB)" -LogLevel "Information"
                            # Create product object with all relevant details
                            $productObj = [PSCustomObject]@{
                                Name                  = $props.DisplayName
                                Version               = if ($props.DisplayVersion) { $props.DisplayVersion } else { $null }
                                UninstallCmd          = if ($props.UninstallString) { $props.UninstallString } else { $null }
                                QuietUninstall        = if ($props.QuietUninstallString) { $props.QuietUninstallString } else { $null }
                                # Size is stored in Registry as Kilobytes (KB)
                                SizeMB                = if ($props.EstimatedSize) { [math]::Round($props.EstimatedSize / 1024, 2) } else { 0 }
                                SizeKB                = if ($props.EstimatedSize) { $props.EstimatedSize } else { 0 }
                                InstallDate           = $props.InstallDate
                                RegKey                = $item.PSChildName # This is often the Product Code GUID
                                Publisher             = if ($props.Publisher) { $props.Publisher } else { $null }
                                InstallLocation       = $props.InstallLocation
                                BundleCachePath       = if ($props.BundleCachePath) { $props.BundleCachePath } else { $null }
                                BundleUpgradeCode     = if ($props.BundleUpgradeCode) { $props.BundleUpgradeCode } else { $null }
                                BundleAddonCode       = if ($props.BundleAddonCode) { $props.BundleAddonCode } else { $null }
                                BundleDetectCode      = if ($props.BundleDetectCode) { $props.BundleDetectCode } else { $null }
                                BundlePatchCode       = if ($props.BundlePatchCode) { $props.BundlePatchCode } else { $null }
                                BundleVersion         = if ($props.BundleVersion) { $props.BundleVersion } else { $null }
                                VersionMajor          = if ($props.VersionMajor) { $props.VersionMajor } else { $null }
                                VersionMinor          = if ($props.VersionMinor) { $props.VersionMinor } else { $null }
                                BundleProviderKey     = if ($props.BundleProviderKey) { $props.BundleProviderKey } else { $null }
                                BundleTag             = if ($props.BundleTag) { $props.BundleTag } else { $null }
                                EngineVersion         = if ($props.EngineVersion) { $props.EngineVersion } else { $null }
                                EngineProtocolVersion = if ($props.EngineProtocolVersion) { $props.EngineProtocolVersion } else { $null }
                                DisplayIcon           = if ($props.DisplayIcon) { $props.DisplayIcon } else { $null }
                                HelpLink              = if ($props.HelpLink) { $props.HelpLink } else { $null }
                                HelpTelephone         = if ($props.HelpTelephone) { $props.HelpTelephone } else { $null }
                                URLInfoAbout          = if ($props.URLInfoAbout) { $props.URLInfoAbout } else { $null }
                                URLUpdateInfo         = if ($props.URLUpdateInfo) { $props.URLUpdateInfo } else { $null }
                                AuthorizedCDFPrefix   = if ($props.AuthorizedCDFPrefix) { $props.AuthorizedCDFPrefix } else { $null }
                                Comments              = if ($props.Comments) { $props.Comments } else { $null }
                                Contact               = if ($props.Contact) { $props.Contact } else { $null }
                                SettingsIdentifier    = if ($props.SettingsIdentifier) { $props.SettingsIdentifier } else { $null }
                                InstallSource         = if ($props.InstallSource) { $props.InstallSource } else { $null }
                                Readme                = if ($props.Readme) { $props.Readme } else { $null }
                                SystemComponent       = if ($props.SystemComponent -eq 1) { $true } else { $false }
                                WindowsInstaller      = if ($props.WindowsInstaller -eq 1) { $true } else { $false }
                                Language              = if ($props.Language) { $props.Language } else { $null }
                                RegistryPath          = $key
                                IsMostLikely          = $false
                            }
                            $allProducts += $productObj
                            Write-Verbose "[$functionName] Product details: Name='$($productObj.Name)', Size=$($productObj.SizeMB)MB, UninstallCmd='$($productObj.UninstallCmd)'"
                            write-log -logFile $LogFile -Module $scriptName -Message "Product details: RegKey='$($productObj.RegKey)', Size=$($productObj.SizeMB)MB, Publisher='$($productObj.Publisher)'" -LogLevel "Verbose"
                        }
                    }
                    catch {
                        Write-Verbose "[$functionName] Error processing registry item $($item.PSPath): $_"
                        write-log -logFile $LogFile -Module $scriptName -Message "Error processing registry item $($item.PSPath): $_" -LogLevel "Warning"
                    }
                }
            }
        }
        catch {
            Write-Error "[$functionName] Error occurred while searching for products with keyword '$keyword': $_"
            write-log -logFile $LogFile -Module $scriptName -Message "Error occurred while searching for products with keyword '$keyword': $_" -LogLevel "Error"
            $uninstallationCommands.hasErrors = $true
            $uninstallationCommands.message += "Error occurred while searching for products with keyword '$keyword': $_`n"
        }
    }

    Write-Verbose "[$functionName] Total products found before deduplication: $($allProducts.Count)"
    write-log -logFile $LogFile -Module $scriptName -Message "Total products found before deduplication: $($allProducts.Count)" -LogLevel "Information"

    # Deduplicate products using UninstallCmd and RegKey as unique identifiers
    # Handle null/empty UninstallCmd values properly
    $uniqueProducts = @{}
    $skippedDuplicates = 0

    foreach ($product in $allProducts) {
        # Create a unique key - use RegKey if UninstallCmd is null/empty
        $uniqueKey = if ([string]::IsNullOrWhiteSpace($product.UninstallCmd)) {
            if ([string]::IsNullOrWhiteSpace($product.RegKey)) {
                # If both are null, skip this product (shouldn't happen but handle gracefully)
                Write-Verbose "[$functionName] Skipping product with no UninstallCmd or RegKey: '$($product.Name)'"
                write-log -logFile $LogFile -Module $scriptName -Message "Skipping product with no UninstallCmd or RegKey: '$($product.Name)'" -LogLevel "Warning"
                continue
            }
            "RegKey:$($product.RegKey)"
        }
        else {
            "Uninstall:$($product.UninstallCmd)"
        }

        if (-not $uniqueProducts.ContainsKey($uniqueKey)) {
            $uniqueProducts[$uniqueKey] = $product
            Write-Verbose "[$functionName] Added unique product: '$($product.Name)' with key: $uniqueKey"
            write-log -logFile $LogFile -Module $scriptName -Message "Added unique product: '$($product.Name)' with key: $uniqueKey" -LogLevel "Verbose"
        }
        else {
            $skippedDuplicates++
            Write-Verbose "[$functionName] Skipped duplicate product: '$($product.Name)' (key already exists: $uniqueKey)"
            write-log -logFile $LogFile -Module $scriptName -Message "Skipped duplicate product: '$($product.Name)'" -LogLevel "Verbose"
        }
    }
    Write-Verbose "[$functionName] Removed $skippedDuplicates duplicate entries"
    write-log -logFile $LogFile -Module $scriptName -Message "Removed $skippedDuplicates duplicate entries. Unique products: $($uniqueProducts.Count)" -LogLevel "Information"
    # Convert to array and find the most likely candidate based on largest size
    $uniqueProductArray = @($uniqueProducts.Values)
    if ($uniqueProductArray.Count -gt 0) {
        # Find product with largest size (most likely the main application)
        $mostLikelyCandidate = $uniqueProductArray | Sort-Object -Property SizeKB -Descending | Select-Object -First 1
        if ($mostLikelyCandidate) {
            $mostLikelyCandidate.IsMostLikely = $true
            $uninstallationCommands.mostLikelyMatch = $mostLikelyCandidate
            Write-Verbose "[$functionName] Most likely candidate: '$($mostLikelyCandidate.Name)' (Size: $($mostLikelyCandidate.SizeMB)MB)"
            write-log -logFile $LogFile -Module $scriptName -Message "Most likely candidate identified: '$($mostLikelyCandidate.Name)' (Size: $($mostLikelyCandidate.SizeMB)MB, Version: $($mostLikelyCandidate.Version))" -LogLevel "Information"
        }

        # Sort products by size (largest first) for better organization
        $uniqueProductArray = $uniqueProductArray | Sort-Object -Property SizeKB -Descending
    }
    $uninstallationCommands.products = @($uniqueProductArray)
    Write-Verbose "[$functionName] Returning total of $($uninstallationCommands.products.count) unique products found."
    write-log -logFile $LogFile -Module $scriptName -Message "Returning total of $($uninstallationCommands.products.count) unique products. Most likely: '$($mostLikelyCandidate.Name)'" -LogLevel "Information"
    return $uninstallationCommands
}
