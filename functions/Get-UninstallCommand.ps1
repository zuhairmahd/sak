function Get-UninstallCommand() {
    <#
.SYNOPSIS
Retrieves uninstall command information for installed applications.

.DESCRIPTION
Searches the 32-bit, 64-bit, and per-user Windows uninstall registry locations under HKLM
and HKCU for installed products. Accepts either a structured product object or one or more
keyword strings. Deduplicates results, computes size information, and identifies the most
likely main product (largest EstimatedSize). Returns a summary object containing all matching
products and the mostLikelyMatch if requested.

Two search modes:
  - Product mode   ($product): uses productName as the keyword. Filters results against
                               productName, productVersion, and productPublisher according
                               to the -strictMatch switch.
  - Keyword mode   ($keywords): performs a wildcard contains-match ("*keyword*") against
                                DisplayName for each supplied keyword.

.PARAMETER product
A PSCustomObject with the following properties used for registry matching:
  - productName      [string] Required. Used as the search keyword.
  - productVersion   [string] Optional. Used for match filtering.
  - productPublisher [string] Optional. Used for match filtering.
When supplied, -keywords is ignored.

.PARAMETER keywords
One or more strings to match against DisplayName (case-insensitive wildcard contains-match).
Used only when -product is not provided. Returns hasErrors = $true when neither parameter
is supplied.

.PARAMETER GuessMostLikely
When set, the product with the largest EstimatedSize is marked as the most likely candidate

.PARAMETER strictMatch
Only applies in product mode (-product). When set, all three fields (DisplayName,
DisplayVersion, Publisher) must match exactly. When omitted, a product is included if at
least one of the three fields matches.

.OUTPUTS
PSCustomObject with the following properties:
  hasErrors       [bool]            True if input was invalid or an error occurred.
  message         [string]          Informational or error description.
  products        [PSCustomObject[]] All unique matching products, sorted by size descending.
                                    Each product exposes: Name, Version, UninstallCmd,
                                    QuietUninstall, SizeMB, SizeKB, InstallDate, RegKey,
                                    Publisher, InstallLocation, RegistryPath, IsMostLikely,
                                    and all available Bundle*/Engine*/URL* registry fields.
  mostLikelyMatch [PSCustomObject]  The product with the largest EstimatedSize (IsMostLikely = $true).

.EXAMPLE
# Keyword search — finds all products whose DisplayName contains "Chrome"
Get-UninstallCommand -keywords "Chrome"

.EXAMPLE
# Multi-keyword search
Get-UninstallCommand -keywords "Google", "Chrome"

.EXAMPLE
# Product-object search with loose matching
$p = [PSCustomObject]@{ productName = "Zoom"; productVersion = "6.0.0"; productPublisher = "Zoom Video Communications" }
Get-UninstallCommand -product $p

.EXAMPLE
# Product-object search requiring all three fields to match exactly
Get-UninstallCommand -product $p -strictMatch

.NOTES
Requires read access to:
  HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
  HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall
  HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
Depends on the write-log helper and expects $LogFile and $scriptName in the caller's scope.

.LINK
https://learn.microsoft.com/windows/win32/msi/uninstall-registry-key
    #>
    [CmdletBinding()]
    param(
        [PSCustomobject]$product,
        [string[]]$keywords,
        [switch]$GuessMostLikely,
        [switch]$strictMatch
    )

    $functionName = $MyInvocation.MyCommand.Name
    if ($product -and $null -ne $product.productName) {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for product with specific properties: $($product | Out-String)" -LogLevel "Information"
        $productSearch = $true
        if ($null -ne $keywords -or $keywords.count -eq 0) {
            $keywords = @($product.productName)
        }
    }
    elseif ($keywords -and $keywords.count -gt 0) {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for products with keywords: $($keywords -join ', ')" -LogLevel "Information"
        $productSearch = $false
    }
    else {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "No keywords or product properties provided. Returning empty list." -LogLevel "Warning"
        return @{
            hasErrors       = $true
            message         = "No keywords or product properties provided. Returning empty list."
            products        = @()
            mostLikelyMatch = $null
        }
    }
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
    if ($GuessMostLikely) {
        Write-Verbose "[$functionName] GuessMostLikely switch is set. Will identify the product with the largest size as the most likely candidate."
        write-log -logFile $LogFile -Module $scriptName -Message "GuessMostLikely switch is set. Will identify the product with the largest size as the most likely candidate." -LogLevel "Information"
        $uninstallationCommands.mostLikelyMatch = $null
    }
    else {
        Write-Verbose "[$functionName] GuessMostLikely switch is not set. Most likely candidate will not be identified."
        write-log -logFile $LogFile -Module $scriptName -Message "GuessMostLikely switch is not set. Most likely candidate will not be identified." -LogLevel "Information"
    }
    Write-Verbose "[$functionName] Get uninstallation commands for applications with keywords: $($keywords -join ', ')"
    write-log -logFile $LogFile -Module $scriptName -Message "Get uninstallation commands for applications with keywords: $($keywords -join ', ')" -LogLevel "Information"
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
                            Write-Verbose "[$functionName] Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion), Publisher: $($props.Publisher), Size: $($props.EstimatedSize) KB)"
                            write-log -logFile $LogFile -Module $scriptName -Message "Found matching product: '$($props.DisplayName)' (Version: $($props.DisplayVersion), Publisher: $($props.Publisher), Size: $($props.EstimatedSize) KB)" -LogLevel "Information"
                            if ($productSearch -and $strictMatch) {
                                if ($props.DisplayName -ne $product.productName -or $props.DisplayVersion -ne $product.productVersion -or $props.Publisher -ne $product.productPublisher) {
                                    Write-Verbose "[$functionName] Skipping product '$($props.DisplayName)' due to strict match criteria."
                                    write-log -logFile $LogFile -Module $scriptName -Message "Skipping product '$($props.DisplayName)' due to strict match criteria." -LogLevel "Verbose"
                                    continue
                                }
                            }
                            elseif ($productSearch -and -not $strictMatch) {
                                if ($props.DisplayName -ne $product.productName -and $props.DisplayVersion -ne $product.productVersion -and $props.Publisher -ne $product.productPublisher) {
                                    Write-Verbose "[$functionName] Skipping product '$($props.DisplayName)' due to non-strict match criteria."
                                    write-log -logFile $LogFile -Module $scriptName -Message "Skipping product '$($props.DisplayName)' due to non-strict match criteria." -LogLevel "Verbose"
                                    continue
                                }
                            }
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
        if ($GuessMostLikely) {
            Write-Verbose "[$functionName] GuessMostLikely switch is set. Identifying the product with the largest size as the most likely candidate."
            write-log -logFile $LogFile -Module $scriptName -Message "GuessMostLikely switch is set. Identifying the product with the largest size as the most likely candidate." -LogLevel "Information"
            # Find product with largest size (most likely the main application)
            $mostLikelyCandidate = $uniqueProductArray | Sort-Object -Property SizeKB -Descending | Select-Object -First 1
            if ($mostLikelyCandidate) {
                $mostLikelyCandidate.IsMostLikely = $true
                $uninstallationCommands.mostLikelyMatch = $mostLikelyCandidate
                Write-Verbose "[$functionName] Most likely candidate: '$($mostLikelyCandidate.Name)' (Size: $($mostLikelyCandidate.SizeMB)MB)"
                write-log -logFile $LogFile -Module $scriptName -Message "Most likely candidate identified: '$($mostLikelyCandidate.Name)' (Size: $($mostLikelyCandidate.SizeMB)MB, Version: $($mostLikelyCandidate.Version))" -LogLevel "Information"
                #remove it from the array so we don't have duplicates
                $uniqueProductArray = $uniqueProductArray | Where-Object { $_.RegKey -ne $mostLikelyCandidate.RegKey }
            }
        }
        else {
            Write-Verbose "[$functionName] GuessMostLikely switch is not set. Skipping identification of most likely candidate."
            write-log -logFile $LogFile -Module $scriptName -Message "GuessMostLikely switch is not set. Skipping identification of most likely candidate." -LogLevel "Information"
        }
        # Sort products by size (largest first) for better organization
        $uniqueProductArray = $uniqueProductArray | Sort-Object -Property SizeKB -Descending
    }
    $uninstallationCommands.products = @($uniqueProductArray)
    Write-Verbose "[$functionName] Returning total of $($uninstallationCommands.products.count) unique products found."
    write-log -logFile $LogFile -Module $scriptName -Message "Returning total of $($uninstallationCommands.products.count) unique products. Most likely: '$($mostLikelyCandidate.Name)'" -LogLevel "Information"
    return $uninstallationCommands
}
