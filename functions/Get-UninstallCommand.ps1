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
        [switch]$strictMatch,
        [switch]$All
    )

    $functionName = $MyInvocation.MyCommand.Name
    if ($product -and $null -ne $product.productName -and -not $all) {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for product with specific properties: $($product | Out-String)" -LogLevel "Information"
        $productSearch = $true
        if ($null -ne $keywords -or $keywords.count -eq 0) {
            $keywords = @($product.productName)
        }
    }
    elseif ($keywords -and $keywords.count -gt 0 -and -not $all) {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for products with keywords: $($keywords -join ', ')" -LogLevel "Information"
        $productSearch = $false
    }
    elseif ($all) {
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for all products." -LogLevel "Information"
        $productSearch = $false
        $keywords = @('All')
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
                        if ($keyword -and $props.DisplayName -and $props.DisplayName -like "*$keyword*" -or $all) {
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

                            $productObj = [PSCustomObject]@{}
                            foreach ($prop in $props.PSObject.Properties) {
                                $name = $prop.Name
                                # Unwrap complex objects; strip PS provider prefix from path strings
                                $value = if ($null -ne $prop.Value -and $prop.Value -isnot [string] -and $prop.Value -isnot [ValueType] -and $prop.Value.PSObject.Properties['Name']) {
                                    $prop.Value.Name
                                }
                                elseif ($prop.Value -is [string] -and $prop.Value -match '^\w[\w.]+\\(\w+)::(.+)$') {
                                    $Matches[2]
                                }
                                else {
                                    $prop.Value
                                }
                                Write-Verbose "[$functionName] Processing property: Name='$name', Value='$V alue'"
                                write-log -logFile $LogFile -Module $scriptName -Message "Processing property: Name='$name', Value='$Value'" -LogLevel "Verbose"
                                $productObj | Add-Member -MemberType NoteProperty -Name $name -Value $value -Force
                            }
                            #Add additional transformation properties
                            $UninstallCmd = if ($props.QuietUninstallString) { $props.QuietUninstallString } elseif ($props.UninstallString) { $props.UninstallString } else { $null }
                            $productObj | Add-Member -MemberType NoteProperty -Name UninstallCmd -Value $UninstallCmd -Force
                            $SizeMB = if ($props.EstimatedSize) { [math]::Round($props.EstimatedSize / 1024, 2) } else { 0 }
                            $productObj | Add-Member -MemberType NoteProperty -Name SizeMB -Value $SizeMB -Force
                            $SizeKB = if ($props.EstimatedSize) { $props.EstimatedSize } else { 0 }
                            $productObj | Add-Member -MemberType NoteProperty -Name SizeKB -Value $SizeKB -Force
                            $RegKey = $item.PSChildName # This is often the Product Code GUID
                            $productObj | Add-Member -MemberType NoteProperty -Name RegKey -Value $RegKey -Force
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

    $uninstallationCommands.products = @($allProducts)
    Write-Verbose "[$functionName] Returning total of $($uninstallationCommands.products.count) unique products found."
    write-log -logFile $LogFile -Module $scriptName -Message "Returning total of $($uninstallationCommands.products.count) unique products. Most likely: '$($mostLikelyCandidate.Name)'" -LogLevel "Information"
    return $uninstallationCommands
}
