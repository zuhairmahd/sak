function Get-ProductInstallStatus()
{
    [CmdletBinding()]
    param(
        [string[]]$Products
    )

    $functionName = $MyInvocation.MyCommand.Name
    # Comprehensive list of uninstall registry keys (64-bit and 32-bit)
    $UninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    $result = @{
        HasErrors = $false
        Message   = ''
        Products  = @()
    }
    
    Write-Verbose "[$functionName] Checking installation status for applications with keywords: $($Products -join ', ')"     
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Checking installation status for applications with keywords: $($Products -join ', ')" -LogLevel "Information"
    
    if ($Products.Count -eq 0)
    {
        Write-Verbose "[$functionName] No keywords provided. Returning empty list." 
        Write-Log -LogFile $LogFile -Module $scriptName -Message "No keywords provided. Returning empty list." -LogLevel "Warning"
        $result.Message = "No keywords provided. Returning empty list."
        $result.HasErrors = $true
        return $result
    }
    
    Write-Verbose "[$functionName] Searching for $($Products.Count) keyword(s) across $($UninstallKeys.Count) registry locations."
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for $($Products.Count) keyword(s) across $($UninstallKeys.Count) registry locations." -LogLevel "Information"
    
    # Collect all matching products
    $allProducts = @()
    foreach ($keyword in $Products)
    {
        Write-Verbose "[$functionName] Searching for products with keyword: '$keyword'"
        Write-Log -LogFile $LogFile -Module $scriptName -Message "Searching for products with keyword: '$keyword'" -LogLevel "Information"
        
        # Check if keyword looks like a GUID (Product Code format)
        $isProductCode = $keyword -match '^(\{)?[0-9A-Fa-f]{8}[-]?([0-9A-Fa-f]{4}[-]?){3}[0-9A-Fa-f]{12}(\})?$'
        if ($isProductCode)
        {
            Write-Verbose "[$functionName] Keyword '$keyword' appears to be a Product Code (GUID)"
            Write-Log -LogFile $LogFile -Module $scriptName -Message "Keyword '$keyword' appears to be a Product Code (GUID)" -LogLevel "Information"
        }
        
        try
        {
            foreach ($key in $UninstallKeys)
            {
                Write-Verbose "[$functionName] Checking registry key: $key"
                Write-Log -LogFile $LogFile -Module $scriptName -Message "Checking registry key: $key" -LogLevel "Verbose"
                
                if (-not (Test-Path $key))
                {
                    Write-Verbose "[$functionName] Registry key does not exist: $key"
                    Write-Log -LogFile $LogFile -Module $scriptName -Message "Registry key does not exist: $key" -LogLevel "Verbose"
                    continue
                }
                
                $registryItems = Get-ChildItem $key -ErrorAction SilentlyContinue
                Write-Verbose "[$functionName] Found $($registryItems.Count) items in $key"
                Write-Log -LogFile $LogFile -Module $scriptName -Message "Found $($registryItems.Count) items in $key" -LogLevel "Verbose"
                
                foreach ($item in $registryItems)
                {
                    try
                    {
                        $props = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                        
                        # Check for match: either by DisplayName or by ProductCode
                        $matchFound = $false
                        $matchType = ""
                        
                        if ($isProductCode -and $item.PSChildName -like "*$keyword*")
                        {
                            $matchFound = $true
                            $matchType = "ProductCode"
                        }
                        elseif ($keyword -and $props.DisplayName -and $props.DisplayName -like "*$keyword*")
                        {
                            $matchFound = $true
                            $matchType = "DisplayName"
                        }
                        
                        if ($matchFound)
                        {
                            Write-Verbose "[$functionName] Found matching product by $matchType`: '$($props.DisplayName)' (Version: $($props.DisplayVersion))"
                            Write-Log -LogFile $LogFile -Module $scriptName -Message "Found matching product by $matchType`: '$($props.DisplayName)' (Version: $($props.DisplayVersion))" -LogLevel "Information"
                            
                            # Create product object with basic installation details
                            $productObj = [PSCustomObject]@{
                                Name            = $props.DisplayName
                                Version         = $props.DisplayVersion
                                ProductCode     = $item.PSChildName # This is often the Product Code GUID
                                Publisher       = $props.Publisher
                                InstallDate     = $props.InstallDate
                                InstallLocation = $props.InstallLocation
                                IsInstalled     = $true
                            }
                            
                            $allProducts += $productObj
                            Write-Verbose "[$functionName] Product details: Name='$($productObj.Name)', ProductCode='$($productObj.ProductCode)', Publisher='$($productObj.Publisher)'"
                            Write-Log -LogFile $LogFile -Module $scriptName -Message "Product details: ProductCode='$($productObj.ProductCode)', Publisher='$($productObj.Publisher)'" -LogLevel "Verbose"
                        }
                    }
                    catch
                    {
                        Write-Verbose "[$functionName] Error processing registry item $($item.PSPath): $_"
                        Write-Log -LogFile $LogFile -Module $scriptName -Message "Error processing registry item $($item.PSPath): $_" -LogLevel "Warning"
                    }
                }
            }
        }
        catch
        {
            Write-Error "[$functionName] Error occurred while searching for products with keyword '$keyword': $_"
            Write-Log -LogFile $LogFile -Module $scriptName -Message "Error occurred while searching for products with keyword '$keyword': $_" -LogLevel "Error"
            $result.HasErrors = $true
            $result.Message += "Error occurred while searching for products with keyword '$keyword': $_`n"
        }
    }
    
    Write-Verbose "[$functionName] Total products found before deduplication: $($allProducts.Count)"
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Total products found before deduplication: $($allProducts.Count)" -LogLevel "Information"
    
    # Deduplicate products using ProductCode as unique identifier
    $uniqueProducts = @{}
    $skippedDuplicates = 0
    
    foreach ($product in $allProducts)
    {
        # Use ProductCode as unique key
        if ([string]::IsNullOrWhiteSpace($product.ProductCode))
        {
            # If no ProductCode, use Name+Version as fallback
            $uniqueKey = "$($product.Name)|$($product.Version)"
        }
        else
        {
            $uniqueKey = $product.ProductCode
        }
        
        if (-not $uniqueProducts.ContainsKey($uniqueKey))
        {
            $uniqueProducts[$uniqueKey] = $product
            Write-Verbose "[$functionName] Added unique product: '$($product.Name)' with key: $uniqueKey"
            Write-Log -LogFile $LogFile -Module $scriptName -Message "Added unique product: '$($product.Name)' with key: $uniqueKey" -LogLevel "Verbose"
        }
        else
        {
            $skippedDuplicates++
            Write-Verbose "[$functionName] Skipped duplicate product: '$($product.Name)' (key already exists: $uniqueKey)"
            Write-Log -LogFile $LogFile -Module $scriptName -Message "Skipped duplicate product: '$($product.Name)'" -LogLevel "Verbose"
        }
    }
    
    Write-Verbose "[$functionName] Removed $skippedDuplicates duplicate entries"
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Removed $skippedDuplicates duplicate entries. Unique products: $($uniqueProducts.Count)" -LogLevel "Information"
    
    # Convert to array and sort by name
    [array]$uniqueProductArray = @($uniqueProducts.Values) | Sort-Object -Property Name
    $result.Products = $uniqueProductArray
    
    Write-Verbose "[$functionName] Returning total of $($result.Products.Count) unique installed products."        
    Write-Log -LogFile $LogFile -Module $scriptName -Message "Returning total of $($result.Products.Count) unique installed products." -LogLevel "Information"
    
    return $result
}       
