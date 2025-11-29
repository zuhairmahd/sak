function Test-IsProductInstalled()
{
    <#
.SYNOPSIS
    Checks if a specified product is installed on the system using various identification methods.

.DESCRIPTION
    The Test-IsProductInstalled function determines whether a product is installed by checking for a file, registry entry, MSI product code, or MSI product name.
    Optionally, it can return the installed file version if the identifier type is 'File' and a version is requested.

.PARAMETER ProductName
    The display name of the product to check. Used for logging and output messages.

.PARAMETER ProductIdentifyer
    The identifier used to check for the product. This can be a file path, registry path, MSI product code (GUID), or MSI product name, depending on the identifier type.

.PARAMETER IdentifyerType
    The type of identifier to use for checking installation. Valid values are:
        - 'File': Checks if a file exists at the specified path.
        - 'Registry': Checks if a registry key or value exists at the specified path.
        - 'MSIProductCode': Checks for an installed MSI product by its product code (GUID).
        - 'MSIProductName': Checks for an installed MSI product by its display name.
    Default is 'File'.

.OUTPUTS
    [hashtable] with keys:
        - installed: [bool] $true if the product is installed, $false otherwise.
        - version: [string] The installed file version (only when using 'File' identifier type and FileVersion is specified), otherwise $null.

.EXAMPLE
    Test-IsProductInstalled -ProductName "App" -ProductIdentifyer "C:\Program Files\App\app.exe" -IdentifyerType "File"
    Checks if the file exists and returns its version.

.EXAMPLE
    Test-IsProductInstalled -ProductName "App" -ProductIdentifyer "HKLM:\Software\App" -IdentifyerType "Registry"
    Checks if the registry key exists.

.EXAMPLE
    Test-IsProductInstalled -ProductName "App" -ProductIdentifyer "{1234-5678-90AB-CDEF}" -IdentifyerType "MSIProductCode"
    Checks if the MSI product with the specified product code is installed.

.EXAMPLE
    Test-IsProductInstalled -ProductName "App" -ProductIdentifyer "App Name" -IdentifyerType "MSIProductName"
    Checks if the MSI product with the specified display name is installed.

.NOTES
    Returns a hashtable with installation status and (optionally) file version.
    Writes verbose and log messages for tracing.
    #>
    [CmdletBinding()]
    param (
        [string]$ProductName,
        [string]$ProductIdentifyer,
        [ValidateSet('File', 'Registry', 'MSIProductCode', 'MSIProductName')]
        [string]$IdentifyerType = 'File',
        [string]$serviceName
    )
    $functionName = $MyInvocation.MyCommand.Name
    #region write verbose logs of received parameters
    Write-Verbose "[$functionName] Product identifyer: $ProductIdentifyer"
    Write-Log -Message "Product identifyer: $ProductIdentifyer" -LogFile $logFile -Module $functionName -LogLevel "Verbose"
    Write-Verbose "[$functionName] Product Name: $ProductName"
    Write-Log -Message "Product Name: $ProductName" -LogFile $logFile -Module $functionName
    Write-Verbose "[$functionName] Identifyer Type: $IdentifyerType"
    Write-Log -Message "Identifyer Type: $IdentifyerType" -LogFile $logFile -Module $functionName -LogLevel "Verbose"
    Write-Verbose "[$functionName] Service: $serviceName"
    Write-Log -Message "Service: $serviceName" -LogFile $logFile -Module $functionName -LogLevel "Verbose"
    #endregion
    $returnObject = @{
        installed         = $false
        name              = $null
        version           = $null
        vender            = $null
        caption           = $null
        IdentifyingNumber = $null
    }    
    Write-Verbose "[$functionName] Starting installation check for $ProductName"                            
    Write-Verbose "[$functionName] Checking if $ProductName is already installed with Product identifyer: $ProductIdentifyer"
    Write-Log -Message "Checking if $ProductName is already installed with Product identifyer: $ProductIdentifyer" -LogFile $logFile -Module $functionName -LogLevel "Verbose"
    switch ($IdentifyerType)
    {
        'File'
        {
            Write-Verbose "[$functionName] Checking file at: $ProductIdentifyer"
            write-log -logFile $logFile -module $functionName -message "Checking file at: $ProductIdentifyer"
            if (Test-Path -Path $ProductIdentifyer)
            {
                $returnObject.version = (Get-Item -Path $ProductIdentifyer).VersionInfo.FileVersion
                $returnObject.installed = $true
                Write-Verbose "[$functionName] $ProductName is installed."
                Write-Log -Message "$ProductName is installed." -LogFile $logFile -Module $functionName
            }
            else
            {
                Write-Verbose "[$functionName] $ProductName is not installed."
                Write-Log -Message "$ProductName is not installed." -LogFile $logFile -Module $functionName
            }
        }
        'Registry'
        {
            Write-Verbose "[$functionName] Checking registry at: $ProductIdentifyer"
            write-log -logFile $logFile -module $functionName -message "Checking registry at: $ProductIdentifyer"
            if (Get-ItemProperty -Path $ProductIdentifyer -ErrorAction SilentlyContinue)
            {
                Write-Verbose "[$functionName] $ProductName is installed."
                Write-Log -Message "$ProductName is installed." -LogFile $logFile -Module $functionName
                $returnObject.installed = $true
            }
            else
            {
                Write-Verbose "[$functionName] $ProductName is not installed."
                Write-Log -Message "$ProductName is not installed." -LogFile $logFile -Module $functionName
            }
        }
        'MSIProductCode'
        {
            Write-Verbose "[$functionName] Checking WMI for product with MSI code: $ProductIdentifyer"
            write-log -logFile $logFile -module $functionName -message "Checking WMI for product with MSI code: $ProductIdentifyer"             
            $productInstalled = Get-WmiObject -Class Win32_Product | Where-Object { $_.IdentifyingNumber -eq $ProductIdentifyer }
            Write-Verbose "[$functionName] Product found: $($productInstalled |Out-String)"
            write-log -logFile $logFile -module $functionName -message "Product found: $($productInstalled |Out-String)"            
            if ($productInstalled)
            {
                Write-Verbose "[$functionName] $ProductName is installed."
                Write-Log -Message "$ProductName is installed." -LogFile $logFile -Module $functionName
                $returnObject.installed = $true
                $returnObject.IdentifyingNumber = $productInstalled.IdentifyingNumber
                $returnObject.name = $productInstalled.Name
                $returnObject.version = $productInstalled.Version
                $returnObject.vender = $productInstalled.Vendor 
                $returnObject.caption = $productInstalled.Caption
            }
            else
            {
                Write-Verbose "[$functionName] $ProductName is not installed."
                Write-Log -Message "$ProductName is not installed." -LogFile $logFile -Module $functionName
            }
        }
        'MSIProductName'
        {
            Write-Verbose "[$functionName] Checking WMI for product with name: $ProductIdentifyer"
            write-log -logFile -logFile -logFile $logFile -module $functionName -message "Checking WMI for product with name: $ProductIdentifyer"
            $productInstalled = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $ProductIdentifyer }
            Write-Verbose "[$functionName] Product found: $($productInstalled |Out-String)"
            write-log -logFile $logFile -module $functionName -message "Product found: $($productInstalled |Out-String)"        
            if ($productInstalled)  
            {
                Write-Verbose "[$functionName] $ProductName is installed."
                Write-Log -Message "$ProductName is installed." -LogFile $logFile -Module $functionName
                $returnObject.installed = $true
                $returnObject.IdentifyingNumber = $productInstalled.IdentifyingNumber
                $returnObject.name = $productInstalled.Name
                $returnObject.version = $productInstalled.Version
                $returnObject.vender = $productInstalled.Vendor 
                $returnObject.caption = $productInstalled.Caption
            }
            else
            {
                Write-Verbose "[$functionName] $ProductName is not installed."
                Write-Log -Message "$ProductName is not installed." -LogFile $logFile -Module $functionName
            }
        }
    }
    #if the service name is not a whitespace or empty string
    if ([string]::IsNullOrWhiteSpace($serviceName) -eq $false)
    {
        Write-Verbose "Checking whether the service $serviceName is installed"
        write-log -logFile $logFile -module $functionName -message "Checking whether the service $serviceName is installed"
        if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
        {
            Write-Verbose "[$functionName] Service: $serviceName is installed."
            Write-Log -Message "Service: $serviceName is installed." -LogFile $logFile -Module $functionName -LogLevel "Verbose"
        }
        else
        {
            Write-Verbose "[$functionName] Service: $serviceName is not installed."
            Write-Log -Message "Service: $serviceName is not installed." -LogFile $logFile -Module $functionName -LogLevel "Verbose"
            $returnObject.installed = $false
        }
    }
    else 
    {
        Write-Verbose "[$functionName] Service name is either null or empty."
    }
    return $returnObject
}