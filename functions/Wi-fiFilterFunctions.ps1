# Requires Administrator privileges to run
# Check if running as Administrator
function Test-IsAdmin
{
    [CmdletBinding()]
    param()
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Message "Checking if script is running as Administrator" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    Write-Log -Message "Current user: $($currentUser.Name)" -Module $functionName -LogLevel "Debug" -LogFile $LogFile
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-WifiFilters
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$FilterSSIDs
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    Write-Host "Current Wi-Fi Filters:"
    Write-Log -Message "Retrieving Wi-Fi filters" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    
    $output = &netsh wlan show filters 2>&1
    
    if ($LASTEXITCODE -ne 0)
    {
        Write-Warning "Failed to retrieve Wi-Fi filters. Exit code: $LASTEXITCODE"
        Write-Log -Message "Failed to retrieve Wi-Fi filters. Exit code: $LASTEXITCODE" -Module $functionName -LogLevel "Error" -LogFile $LogFile
        return $null
    }
    
    Write-Host "Successfully retrieved Wi-Fi filters."
    Write-Log -Message "Successfully retrieved Wi-Fi filters" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    
    # Initialize result object
    $result = [PSCustomObject]@{
        GroupPolicyAllowList = @()
        UserAllowList        = @()
        GroupPolicyBlockList = @()
        UserBlockList        = @()
        RawOutput            = $output
    }
    
    # Parse the output
    $currentSection = $null
    $lines = $output -split "`r?`n"
    
    foreach ($line in $lines)
    {
        $line = $line.Trim()
        
        # Identify sections
        if ($line -match "Allow list on the system \(group policy\)")
        {
            $currentSection = "GroupPolicyAllowList"
            continue
        }
        elseif ($line -match "Allow list on the system \(user\)")
        {
            $currentSection = "UserAllowList"
            continue
        }
        elseif ($line -match "Block list on the system \(group policy\)")
        {
            $currentSection = "GroupPolicyBlockList"
            continue
        }
        elseif ($line -match "Block list on the system \(user\)")
        {
            $currentSection = "UserBlockList"
            continue
        }
        
        # Parse SSID entries
        if ($currentSection -and $line -match 'SSID: "([^"]+)", Type: (\w+)')
        {
            $ssid = $matches[1]
            $type = $matches[2]
            
            # Check if we should filter by specific SSIDs
            if ($FilterSSIDs -and $FilterSSIDs.Count -gt 0)
            {
                if ($ssid -notin $FilterSSIDs)
                {
                    continue
                }
            }
            
            $filterEntry = [PSCustomObject]@{
                SSID = $ssid
                Type = $type
            }
            
            switch ($currentSection)
            {
                "GroupPolicyAllowList"
                {
                    $result.GroupPolicyAllowList += $filterEntry 
                }
                "UserAllowList"
                {
                    $result.UserAllowList += $filterEntry 
                }
                "GroupPolicyBlockList"
                {
                    $result.GroupPolicyBlockList += $filterEntry 
                }
                "UserBlockList"
                {
                    $result.UserBlockList += $filterEntry 
                }
            }
            
            Write-Verbose "Found SSID '$ssid' of type '$type' in section '$currentSection'"
            Write-Log -Message "Found SSID '$ssid' of type '$type' in section '$currentSection'" -Module $functionName -LogLevel "Debug" -LogFile $LogFile
        }
    }
    
    # Log summary
    $totalFilters = $result.GroupPolicyAllowList.Count + $result.UserAllowList.Count + $result.GroupPolicyBlockList.Count + $result.UserBlockList.Count
    Write-Host "Parsed $totalFilters total filters from netsh output"
    Write-Log -Message "Parsed $totalFilters total filters: GP Allow: $($result.GroupPolicyAllowList.Count), User Allow: $($result.UserAllowList.Count), GP Block: $($result.GroupPolicyBlockList.Count), User Block: $($result.UserBlockList.Count)" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    
    if ($FilterSSIDs -and $FilterSSIDs.Count -gt 0)
    {
        Write-Host "Filtered results for SSIDs: $($FilterSSIDs -join ', ')"
        Write-Log -Message "Filtered results for SSIDs: $($FilterSSIDs -join ', ')" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    }
    
    return $result
}

function Add-WifiBlockFilter()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Ssids
    )

    $functionName = $MyInvocation.MyCommand.Name
    $wifi = @() 
    $wifiNetworks = [ordered] @{}
    Write-Host "Getting current Wi-Fi filters..."
    Write-Log -Message "Getting current Wi-Fi filters..." -Module $functionName -LogLevel "Information" -LogFile $LogFile
    $currentFilters = Show-WifiFilters -FilterSSIDs $networksToBlock
    if ($currentFilters -eq $null)
    {
        Write-Warning "Failed to retrieve current Wi-Fi filters.."
        Write-Log -Message "Failed to retrieve current Wi-Fi filters." -Module $functionName -LogLevel "Error" -LogFile $LogFile
    }
    else
    {
        Write-Host "Current Wi-Fi filters retrieved successfully."
        Write-Log -Message "Current Wi-Fi filters retrieved successfully." -Module $functionName -LogLevel "Information" -LogFile $LogFile
    }
    Write-Host "Adding block filter for $($Ssids.count) SSIDs"
    Write-Log -Message "Adding block filter for $($Ssids.count) SSIDs" -Module $functionName -LogLevel "Information" -LogFile $LogFile
    foreach ($Ssid in $Ssids)
    {
        Write-Host "Checking if $Ssid is already blocked..."
        Write-Log -Message "Checking if $Ssid is already blocked..." -Module $functionName -LogLevel "Information" -LogFile $LogFile
        if ($null -ne $currentFilters -and ($currentFilters.UserBlockList.SSID -contains $Ssid -or $currentFilters.GroupPolicyBlockList.SSID -contains $Ssid))
        {
            # If the SSID is already blocked, update the $wifiNetworks object and skip blocking
            Write-Host "$Ssid is already blocked. Skipping..."
            Write-Log -Message "$Ssid is already blocked. Skipping..." -Module $functionName -LogLevel "Information" -LogFile $LogFile
            $wifiNetworks = @{
                Ssid    = $Ssid        # Consistent key name
                Status  = "Already Blocked"
                Output  = "SSID is already blocked."
                Success = $true 
            }
            $wifi += $wifiNetworks
            continue
        }
        Write-Host "Blocking SSID: $Ssid"
        Write-Log -Message "Blocking SSID: $Ssid" -Module $functionName -LogLevel "Information" -LogFile $LogFile
        $output = &netsh wlan add filter permission=block ssid=`"$Ssid`" networktype=infrastructure 2>&1
        Write-Log -Message "Command output: $output" -Module $functionName -LogLevel "Debug" -logFile $LogFile
        Write-Host "Exit code: $LASTEXITCODE"
        Write-Log -Message "Exit code: $LASTEXITCODE" -Module $functionName -LogLevel "Debug" -logFile $LogFile
        if ($LASTEXITCODE -eq 0)
        {
            Write-Host "Successfully blocked $Ssid."
            Write-Log -Message "Successfully blocked $Ssid." -Module $functionName -LogLevel "Information" -logFile $logFile
            $wifiNetworks = @{
                Ssid    = $Ssid        # Consistent key name
                Status  = "Blocked"
                Output  = $output
                Success = $true 
            }
        }
        else
        {
            Write-Warning "Failed to block $Ssid. Exit code: $LASTEXITCODE"
            Write-Log -Message "Failed to block $Ssid. Exit code: $LASTEXITCODE" -Module $functionName -LogLevel "Error" -logFile $logFile
            $wifiNetworks = @{
                Ssid     = $Ssid        # Consistent key name
                Status   = "Failed"
                Output   = $output
                Success  = $false       
                ExitCode = $LASTEXITCODE  # Added exit code for debugging
            }
        }
        $wifi += $wifiNetworks
    }
    Write-Host "Block filter added for $($Ssids.count) SSIDs."
    Write-Log -Message "Block filter added for $($Ssids.count) SSIDs." -Module $functionName -LogLevel "Information" -logFile $logFile
    return $wifi
}

function Remove-WifiBlockFilter
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Ssids
    )
    
    Write-Host "Removing block filters for $($Ssids.count) SSIDs"
    $wifi = @() 
    $wifiNetworks = [ordered] @{}
    foreach ($Ssid in $Ssids)
    {
        Write-Host "Removing block filter for SSID: $Ssid"
        $output = &netsh wlan delete filter permission=block ssid=`"$Ssid`" networktype=infrastructure 2>&1
        Write-Verbose $output
        if ($LASTEXITCODE -eq 0)
        {
            Write-Host "Successfully unblocked $Ssid."
            $wifiNetworks = @{
                Ssid    = $Ssid        # Consistent key name
                Status  = "Unblocked"
                Output  = $output
                Success = $true 
            }
        }
        else
        {
            Write-Warning "Failed to unblock $Ssid. Exit code: $LASTEXITCODE"
            $wifiNetworks = @{
                Ssid     = $Ssid        # Consistent key name
                Status   = "Failed"
                Output   = $output
                Success  = $false       
                ExitCode = $LASTEXITCODE  # Added exit code for debugging
            }
        }
        $wifi += $wifiNetworks
    }
    return $wifi
}


# Example: Implement a denyall and then allow specific networks (whitelisting)
function Set-WifiWhitelist
{
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$AllowedSsids # Array of SSIDs to allow
    )
    Write-Host "Setting up Wi-Fi Whitelist. All other networks will be denied."

    # First, deny all networks
    Write-Host "Denying all infrastructure networks..."
    Invoke-Expression "netsh wlan add filter permission=denyall networktype=infrastructure"
    if ($LASTEXITCODE -ne 0)
    {
        Write-Warning "Failed to deny all networks."; return 
    }

    # Then, allow specified networks
    foreach ($ssid in $AllowedSsids)
    {
        Write-Host "Allowing SSID: $ssid"
        $command = "netsh wlan add filter permission=allow ssid=`"$ssid`" networktype=infrastructure"
        Invoke-Expression $command
        if ($LASTEXITCODE -ne 0)
        {
            Write-Warning "Failed to allow $ssid." 
        }
    }
    Write-Host "Whitelisting complete."
}

function Clear-WifiWhitelist
{
    Write-Host "Clearing Wi-Fi Whitelist (removing denyall filter)."
    Invoke-Expression "netsh wlan delete filter permission=denyall networktype=infrastructure"
    if ($LASTEXITCODE -eq 0)
    {
        Write-Host "Successfully cleared denyall filter. All networks are now visible."
    }
    else
    {
        Write-Warning "Failed to clear denyall filter. Exit code: $LASTEXITCODE"
    }
}

# Example: Use the whitelisting function
# Set-WifiWhitelist -AllowedSsids "MyHomeNetwork", "WorkWiFi", "MyPhoneHotspot"

# Example: Clear the whitelist
# Clear-WifiWhitelist


# --- How to use the functions in your script ---

# Example: Block networks and capture results
# $blockedNetworks = Add-WifiBlockFilter -Ssids "MyPublicWiFi", "CoffeeShopWiFi"

# Example: Get only successful blocks
# $successfulBlocks = $blockedNetworks | Where-Object { $_.Success -eq $true }

# Example: Block another network with spaces in its name
# $result = Add-WifiBlockFilter -Ssids "Starbucks Free WiFi"

# Example: Show current filters
$filters = Show-WifiFilters

# Example: Unblock a network
# Remove-WifiBlockFilter -Ssid "MyPublicWiFi"