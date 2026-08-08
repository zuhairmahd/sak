function Get-NetworkProfiles {
    <#
    .SYNOPSIS
    Retrieves all network profiles from the registry with an optional keyword filter.

    .DESCRIPTION
    This function retrieves all network profiles stored in the Windows registry under the specified path. It returns an array of objects containing the profile name and category. If a keyword is provided, only profiles matching the keyword will be returned.

    .OUTPUTS
    System.Object

    .EXAMPLE
    Get-NetworkProfiles

    Retrieves and displays all network profiles.

    .EXAMPLE
    Get-NetworkProfiles -keyword "Home"

    Retrieves and displays all network profiles matching the keyword "Home".
    #>
    param(
        [string]$keyword
    )

    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles"
    $profiles = [System.Collections.Generic.List[object]]::new()

    #get all subkeys
    $subKeys = Get-ChildItem -Path $regPath
    foreach ($subKey in $subKeys) {
        #get the profile name and category
        $props = Get-ItemProperty -Path $subKey.PSPath
        $profileName = $props.ProfileName
        $category = $props.Category
        switch ($category) {
            0 { $categoryName = "Public" }
            1 { $categoryName = "Private" }
            2 { $categoryName = "Domain" }
            default { $categoryName = "Unknown" }
        }
        #create a custom object for each profile
        $profileObject = [PSCustomObject]@{
            profileGuid = $subKey.PSChildName
            profilePath = $subKey.PSPath
            ProfileName = $profileName
            Category    = $categoryName
        }
        if (-not $keyword -or $profileName -like "*$keyword*") {
            [void]$profiles.Add($profileObject)
        }
    }
    return $profiles
}
