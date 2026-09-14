function Get-UserProfilePaths {
    <#
    .SYNOPSIS
        Resolves the actual (possibly redirected) AppData/profile paths
        for the currently logged-in user, given their SID.
    .PARAMETER Sid
        The user's SID (reuse the SID resolved in Convert-HKCUPathToUserHive)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Sid
    )

    if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -Scope Script | Out-Null
    }

    $shellFoldersPath = "HKU:\$Sid\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

    if (-not (Test-Path $shellFoldersPath)) {
        throw "Could not find User Shell Folders key for SID '$Sid'."
    }

    $folders = Get-ItemProperty -Path $shellFoldersPath
    #Extract the userprofile variable from the first two folders of the list, which are AppData and Cache.  This is to ensure that the userprofile variable is expanded correctly.
    $userProfilePath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID").ProfileImagePath
    # Values may contain unexpanded vars like %USERPROFILE% - expand them
    [PSCustomObject]@{
        AppData      = [Environment]::ExpandEnvironmentVariables($folders.'AppData')
        Cache        = [Environment]::ExpandEnvironmentVariables($folders.'Cache')
        Cookies      = [Environment]::ExpandEnvironmentVariables($folders.'Cookies')
        Desktop      = [Environment]::ExpandEnvironmentVariables($folders.'Desktop')
        Documents    = [Environment]::ExpandEnvironmentVariables($folders.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}')
        Downloads    = [Environment]::ExpandEnvironmentVariables($folders.'{374DE290-123F-4565-9164-39C4925E467B}')
        Favorites    = [Environment]::ExpandEnvironmentVariables($folders.'Favorites')
        History      = [Environment]::ExpandEnvironmentVariables($folders.'History')
        LocalAppData = [Environment]::ExpandEnvironmentVariables($folders.'Local AppData')
        MyMusic      = [Environment]::ExpandEnvironmentVariables($folders.'My Music')
        MyPictures   = [Environment]::ExpandEnvironmentVariables($folders.'My Pictures')
        MyVideo      = [Environment]::ExpandEnvironmentVariables($folders.'My Video')
        NetHood      = [Environment]::ExpandEnvironmentVariables($folders.'NetHood')
        Personal     = [Environment]::ExpandEnvironmentVariables($folders.'Personal')
        Pictures     = [Environment]::ExpandEnvironmentVariables($folders.'{0DDD015D-B06C-45D5-8C4C-F59713854639}')
        PrintHood    = [Environment]::ExpandEnvironmentVariables($folders.'PrintHood')
        Programs     = [Environment]::ExpandEnvironmentVariables($folders.'Programs')
        Recent       = [Environment]::ExpandEnvironmentVariables($folders.'Recent')
        SendTo       = [Environment]::ExpandEnvironmentVariables($folders.'SendTo')
        StartMenu    = [Environment]::ExpandEnvironmentVariables($folders.'Start Menu')
        Startup      = [Environment]::ExpandEnvironmentVariables($folders.'Startup')
        Templates    = [Environment]::ExpandEnvironmentVariables($folders.'Templates')
        userProfile  = $userProfilePath
    }
}
