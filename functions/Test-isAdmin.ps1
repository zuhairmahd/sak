function Test-IsAdmin()
{
    <#
.SYNOPSIS
    Checks if the current PowerShell session is running with Administrator privileges.

.DESCRIPTION
    The Test-IsAdmin function determines whether the current user context has administrative rights.
    It logs the check and the current user information, then returns a boolean indicating admin status.

.EXAMPLE
    if (Test-IsAdmin) {
        Write-Host "Running as Administrator"
    } else {
        Write-Host "Not running as Administrator"
    }

.NOTES
    Author: MahmoudZ
    This function is intended for use in scripts that require elevated permissions.
    #>
    [CmdletBinding()]
    param()
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Checking if the script is running with admin privileges."
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    Write-Verbose "[$functionName] Current user: $($currentUser.Name)"
    Write-Verbose "[$functionName] Is current user an administrator? $($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))"
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}