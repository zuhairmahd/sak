function Test-IsServiceRunning()
{
    [CmdletBinding()]
    param (
        [string]$ServiceName
    )
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service)
    {
        return $service.Status -eq 'Running'
    }
    return $false
}
