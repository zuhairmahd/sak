[CmdletBinding()]
param(
    [string]$configFile = (Join-Path $PSScriptRoot ".secrets\config.json"),
    [string]$paramsFile = (Join-Path $PSScriptRoot ".secrets\params.json"),
    [int]$renewalLeadTime,
    [switch]$SecureString,
    [parameter(parameterSetName = 'delegated')]
    [switch]$NoSaveRefreshToken,
    [parameter(parameterSetName = 'delegated')]
    [switch]$delegated,
    [parameter(parameterSetName = 'delegated')]
    [string[]]$Scope,
    [parameter(parameterSetName = 'delegated')]
    [ValidateSet('PublicAuthFlow', 'Interactive', 'Private')]
    [string]$AuthType,
    [parameter(parameterSetName = 'delegated')]
    [switch]$ForceNewToken,
    [parameter(parameterSetName = 'delegated')]
    [switch]$ForceNewRefreshToken,
    [parameter(parameterSetName = 'delegated')]
    [ValidateSet('Default', 'Edge', 'Chrome', 'Firefox')]
    [string]$preferredBrowser,
    [parameter(parameterSetName = 'delegated')]
    [switch]$privateSession,
    [ValidateSet('file', 'memory')]
    [string]$CacheType,
    [string]$APIVersion
)

. (Join-Path $PSScriptRoot "functions\get-GraphAccessToken.ps1")
. (Join-Path $PSScriptRoot "functions\Write-Log.ps1")
. (Join-Path $PSScriptRoot "functions\Invoke-GraphAPI.ps1")

$script:logFile = Join-Path $PSScriptRoot "test.log"
$params = @{
    configFile = $configFile
}

$auth = Get-Content -Path $paramsFile -Raw | ConvertFrom-Json
if ($renewalLeadTime) { $params.renewalLeadTime = $renewalLeadTime } elseif ($auth.renewalLeadTime) { $params.renewalLeadTime = $auth.renewalLeadTime }
if ($SecureString) { $params.SecureString = $SecureString } elseif ($auth.SecureString) { $params.SecureString = $auth.SecureString }
if ($NoSaveRefreshToken) { $params.NoSaveRefreshToken = $NoSaveRefreshToken } elseif ($auth.NoSaveRefreshToken) { $params.NoSaveRefreshToken = $auth.NoSaveRefreshToken }
if ($delegated) { $params.delegated = $delegated } elseif ($auth.delegated) { $params.delegated = $auth.delegated }
if ($Scope) { $params.Scope = $Scope } elseif ($auth.Scope) { $params.Scope = $auth.Scope }
if ($AuthType) { $params.AuthType = $AuthType } elseif ($auth.AuthType) { $params.AuthType = $auth.AuthType }
if ($ForceNewToken) { $params.ForceNewToken = $ForceNewToken } elseif ($auth.ForceNewToken) { $params.ForceNewToken = $auth.ForceNewToken }
if ($ForceNewRefreshToken) { $params.ForceNewRefreshToken = $ForceNewRefreshToken } elseif ($auth.ForceNewRefreshToken) { $params.ForceNewRefreshToken = $auth.ForceNewRefreshToken }
if ($preferredBrowser) { $params.preferredBrowser = $preferredBrowser } elseif ($auth.preferredBrowser) { $params.preferredBrowser = $auth.preferredBrowser }
if ($privateSession) { $params.privateSession = $privateSession } elseif ($auth.privateSession) { $params.privateSession = $auth.privateSession }
if ($CacheType) { $params.CacheType = $CacheType } elseif ($auth.CacheType) { $params.CacheType = $auth.CacheType }
if ($APIVersion) { $params.APIVersion = $APIVersion } elseif ($auth.APIVersion) { $params.APIVersion = $auth.APIVersion }

$accessToken = Get-GraphAccessToken @params
$uri = "users/me"
if ($accessToken) {
    $global:users = Invoke-GraphAPI -accessToken $accessToken -ResourcePath $uri -method "GET"
}
else {
    Write-Host "Failed to acquire access token."
}

