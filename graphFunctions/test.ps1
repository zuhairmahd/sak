[CmdletBinding()]
param(
    [string]$configFile = (Join-Path $PSScriptRoot ".secrets\config.json"),
    [int]$renewalLeadTime = 5,
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
    [string]$preferredBrowser = 'Default',
    [parameter(parameterSetName = 'delegated')]
    [switch]$privateSession,
    [ValidateSet('file', 'memory')]
    [string]$CacheType,
    [string]$APIVersion
)

. (Join-Path $PSScriptRoot "get-GraphAccessToken.ps1")
. (Join-Path $PSScriptRoot "Write-Log.ps1")
. (Join-Path $PSScriptRoot "Invoke-GraphAPI.ps1")

$script:logFile = Join-Path $PSScriptRoot "test.log"
$params = @{
    configFile = $configFile
}

if ($renewalLeadTime) { $params.renewalLeadTime = $renewalLeadTime }
if ($SecureString) { $params.SecureString = $SecureString }
if ($NoSaveRefreshToken) { $params.NoSaveRefreshToken = $NoSaveRefreshToken }
if ($delegated) { $params.delegated = $delegated }
if ($Scope) { $params.Scope = $Scope }
if ($AuthType) { $params.AuthType = $AuthType }
if ($ForceNewToken) { $params.ForceNewToken = $ForceNewToken }
if ($ForceNewRefreshToken) { $params.ForceNewRefreshToken = $ForceNewRefreshToken }
if ($preferredBrowser) { $params.preferredBrowser = $preferredBrowser }
if ($privateSession) { $params.privateSession = $privateSession }
if ($CacheType) { $params.CacheType = $CacheType }
if ($APIVersion) { $params.APIVersion = $APIVersion }


$accessToken = Get-GraphAccessToken @params
$uri = "users"
$global:users = Invoke-GraphAPI -accessToken $accessToken -ResourcePath $uri -method "GET"

