[CmdletBinding()]
param(
    [string]$configFile = (Join-Path $PSScriptRoot ".secrets\config.json"),
    [string]$paramsFile = (Join-Path $PSScriptRoot "params.json"),
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
#region define configuration parameters
$params = @{
    configFile = $configFile
}
$auth = Get-Content -Path $paramsFile -Raw -Force -ErrorAction SilentlyContinue | ConvertFrom-Json
if ($null -ne $auth) {
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
}
#endregion define configuration parameters

$accessToken = Get-GraphAccessToken @params
$uri = "deviceAppManagement/mobileApps"
$extraParameters = "expand=assignments"
$filter = "(isof('microsoft.graph.windowsStoreApp') or isof('microsoft.graph.microsoftStoreForBusinessApp') or isof('microsoft.graph.officeSuiteApp') or isof('microsoft.graph.win32LobApp') or isof('microsoft.graph.windowsMicrosoftEdgeApp') or isof('microsoft.graph.windowsPhone81AppX') or isof('microsoft.graph.windowsPhone81StoreApp') or isof('microsoft.graph.windowsPhoneXAP') or isof('microsoft.graph.windowsAppX') or isof('microsoft.graph.windowsMobileMSI') or isof('microsoft.graph.windowsUniversalAppX') or isof('microsoft.graph.webApp') or isof('microsoft.graph.windowsWebApp') or isof('microsoft.graph.winGetApp'))&$orderby=displayName'"
$consistencyLevel = $true

#region define api parameters
$apiParams = @{
    accessToken  = $accessToken
    ResourcePath = $uri
}
if (-not [string]::IsNullOrWhiteSpace($Method)) {
    $apiParams.Method = $Method
}
if (-not [string]::IsNullOrWhiteSpace($extraParameters)) {
    $apiParams.extraParameters = $extraParameters
}
if (-not [string]::IsNullOrWhiteSpace($APIVersion)) {
    $apiParams.apiVersion = $APIVersion
}
if (-not [string]::IsNullOrWhiteSpace($Search)) {
    $apiParams.search = $Search
}
if ($consistencyLevel) {
    $apiParams.consistencyLevel = $consistencyLevel
}
if (-not [string]::IsNullOrWhiteSpace($filter)) {
    $apiParams.filter = $filter
}
if (-not [string]::IsNullOrWhiteSpace($Body)) {
    $apiParams.Body = $Body
}
if ($null -ne $headers) {
    $apiParams.headers = $headers
}
#endregion define api parameters

if ($accessToken) {
    Write-Host "Got access token"
    $global:response = Invoke-GraphAPI @apiParams
    if ($response.statusCode -in 200..299) {
        Write-Host "API call succeeded."
    }
    else {
        Write-Host "API call failed with status code $($response.statusCode)."
        #parce and print the properties of the $response.error object
        if ($response.error) {
            Write-Host "Error message: $($response.error.message)"
            foreach ($property in $response.error.PSObject.Properties) {
                Write-Host "$($property.Name): $($property.Value)"
            }
        }
    }
}
else {
    Write-Host "Failed to acquire access token."
}

