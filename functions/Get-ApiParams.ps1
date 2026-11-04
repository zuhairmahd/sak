function Get-ApiParams {
    param (
        [string]$accessToken,
        [string]$uri,
        [string]$Method,
        [string]$extraParameters,
        [string]$APIVersion,
        [string]$Search,
        [bool]$consistencyLevel = $false,
        [string]$filter,
        [string]$Body,
        [hashtable]$headers
    )

    if (-not $accessToken) { throw "Access token is required." }
    if (-not $uri) { throw "Resource URI is required." }       
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
    if ($consistencyLevel -eq $true) {
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

    return $apiParams
}                                                                       
