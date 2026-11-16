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
        [hashtable]$headers,
        [switch]$ValidateEndpoint
    )

    if (-not $accessToken) { throw "Access token is required." }
    if (-not $uri) { throw "Resource URI is required." }       
    if ($ValidateEndpoint) {
        $endpoints = Get-Content -Path (Join-Path $PSScriptRoot "ListOfMicrosoftGraphEndpoints.json") -Raw -Force -ErrorAction SilentlyContinue | ConvertFrom-Json
        $endpointInfo = $endpoints | Where-Object endpoint -EQ $uri
        if ($endpointInfo) {
            Write-Host "Endpoint found: $($endpointInfo.endpoint)"
            Write-Host "Available in v1.0: $($endpointInfo.'v1.0')"
            Write-Host "Available in beta: $($endpointInfo.'beta')"
        }
        else {
            Write-Host "Endpoint not found: $uri"
        }
    }
    
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
