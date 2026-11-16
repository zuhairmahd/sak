[CmdletBinding()]
param(
    [string]$PolicyID
)

function Get-AccessToken() {
    [CmdletBinding()]
    param(
        [string]$TenantID,
        [string]$clientId,
        [string]$clientSecret,
        [string]$scopes = "https://graph.microsoft.com/.default"
    )

    $global:tokenCache = if ($null -eq $global:tokenCache) {
        @{}
    }
    else {
        $global:tokenCache
    }
    if ($global:tokenCache.access_token -and $global:tokenCache.expiry_time) {
        $currentTime = [DateTime]::UtcNow
        $expiryTime = [DateTime]$global:tokenCache.expiry_time
        $minutesRemaining = [math]::Round(($expiryTime - $currentTime).TotalMinutes, 1)
        # Consider token valid if it expires in more than 5 minutes
        if ($expiryTime -gt $currentTime.AddMinutes(5)) {
            Write-Verbose "[$functionName] Using cached access token (expires at $expiryTime UTC, $minutesRemaining minutes remaining)"
            Write-Host "Using cached access token ($minutesRemaining minutes remaining)"
            return $global:tokenCache.access_token
        }
        else {
            Write-Verbose "[$functionName] Cached access token expired or expiring soon, obtaining new token"
            Write-Host "[$functionName] Cached access token expired or expiring soon, obtaining new token"
        }
    }
    Write-Verbose "[$functionName] Getting access token for TenantID: $TenantID, ClientID: $clientId"
    try {
        $tokenEndpoint = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $body = @{
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = $scopes
            grant_type    = "client_credentials"
        }
        Write-Verbose "[$functionName] Requesting token from: $tokenEndpoint"
        $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
        if ($tokenResponse.access_token) {
            Write-Verbose "[$functionName] Successfully obtained access token (expires in $($tokenResponse.expires_in) seconds)"
            # Store the token and calculate absolute expiry time
            $global:tokenCache = @{
                access_token = $tokenResponse.access_token
                expiry_time  = ([DateTime]::UtcNow).AddSeconds($tokenResponse.expires_in)
            }

            return $tokenResponse.access_token
        }
        else {
            Write-Verbose "[$functionName] Token response did not contain an access token"
            throw "Token response did not contain an access token"
        }
    }
    catch {
        Write-Verbose "[$functionName] Token request failed: $_"
        if ($_.Exception.InnerException) {
            Write-Verbose "[$functionName] Inner Exception: $($_.Exception.InnerException.Message)"
        }
        throw
    }
}


exit 0
#Recursively get all .log files in the current folder and all subfolders
$logFiles = Get-ChildItem -Path . -Recurse -Filter *.log
Write-Host "Found $($logFiles.Count) log files"
#Search each file for the string "timeZone" and other variations "like time zone, timezone, tz"
$searchPatterns = @("timeZone", "time zone", "timezone", "Pacific", "eastern")
$logLines = @()

foreach ($file in $logFiles) {
    foreach ($pattern in $searchPatterns) {
        $string = $null
        $string = Select-String -Path $file.FullName -Pattern $pattern
        if ([string]::IsNullOrWhiteSpace($string)) {
            continue
        }
        $logObject = [PSCustomObject]@{
            FileName = (Resolve-Path $file.FullName).Path.Replace((Resolve-Path .).Path + "\", "")
            Pattern  = $pattern
            Line     = $string.Line
        }
        $logLines += $logObject
    }
}
#export the object to a csv
Write-Host "Exporting $($logLines.Count) log lines"
$logLines | Export-Csv -Path "logLines.csv" -NoTypeInformation -Force

$global:logs = $logLines

