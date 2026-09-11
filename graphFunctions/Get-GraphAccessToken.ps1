function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$configFile,
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
        [string]$AuthType = 'Private',
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
        [string]$CacheType = 'Memory',
        [string]$APIVersion = 'Beta'
    )
    $maxJSONDepth = 10
    #region helper functions
    function FormatScopes {
        [CmdletBinding()]
        param(
            [string[]]$scopes,
            [switch]$Reverse
        )
        $functionName = $MyInvocation.MyCommand.Name
        $openIdScopes = @('offline_access', 'openid', 'profile')
        Write-Verbose "[$functionName] Called with Reverse=$Reverse"
        # Write-Verbose "[$functionName] Input scopes: '$scopes'"
        Write-Verbose "Passed parameter type: $($scopes.GetType().Name)"
        #Check for null or empty scopes.
        if (-not $scopes -or $scopes -eq "") {
            Write-Verbose "[$functionName] No scopes provided. Returning empty string."
            Write-Warning "[$functionName] WARNING: Scopes parameter is null or empty!"
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "No scopes provided - scopes parameter is null or empty" -LogLevel "Warning"
            return ""
        }

        #region Format scopes properly if necessary
        $scopesFormatted = $scopes
        Write-Verbose "[$functionName] Received $($scopes.count) scopes"
        if ($Reverse) {
            # Reverse mode: Remove Graph API prefixes and don't add default scopes
            Write-Verbose "[$functionName] Reverse mode: Removing Graph API prefixes"
            Write-Verbose "[$functionName] Converting scopes to array for processing"
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Converting scopes to array for processing" -LogLevel "Verbose"
            $scopesArray = $scopes.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
            $formattedScopesArray = @()
            Write-Verbose "[$functionName] Processing scope array with $($scopesArray.Count) items"
            foreach ($scope in $scopesArray) {
                Write-Verbose "[$functionName] Processing scope: $scope"
                if ($scope.StartsWith("https://graph.microsoft.com/")) {
                    Write-Verbose "[$functionName] Removing prefix from scope: $scope"
                    $formattedScope = $scope -replace "https://graph.microsoft.com/", ""
                    $formattedScopesArray += $formattedScope
                    Write-Verbose "[$functionName] Scope is now: $formattedScope"
                }
                else {
                    $formattedScopesArray += $scope
                    Write-Verbose "[$functionName] Added as is (no prefix): $scope"
                }
            }
            Write-Verbose "[$functionName] Count of scopes with prefixes removed: $($formattedScopesArray.count)"
            $scopesFormatted = $formattedScopesArray
        }
        else {
            Write-Verbose "[$functionName] Formatting scopes for normal mode"
            # Normal mode: Add Graph API prefixes and default scopes
            Write-Verbose "[$functionName] Normal mode: Adding Graph API prefixes"
            $scopesArray = $scopes.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
            $formattedScopesArray = @()
            Write-Verbose "[$functionName] Processing scope array with $($scopesArray.Count) items"
            foreach ($scope in $scopesArray) {
                Write-Verbose "[$functionName] Processing scope: $scope"
                if ($scope -in $openIdScopes) {
                    $formattedScopesArray += $scope
                    Write-Verbose "[$functionName] Added default scope as-is: $scope"
                }
                else {
                    # Check if the scope already has the Graph prefix
                    if (-not $scope.StartsWith("https://graph.microsoft.com/")) {
                        $formattedScopesArray += "https://graph.microsoft.com/$scope"
                        Write-Verbose "[$functionName] Added prefix: https://graph.microsoft.com/$scope"
                    }
                    else {
                        $formattedScopesArray += $scope
                        Write-Verbose "[$functionName] Added as-is (already has prefix): $scope"
                    }
                }
            }
            Write-Verbose "[$functionName] Count of scopes with prefixes added: $($formattedScopesArray.count)"
            Write-Verbose "[$functionName] Adding default scopes (openid and offline_access)"
            #If the $formattedScopesArray does not contain a scope in the $openIdScopes, add the missing scope.
            foreach ($defaultScope in $openIdScopes) {
                if ($formattedScopesArray -notcontains $defaultScope) {
                    Write-Verbose "[$functionName] Adding default scope: $defaultScope"
                    $formattedScopesArray += $defaultScope
                }
                else {
                    Write-Verbose "[$functionName] Default scope already present: $defaultScope"
                }
            }
            $scopesFormatted = $formattedScopesArray -join ' '
            # Remove any extra spaces
            $scopesFormatted = $scopesFormatted.Trim()
        }
        Write-Verbose "[$functionName] Final formatted scopes: $scopesFormatted"
        #endregion Format scopes
        return $scopesFormatted
    }

    function Start-HttpListener {
        [CmdletBinding()]
        param (
            [string]$redirectUri
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting HTTP listener function with redirectUri: $redirectUri"
        # Result object to return
        $result = @{
            Success      = $false
            Code         = $null
            ErrorMessage = $null
        }
        try {
            # Get local IP address for diagnostics
            try {
                $localIp = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" -and $_.InterfaceAlias -notlike "*Virtual*" }).IPAddress
                Write-Verbose "[$functionName] Local IP addresses: $($localIp -join ', ')"
            }
            catch {
                Write-Verbose "[$functionName] Could not get local IP address: $_"
            }
            # Create HTTP listener
            $listener = New-Object System.Net.HttpListener
            # Make sure redirect URI ends with a slash for matching
            $redirectUri = $redirectUri.TrimEnd('/') + '/'
            Write-Verbose "[$functionName] Using redirect URI: $redirectUri"
            # Add prefix
            $listener.Prefixes.Add($redirectUri)
            Write-Verbose "[$functionName] Added listener prefix: $redirectUri"
            # Start listener
            Write-Verbose "[$functionName] Starting HTTP listener..."
            $listener.Start()
            Write-Host "Waiting for authorization response. Please complete the sign in within your browser..."
            # Set timeout for listener
            $timeout = New-TimeSpan -Minutes 5
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            # Wait for the callback
            while ($stopwatch.Elapsed -lt $timeout) {
                # Check for pending request with a timeout
                if ($listener.IsListening) {
                    try {
                        # Try to get context with a 15-second timeout
                        $task = $listener.GetContextAsync()
                        $timeoutTask = [System.Threading.Tasks.Task]::Delay(15000)
                        $completedTask = [System.Threading.Tasks.Task]::WhenAny($task, $timeoutTask).GetAwaiter().GetResult()
                        # If the HTTP request came in before timeout
                        if ($completedTask -eq $task) {
                            $context = $task.GetAwaiter().GetResult()
                            Write-Verbose "[$functionName] Received HTTP request"
                            # Get request details for diagnostics
                            $request = $context.Request
                            Write-Verbose "[$functionName] Request URL: $($request.Url)"
                            Write-Verbose "[$functionName] Request headers: $($request.Headers)"
                            # Check if QueryString exists and has keys
                            if ($null -ne $request.QueryString -and $request.QueryString.Count -gt 0) {
                                Write-Verbose "[$functionName] Query string keys: $($request.QueryString.AllKeys -join ', ')"
                                # Check for code
                                $code = $request.QueryString.Get("code")
                                if ($null -ne $code) {
                                    Write-Verbose "[$functionName] Successfully retrieved authorization code"
                                    $result.Success = $true
                                    $result.Code = $code
                                }
                                else {
                                    # Check if there's an error
                                    $authError = $request.QueryString.Get("error")
                                    $errorDescription = $request.QueryString.Get("error_description")
                                    if ($authError) {
                                        Write-Verbose "[$functionName] Error in response: $authError - $errorDescription"
                                        $result.ErrorMessage = "Authentication error: $authError - $errorDescription"
                                    }
                                    else {
                                        Write-Verbose "[$functionName] No code or error found in query string"
                                        $result.ErrorMessage = "Authorization code not found in the response"
                                    }
                                }
                            }
                            else {
                                Write-Verbose "[$functionName] Query string is empty or null"
                                $result.ErrorMessage = "Empty query string in redirect response"
                            }
                            # Send response to browser
                            $response = $context.Response
                            $responseHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Authentication Complete</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #f0f0f0; }
        .container { background-color: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); padding: 40px; text-align: center; max-width: 600px; }
        h1 { color: #0078d4; margin-bottom: 20px; }
        p { color: #333; font-size: 16px; line-height: 1.6; }
        .status { font-weight: bold; margin: 20px 0; padding: 10px; border-radius: 4px; }
        .success { background-color: #dff6dd; color: #107c10; }
        .error { background-color: #fde7e9; color: #d13438; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Authentication Response</h1>
        $(if ($result.Success) {
            '<div class="status success">Authentication successful! You can close this window and return to the application.</div>'
        } else {
            "<div class='status error'>Authentication error: $($result.ErrorMessage)</div>"
        })
        <p>You may close this window and return to the PowerShell window.</p>
    </div>
</body>
</html>
"@
                            $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
                            $response.ContentLength64 = $buffer.Length
                            $response.ContentType = "text/html"
                            $response.StatusCode = 200
                            $response.OutputStream.Write($buffer, 0, $buffer.Length)
                            $response.Close()
                            # Break out of the loop
                            break
                        }
                    }
                    catch {
                        Write-Verbose "[$functionName] Error while waiting for request: $_"
                        Write-Verbose "[$functionName] Exception details: $($_.Exception.ToString())"
                        $result.ErrorMessage = "HTTP listener error: $($_.Exception.Message)"
                    }
                }
                else {
                    Write-Verbose "[$functionName] Listener is no longer listening"
                    $result.ErrorMessage = "HTTP listener stopped unexpectedly"
                    break
                }
                # Brief pause before checking again
                Start-Sleep -Milliseconds 100
            }
            # Check for timeout
            if ($stopwatch.Elapsed -ge $timeout) {
                Write-Verbose "[$functionName] Timeout waiting for authentication response"
                $result.ErrorMessage = "Timed out waiting for authentication response"
            }
        }
        catch {
            Write-Verbose "[$functionName] Error in HTTP listener: $_"
            Write-Verbose "[$functionName] Exception details: $($_.Exception.ToString())"
            $result.ErrorMessage = "Failed to start HTTP listener: $($_.Exception.Message)"
        }
        finally {
            # Clean up
            if ($listener -and $listener.IsListening) {
                $listener.Stop()
                $listener.Close()
                Write-Verbose "[$functionName] HTTP listener stopped"
            }
        }
        return $result
    }

    function DecodeJwtToken {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Token,
            [switch]$raw,
            [switch]$RawJSON
        )
        $functionName = $MyInvocation.MyCommand.Name
        $HRIdentifyers = @{
            aud                = 'Audience'
            iss                = 'Issuer'
            wids               = 'WindowsIdentifiers'
            roles              = 'Roles'
            sub                = 'Subject'
            oid                = 'Object Identifier'
            iat                = 'IssuedAt'
            nbf                = 'NotBefore'
            exp                = 'ExpirationTime'
            idp                = 'IdentityProvider'
            appidacr           = 'AuthenticationMechanism'
            idtyp              = 'IdentityType'
            tid                = 'TenantID'
            uti                = 'UniqueTokenIdentifier'
            ver                = 'Version'
            preferred_username = 'UserPrincipalName'
            email              = 'Email'
            upn                = 'UserPrincipalName'
            unique_name        = 'UniqueName'
            mail               = 'Email'
        }
        $parts = $Token -split '\.'
        Write-Verbose "[$functionName] Token parts: $($parts.Length)"
        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Processing JWT token with $($parts.Length) parts" -LogLevel "Verbose"
        if ($parts.Length -lt 2) {
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Invalid JWT token format. Expected at least 2 parts." -LogLevel "Error"
            return $returnValues.invalidJWTTokenMessage
        }
        $payload = $parts[1].Replace('-', '+').Replace('_', '/')
        Write-Verbose "[$functionName] Payload part: $($payload.Length)"
        switch ($payload.Length % 4) {
            2 {
                Write-Verbose "[$functionName] Adjusting payload length by adding two padding characters."; $payload += '=='
            }
            3 {
                Write-Verbose "[$functionName] Adjusting payload length by adding one padding character."; $payload += '='
            }
        }
        Write-Verbose "[$functionName] Adjusted payload for Base64 decoding: $($payload.Length)"
        $bytes = [System.Convert]::FromBase64String($payload)
        Write-Verbose "[$functionName] Decoded bytes length: $($bytes.Length)"
        if ($bytes.Length -eq 0) {
            Write-Verbose "[$functionName] Decoded bytes are empty. Invalid JWT payload."
            return $returnValues.invalidJWTPayloadMessage
        }
        Write-Verbose "[$functionName] Decoding payload to JSON string."
        $json = [System.Text.Encoding]::UTF8.GetString($bytes)
        $claims = $json | ConvertFrom-Json
        if (-not $raw) {
            # Convert all claims to human readable if possible
            Write-Verbose "[$functionName] Converting JWT claims to human readable format."
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Converting JWT claims to human readable format" -LogLevel "Debug"
            $humanClaims = [ordered]@{}
            foreach ($key in $claims.PSObject.Properties.Name) {
                Write-Verbose "[$functionName] Processing claim: $key"
                $value = $claims.$key
                Write-Verbose "[$functionName] Claim value: $value"
                Write-Verbose "[$functionName] Claim value type: $($value.GetType().Name)"
                # Check if the key is a known identifier and convert it to human readable format
                Write-Verbose "[$functionName] Checking if claim key $key is a known identifier."
                if ($HRIdentifyers.ContainsKey($key)) {
                    Write-Verbose "[$functionName] Claim key $key is a known identifier, converting to human readable format."
                    $key = $HRIdentifyers[$key]
                }
                else {
                    Write-Verbose "[$functionName] Claim key $key is not a known identifier, Skipping."
                    continue
                }
                # Try to convert unix time fields
                if ($value -is [int] -or $value -is [long]) {
                    # Heuristic: treat as unix time if key is a known time claim or value is in a reasonable unix time range
                    Write-Verbose "[$functionName] Claim is a number, checking if it is a unix time."
                    if ($value -gt 1000000000 -and $value -lt 3000000000) {
                        Write-Verbose "[$functionName] Claim $key is a unix time, converting to human readable format."
                        $dt = [DateTimeOffset]::FromUnixTimeSeconds([long]$value).UtcDateTime
                        $humanClaims[$key] = FormatDateWithTimeZone -DateTime $dt
                        continue
                    }
                }
                # Try to convert arrays to comma-separated string
                Write-Verbose "[$functionName] Checking if claim value is an array."
                if ($value -is [array]) {
                    Write-Verbose "[$functionName] Claim value is an array, converting to comma-separated string."
                    $humanClaims[$key] = $value -join ', '
                    continue
                }
                else {
                    Write-Verbose "[$functionName] Claim value is not an array."
                }
                if ($key -eq 'AuthenticationMechanism') {
                    switch ($value) {
                        0 {
                            $value = 'Public'
                            Write-Verbose "[$functionName] AuthenticationMechanism is 0, setting value to $value"
                        }
                        1 {
                            $value = 'App Secret'
                            Write-Verbose "[$functionName] AuthenticationMechanism is 1, setting value to $value"
                        }
                        2 {
                            $value = 'Certificate'
                            Write-Verbose "[$functionName] AuthenticationMechanism is 2, setting value to $value"
                        }
                        default {
                            $value = $value
                            Write-Verbose "[$functionName] AuthenticationMechanism is unknown, keeping value as $value"
                        }
                    }
                    $humanClaims[$key] = $value
                    continue
                }
                # Otherwise, just copy
                $humanClaims[$key] = $value
            }
            if ($RawJSON) {
                Write-Verbose "[$functionName] Returning all decoded JWT claims as raw JSON."
                return $humanClaims | ConvertTo-Json -Depth $maxJSONDepth
            }
            else {
                Write-Verbose "[$functionName] Returning all decoded JWT claims as object."
                return $humanClaims
            }
        }
        else {
            # Raw mode: Return unfiltered JWT claims
            Write-Verbose "[$functionName] Raw mode: Returning unfiltered JWT claims"
            if ($RawJSON) {
                Write-Verbose "[$functionName] Returning raw decoded JWT payload as raw JSON."
                return $json
            }
            else {
                Write-Verbose "[$functionName] Returning raw decoded JWT payload as object."
                return $claims
            }
        }
    }

    function Get-NormalizedExpiryTime {
        [CmdletBinding()]
        param(
            [object]$accessTokenObject
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting function execution"
        if (-not $accessTokenObject.AbsoluteExpiryTime) {
            Write-Verbose "[$functionName] No AbsoluteExpiryTime found in token object"
            Write-Log -logFile $logFile -moduleName $moduleName -logLevel Warning -message "No AbsoluteExpiryTime found in token object"
            return [datetime]::MinValue
        }

        try {
            if ($accessTokenObject.AbsoluteExpiryTime -is [string]) {
                Write-Verbose "[$functionName] Converting string expiry time to datetime"
                Write-Log -logFile $logFile -moduleName $moduleName -logLevel Verbose -message "Converting string expiry time to datetime"
                $parsedTime = [datetime]::Parse($accessTokenObject.AbsoluteExpiryTime).ToLocalTime()
                Write-Log -logFile $logFile -moduleName $moduleName -logLevel Verbose -message "Parsed expiry time: $parsedTime"
                # Handle timezone differences
                if ($parsedTime -lt $accessTokenObject.AbsoluteExpiryTime) {
                    Write-Verbose "[$functionName] Using original expiry time to resolve timezone differences"
                    return $accessTokenObject.AbsoluteExpiryTime
                }
                return $parsedTime
            }
            elseif ($accessTokenObject.AbsoluteExpiryTime.kind -eq 'Utc') {
                Write-Verbose "[$functionName] Converting UTC expiry time to local time"
                return $accessTokenObject.AbsoluteExpiryTime.ToLocalTime()
            }
            else {
                Write-Verbose "[$functionName] Using datetime expiry time as-is"
                return $accessTokenObject.AbsoluteExpiryTime
            }
        }
        catch {
            Write-Warning "[$functionName] Failed to parse expiry time: $_"
            Write-Log -logFile $logFile -moduleName $moduleName -logLevel Error -message "Failed to parse expiry time: $_"
            return [datetime]::MinValue
        }
    }

    function Get-TokenFromCache {
        [CmdletBinding()]
        param(
            [ValidateSet('file', 'memory')]
            [string]$cacheType,
            [string]$domain,
            [int]$renewalLeadTime,
            [string]$clientId,
            [string]$clientSecret,
            [string]$tenantId,
            [string[]]$scopes,
            [bool]$delegated,
            [string]$cacheFolder,
            [string]$cacheTokenFile,
            [bool]$secureString,
            [string]$configFilePath,
            [string]$configRefreshToken
        )

        function Get-CachedTokenObject {
            [CmdletBinding()]
            param(
                [string]$cacheType,
                [string]$cacheTokenFile,
                [string]$domain
            )

            $functionName = $MyInvocation.MyCommand.Name
            switch ($cacheType) {
                'memory' {
                    Write-Verbose "[$functionName] Checking memory cache for access token"
                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Checking memory cache for access token" -LogLevel "Verbose"
                    # Initialize memory cache if it doesn't exist
                    if (-not (Get-Variable -Name 'MemoryCache' -Scope Global -ErrorAction SilentlyContinue)) {
                        Write-Verbose "[$functionName] Initializing memory cache"
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Initializing memory cache" -LogLevel "Verbose"
                        New-Variable -Name 'MemoryCache' -Scope Global -Value @{} -Force
                    }

                    if ($Global:MemoryCache.ContainsKey('accessToken')) {
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Found token in memory cache"
                        Write-Verbose "[$functionName] Found token in memory cache"
                        $tokenObject = $Global:MemoryCache['accessToken']
                        if ($tokenObject.domain -eq $domain) {
                            Write-Verbose "[$functionName] Found matching token in memory cache for domain: $domain"
                            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Found matching token in memory cache for domain: $domain"
                            return $tokenObject
                        }
                        else {
                            Write-Verbose "[$functionName] Memory cache token domain ($($tokenObject.domain)) doesn't match requested domain ($domain)"
                            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Memory cache token domain ($($tokenObject.domain)) doesn't match requested domain ($domain)"
                        }
                    }
                    else {
                        Write-Verbose "[$functionName] No token found in memory cache"
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "No token found in memory cache"
                    }
                    return $null
                }
                'file' {
                    Write-Verbose "[$functionName] Checking file cache for access token: $cacheTokenFile"
                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Checking file cache for access token: $cacheTokenFile"
                    if (-not (Test-Path -Path $cacheTokenFile)) {
                        Write-Verbose "[$functionName] Cache file not found: $cacheTokenFile"
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Cache file not found: $cacheTokenFile" -LogLevel "Warning"
                        return $null
                    }

                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Reading token from file cache: $cacheTokenFile"
                    Write-Verbose "[$functionName] Reading token from file cache: $cacheTokenFile"

                    $tokenContent = Get-Content -Path $cacheTokenFile -Raw -Force
                    # Parse the token JSON
                    $tokenObject = ConvertFrom-Json $tokenContent
                    if ($tokenObject.domain -eq $domain) {
                        Write-Verbose "[$functionName] Found matching token in file cache for domain: $domain"
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Found matching token in file cache for domain: $domain"
                        return $tokenObject
                    }
                    else {
                        Write-Verbose "[$functionName] File cache token domain ($($tokenObject.domain)) doesn't match requested domain ($domain)"
                        Write-Log -LogFile $LogFile -Module "$functionName" -Message "File cache token domain ($($tokenObject.domain)) doesn't match requested domain ($domain)"
                    }

                    return $null
                }
                default {
                    Write-Error "[$functionName] Invalid cache type: $cacheType. Use 'file' or 'memory'."
                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Invalid cache type: $cacheType. Use 'file' or 'memory'." -LogLevel "Error"
                    return $null
                }
            }
        }

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting token cache retrieval for domain: $domain"
        Write-Verbose "[$functionName] Cache type: $cacheType, Delegated: $delegated"
        Write-Log -logFile $logFile -Module $functionName -Message "Starting token cache retrieval for domain: $domain, Cache type: $cacheType, Delegated: $delegated"
        # Calculate time buffer for token renewal
        $timeBuffer = (Get-Date).AddMinutes($renewalLeadTime)
        Write-Verbose "[$functionName] Token renewal buffer time: $timeBuffer"
        Write-Log -logFile $logFile -Module $functionName -Message "Token renewal buffer time: $timeBuffer"
        # Get cached token object based on cache type
        $accessTokenObject = Get-CachedTokenObject -cacheType $cacheType -cacheTokenFile $cacheTokenFile -domain $domain

        # If we have a cached token, validate and return it
        if ($accessTokenObject) {
            Write-Verbose "[$functionName] Cached token object found, validating token"
            Write-Log -logFile $logFile -Module $functionName -Message "Cached token object found, validating token"
            $validToken = Test-CachedTokenValidity -accessTokenObject $accessTokenObject -timeBuffer $timeBuffer -domain $domain -cacheType $cacheType -requestedScopes $scopes
            if ($validToken) {
                Write-Verbose "[$functionName] Valid cached token found, returning token"
                Write-Log -logFile $logFile -Module $functionName -Message "Valid cached token found, returning token"
                return $validToken
            }

            # Token is expired, try to refresh it
            $refreshedToken = Invoke-TokenRefresh -accessTokenObject $accessTokenObject -delegated $delegated -configRefreshToken $configRefreshToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder -configFilePath $configFilePath
            if ($refreshedToken) {
                Write-Verbose "[$functionName] Token refreshed successfully, returning refreshed token"
                Write-Log -logFile $logFile -Module $functionName -Message "Token refreshed successfully, returning refreshed token"
                return $refreshedToken
            }
        }

        # No cached token found, try config refresh token as fallback
        if ($delegated -and $configRefreshToken) {
            Write-Verbose "[$functionName] No valid cached token found, attempting to use config refresh token"
            Write-Log -logFile $logFile -Module $functionName -Message "No valid cached token found, attempting to use config refresh token"
            $refreshTokenObject = @{ refresh_token = $configRefreshToken }
            return Get-RefreshToken -accessTokenObject $refreshTokenObject -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder -configFilePath $configFilePath
        }

        Write-Verbose "[$functionName] No valid token found in cache or config"
        Write-Log -logFile $logFile -Module $functionName -Message "No valid token found in cache or config"
        return $null
    }

    function Test-CachedTokenValidity {
        [CmdletBinding()]
        param(
            [object]$accessTokenObject,
            [datetime]$timeBuffer,
            [string]$domain,
            [string]$cacheType,
            [string[]]$requestedScopes = @()
        )
        $functionName = $MyInvocation.MyCommand.Name
        if (-not $accessTokenObject.access_token) {
            Write-Verbose "[$functionName] No access token in cached object"
            Write-Log -logFile $logFile -Module "$functionName" -Message "No access token found in cached token object for $domain in $cacheType cache" -LogLevel "Warning"
            return $null
        }

        # Handle different time formats for expiry time
        $absoluteExpiryTime = Get-NormalizedExpiryTime -accessTokenObject $accessTokenObject
        Write-Log -logFile $logFile -Module "$functionName" -Message "Cached token expiry time for $domain in $cacheType cache: $absoluteExpiryTime"
        Write-Verbose "[$functionName] Normalized expiry time: $absoluteExpiryTime"
        if ($absoluteExpiryTime -gt $timeBuffer) {
            # Token is not expired, but check if scopes match (for delegated auth)
            Write-Log -logFile $logFile -Module "$functionName" -Message "Validating cached token scopes for $domain in $cacheType cache"
            Write-Verbose "[$functionName] Access token is not expired, validating scopes if requested"

            # Normalize requestedScopes to array first (may be passed as space-separated string or array)
            # BUGFIX: Filter empty elements created by multiple consecutive spaces
            # Note: Parameter is [string[]] so strings get wrapped in array automatically
            $requestedScopesArray = @()
            if ($requestedScopes.Count -eq 1 -and $requestedScopes[0] -match ' ') {
                # Single element array containing space-separated scopes (string was passed)
                Write-Verbose "[$functionName] Normalizing space-separated scope string to array - VERSION 2024-11-15-FINAL"
                $splitScopes = $requestedScopes[0] -split ' '
                foreach ($scope in $splitScopes) {
                    if (-not [string]::IsNullOrWhiteSpace($scope)) {
                        $requestedScopesArray += $scope.Trim()
                    }
                }
            }
            else {
                # Already an array or empty
                $requestedScopesArray = $requestedScopes
            }
            Write-Verbose "[$functionName] Requested scopes after normalization (count=$($requestedScopesArray.Count)): $($requestedScopesArray -join ', ')"

            if ($requestedScopesArray -and $requestedScopesArray.Count -gt 0) {
                Write-Verbose "[$functionName] Validating cached token has required scopes"
                Write-Log -logFile $logFile -Module "$functionName" -Message "Validating cached token has required scopes for $domain in $cacheType cache: $($requestedScopesArray -join ', ')"
                try {
                    # Decode the cached token to check its scopes
                    $decodedToken = DecodeJwtToken -Token $accessTokenObject.access_token -raw

                    # Extract granted scopes from scp claim (delegated) or roles claim (application)
                    $grantedScopes = @()
                    if ($decodedToken.scp) {
                        # Delegated auth - scp is space-separated string
                        Write-Verbose "[$functionName] Cached token has delegated scopes (scp)"
                        $grantedScopes = $decodedToken.scp -split ' ' | Where-Object { $_ -and $_.Trim() }
                        Write-Verbose "[$functionName] Cached token has delegated scopes (scp): $($grantedScopes -join ', ')"
                        Write-Log -logFile $logFile -Module "$functionName" -Message "Cached token has delegated scopes (scp): $($grantedScopes -join ', ')"
                    }
                    elseif ($decodedToken.roles) {
                        # Application auth - roles is array
                        $grantedScopes = $decodedToken.roles
                        Write-Verbose "[$functionName] Cached token has application scopes (roles): $($grantedScopes -join ', ')"
                        Write-Log -logFile $logFile -Module "$functionName" -Message "Cached token has application scopes (roles): $($grantedScopes -join ', ')"
                    }

                    # Check if all requested scopes are present in granted scopes
                    $missingScopes = $requestedScopesArray | Where-Object { $grantedScopes -notcontains $_ }
                    if ($missingScopes.Count -gt 0) {
                        Write-Verbose "[$functionName] Cached token is missing required scopes: $($missingScopes -join ', ')"
                        Write-Verbose "[$functionName] Token will be invalidated and re-acquired with updated scopes"
                        Write-Log -LogFile $logFile -Module "$functionName" -Message "Cached token missing scopes: $($missingScopes -join ', ') - invalidating cache" -LogLevel "Warning"
                        return $null
                    }
                    else {
                        Write-Verbose "[$functionName] Cached token has all required scopes"
                        Write-Log -logFile $logFile -Module "$functionName" -Message "Cached token has all required scopes for $domain in $cacheType cache"
                    }
                }
                catch {
                    Write-Verbose "[$functionName] Failed to decode cached token for scope validation: $($_.Exception.Message)"
                    Write-Log -LogFile $logFile -Module "$functionName" -Message "Failed to decode cached token: $($_.Exception.Message)" -LogLevel "Warning"
                    # Continue with token - expiry is still valid
                }
            }

            Write-Verbose "[$functionName] Access token for $domain is valid until $absoluteExpiryTime"
            Write-Verbose "[$functionName] Using cached access token from $cacheType cache"
            Write-Log -logFile $logFile -Module "$functionName" -Message "Using valid cached access token for $domain from $cacheType cache (expires: $absoluteExpiryTime)"
            Write-Host "Valid Token retrieved from cache." -ForegroundColor Green
            [console]::beep(200, 200)
            return $accessTokenObject.access_token
        }
        else {
            Write-Verbose "[$functionName] Access token for $domain is expired or invalid"
            Write-Verbose "[$functionName] Absolute expiry time: $absoluteExpiryTime, Time buffer: $timeBuffer"
            Write-Verbose "[$functionName] Will attempt to refresh token"
            Write-Log -LogFile $logFile -Module "$functionName" -Message "Access token in $cacheType cache is expired or invalid (expires: $absoluteExpiryTime, buffer: $timeBuffer)" -LogLevel "Warning"
            return $null
        }
    }

    function Save-TokenToCache {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$cachedToken,
            [Parameter(Mandatory = $true)]
            [ValidateSet('file', 'memory')]
            [string]$cacheType,
            [string]$cacheTokenFile,
            [string]$cacheFolder
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting token cache save operation"
        Write-Verbose "[$functionName] Cache type: $cacheType"

        # Extract token scope from JWT token (scp for delegated, roles for application)
        Write-Verbose "[$functionName] Decoding JWT token to extract scope information"
        $decodedToken = DecodeJwtToken -Token $cachedToken.access_token -raw

        # Determine scope based on auth type
        $tokenScope = @()
        if ($decodedToken.scp) {
            # Delegated auth - scp is space-separated string
            $tokenScope = $decodedToken.scp -split ' ' | Where-Object { $_ -and $_.Trim() }
            Write-Verbose "[$functionName] Extracted delegated scopes (scp): $($tokenScope -join ', ')"
        }
        elseif ($decodedToken.roles) {
            # Application auth - roles is array
            $tokenScope = $decodedToken.roles
            Write-Verbose "[$functionName] Extracted application scopes (roles): $($tokenScope -join ', ')"
        }
        else {
            Write-Warning "[$functionName] No scope information found in token (neither scp nor roles claims present)"
        }

        Write-Verbose "[$functionName] Successfully extracted token scope: $($tokenScope -join ', ')"

        # Add scope to cached token object
        Write-Verbose "[$functionName] Adding scope property to cached token object"
        if (-not $cachedToken.scope) {
            Write-Verbose "[$functionName] Scope property not found in cached token, adding it"
            $cachedToken.add('scope', $tokenScope)
        }
        else {
            Write-Verbose "[$functionName] Scope property already exists in cached token. Updating with extracted scopes."
            $cachedToken.scope = $tokenScope
        }

        # Save access token according to cache type
        if ($cacheType -eq 'memory') {
            Write-Verbose "[$functionName] Saving access token to memory cache"

            # Initialize global memory cache if it doesn't exist
            if (-not (Get-Variable -Name 'MemoryCache' -Scope Global -ErrorAction SilentlyContinue)) {
                Write-Verbose "[$functionName] Initializing global memory cache"
                New-Variable -Name 'MemoryCache' -Scope Global -Value @{} -Force
            }

            # Save to memory cache
            $Global:MemoryCache['accessToken'] = $cachedToken
            Write-Verbose "[$functionName] Token successfully saved to memory cache"

            # Debug: Log what was actually saved
            if ($Global:MemoryCache['accessToken'].scope) {
                Write-Verbose "[$functionName] Verified scope is accessible: $($Global:MemoryCache['accessToken'].scope -join ', ')"
            }
            else {
                Write-Warning "[$functionName] Scope property not found in saved token"
            }
        }
        else {
            Write-Verbose "[$functionName] Saving access token to cache file: $cacheTokenFile"

            if (-not (Test-Path -Path $cacheFolder)) {
                Write-Verbose "[$functionName] Creating cache folder: $cacheFolder"
                New-Item -Path $cacheFolder -ItemType Directory -Force | Out-Null
            }

            try {
                # Convert token to JSON
                $tokenJson = $cachedToken | ConvertTo-Json -Depth $maxJSONDepth

                # Check if user encryption password is available (same password used for config file)
                if ($script:UserEncryptionPassword -or $global:UserEncryptionPassword) {
                    $userPassword = if ($script:UserEncryptionPassword) { $script:UserEncryptionPassword } else { $global:UserEncryptionPassword }

                    Write-Verbose "[$functionName] Encrypting token before saving to file cache"
                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Encrypting token before saving to file cache" -LogLevel "Debug"

                    # Create a temporary file for encryption
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    try {
                        Set-Content -Path $tempFile -Value $tokenJson -Encoding UTF8 -NoNewline

                        # Encrypt the token using the user's password (same as config file)
                        # The -InMemoryOnly parameter causes Invoke-JsonFileEncryption to return the encrypted content in memory,
                        # rather than writing it back to the file. A temporary file is still needed because the encryption function
                        # expects a file input.
                        $encryptResult = Invoke-JsonFileEncryption -FilePath $tempFile -Key $userPassword -InMemoryOnly

                        if ($encryptResult.Success) {
                            # Save the encrypted content to the cache file
                            Set-Content -Path $cacheTokenFile -Value $encryptResult.Content -Force -ErrorAction Stop
                            Write-Verbose "[$functionName] Access token encrypted and saved successfully to $cacheTokenFile"
                            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Access token encrypted and saved successfully" -LogLevel "Information"
                        }
                        else {
                            Write-Warning "[$functionName] Failed to encrypt token, saving unencrypted: $($encryptResult.ErrorMessage)"
                            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Failed to encrypt token, saving unencrypted: $($encryptResult.ErrorMessage)" -LogLevel "Warning"
                            Set-Content -Path $cacheTokenFile -Value $tokenJson -Force -ErrorAction Stop
                        }
                    }
                    finally {
                        try {
                            # Overwrite the file with random data before deletion
                            if (Test-Path $tempFile) {
                                $fileInfo = Get-Item $tempFile
                                $fileLength = $fileInfo.Length
                                if ($fileLength -gt 0) {
                                    $randomBytes = New-Object byte[] $fileLength
                                    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($randomBytes)
                                    [System.IO.File]::WriteAllBytes($tempFile, $randomBytes)
                                }
                                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                            }
                            else {
                                Write-Verbose "[$functionName] Temporary file $($tempFile) does not exist, skipping secure deletion."
                            }
                        }
                        catch {
                            Write-Warning "[$functionName] Failed to securely delete temporary file $($tempFile): $_"
                        }
                        if (Test-Path $tempFile) {
                            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue | Out-Null
                        }
                    }
                }
                else {
                    Write-Verbose "[$functionName] No user encryption password available, saving token unencrypted"
                    Write-Log -LogFile $LogFile -Module "$functionName" -Message "No user encryption password available, saving token unencrypted" -LogLevel "Warning"
                    Set-Content -Path $cacheTokenFile -Value $tokenJson -Force -ErrorAction Stop
                }

                Write-Verbose "[$functionName] Access token successfully saved to $cacheTokenFile"
            }
            catch {
                Write-Error "[$functionName] Failed to save token to cache file: $_"
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Failed to save token to cache file: $_" -LogLevel "Error"
                throw
            }
        }
    }

    function Invoke-TokenRefresh {
        [CmdletBinding()]
        param(
            [object]$accessTokenObject,
            [bool]$delegated,
            [string]$configRefreshToken,
            [string]$clientId,
            [string]$clientSecret,
            [string]$tenantId,
            [string[]]$scopes,
            [string]$domain,
            [string]$cacheType,
            [string]$cacheTokenFile,
            [string]$cacheFolder,
            [string]$configFilePath
        )
        $functionName = $MyInvocation.MyCommand.Name
        if (-not $delegated) {
            Write-Verbose "[$functionName] Not using delegated authentication, skipping refresh token logic"
            return $null
        }

        # Try cached refresh token first
        if ($accessTokenObject.refresh_token) {
            Write-Verbose "[$functionName] Attempting to refresh token using cached refresh token"
            return Get-RefreshToken -accessTokenObject $accessTokenObject -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder -configFilePath $configFilePath
        }

        # Try config refresh token as fallback
        if ($configRefreshToken) {
            Write-Verbose "[$functionName] Attempting to refresh token using config refresh token"
            $refreshTokenObject = @{ refresh_token = $configRefreshToken }
            return Get-RefreshToken -accessTokenObject $refreshTokenObject -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder -configFilePath $configFilePath
        }

        Write-Verbose "[$functionName] No refresh token available for token refresh"
        return $null
    }

    function Get-TokenFromResponse {
        [CmdletBinding()]
        param($tokenResponse, $domain, $refreshToken)
        $functionName = $MyInvocation.MyCommand.Name
        $tokenExpiryTime = (Get-Date).AddSeconds($tokenResponse.expires_in)
        Write-Verbose "[$functionName] Token absolute expiry time: $($tokenExpiryTime)"
        $cachedToken = [ordered] @{
            'domain'           = $domain
            access_token       = $tokenResponse.access_token
            AbsoluteExpiryTime = $tokenExpiryTime
            'expires_in'       = $tokenResponse.expires_in
        }

        # Add refresh token if available
        if ($refreshToken -or $tokenResponse.refresh_token) {
            $cachedToken.Add('refresh_token', $tokenResponse.refresh_token)
        }

        # Add scope if available
        if ($tokenResponse.scope) {
            $cachedToken.Add('scope', $tokenResponse.scope)
        }

        return $cachedToken
    }

    function Format-TokenOutput {
        [CmdletBinding()]
        param($token, $secureString)
        $functionName = $MyInvocation.MyCommand.Name
        if ($secureString) {
            Write-Verbose "[$functionName] Converting access token to secure string"
            $secureAccessToken = ConvertTo-SecureString -String $token -AsPlainText -Force
            Write-Verbose "[$functionName] Returning secure token"
            return $secureAccessToken
        }
        else {
            Write-Verbose "[$functionName] Returning plain text access token"
            return $token
        }
    }

    function Test-RefreshTokenValidity {
        [CmdletBinding()]
        param(
            $refreshToken,
            $clientId,
            $clientSecret,
            $tenantId,
            $scopes,
            $domain,
            $AuthType
        )
        $functionName = $MyInvocation.MyCommand.Name
        #write verbose and log the function name and parameters
        Write-Verbose "[$functionName] Testing refresh token validity with parameters:clientId=$clientId, tenantId=$tenantId, domain=$domain, AuthType=$AuthType"
        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Testing refresh token validity with parameters: refreshToken=$refreshToken, clientId=$clientId, tenantId=$tenantId, scopes=$scopes, domain=$domain, AuthType=$AuthType"
        try {
            # Ensure scopes are properly formatted for the token refresh
            $scopesFormatted = $scopes
            if ($scopes -and -not $scopes.Contains("https://graph.microsoft.com/")) {
                Write-Verbose "[$functionName] Scopes provided, formatting them for token refresh"
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Scopes provided, formatting them for token refresh" -LogLevel "Verbose"
                $scopesArray = $scopes.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
                $formattedScopesArray = @()
                foreach ($scope in $scopesArray) {
                    if ($scope -eq "offline_access") {
                        $formattedScopesArray += $scope
                    }
                    else {
                        $formattedScopesArray += "https://graph.microsoft.com/$scope"
                    }
                }
                $scopesFormatted = $formattedScopesArray -join ' '
            }
            $refreshTokenRequestBody = @{
                client_id     = $clientId
                refresh_token = $refreshToken
                grant_type    = 'refresh_token'
                scope         = $scopesFormatted
            }
            if ($auth.AuthType -eq 'PublicAuthFlow') {
                Write-Verbose "[$functionName] Public authentication flow detected, client secret will not be included in the request"
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Public authentication flow detected, client secret will not be included in the request" -LogLevel "Verbose"
            }
            else {
                Write-Verbose "[$functionName] Non-public authentication flow detected, client secret will be included in the request"
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Non-public authentication flow detected, client secret will be included in the request" -LogLevel "Verbose"
                $refreshTokenRequestBody.Add('client_secret', $clientSecret)
            }
            $refreshTokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
            Write-Verbose "[$functionName] Attempting to validate refresh token by getting a new access token.."
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $refreshTokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $refreshTokenRequestBody
            Write-Verbose "[$functionName] Refresh token is valid. Successfully obtained a new access token."
            return $true, $tokenResponse
        }
        catch {
            Write-Verbose "[$functionName] Refresh token validation failed: $_"
            Write-Verbose "[$functionName] Refresh token appears to be invalid or expired."
            return $false, $null
        }
    }

    function Save-RefreshTokenToConfig {
        [CmdletBinding()]
        param($refreshToken, $configFilePath)

        $functionName = $MyInvocation.MyCommand.Name
        $delegatedCredentials = @{}

        Write-Log -LogFile $LogFile -Module $functionName -Message "Starting refresh token save operation" -LogLevel "Verbose"
        Write-Verbose "[$functionName] Called with configFilePath=$configFilePath, refreshToken provided=$($null -ne $refreshToken)"

        if (-not (Test-Path -Path $configFilePath)) {
            Write-Verbose "[$functionName] Config file does not exist at path $configFilePath"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Config file does not exist at path $configFilePath" -LogLevel "Warning"
            return
        }

        if (-not $refreshToken) {
            Write-Verbose "[$functionName] No refresh token to save"
            Write-Log -LogFile $LogFile -Module $functionName -Message "No refresh token provided - operation skipped" -LogLevel "Warning"
            return
        }

        $config = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json

        # Format scopes
        if ($refreshToken.scope) {
            Write-Verbose "[$functionName] Adding scope to new delegatedCredentials"
            $formattedScopes = FormatScopes -scopes $refreshToken.scope -Reverse
            Write-Verbose "[$functionName] Formatted scopes object type: $($formattedScopes.GetType().Name)"
            #If it is not an array, make it int an array
            if (-not ($formattedScopes -is [array])) {
                Write-Verbose "[$functionName] Scopes is not an array, converting to array"
                $formattedScopes = @($formattedScopes)
            }
        }
        else {
            Write-Verbose "[$functionName] No scopes provided in refresh token, using empty array"
            $formattedScopes = @()
        }

        # Update the config with the new refresh token
        if ($config.delegatedCredentials) {
            Write-Verbose "[$functionName] Updating existing delegatedCredentials property"
            $config.delegatedCredentials.refresh_token = $refreshToken.refresh_token
            $config.delegatedCredentials.scope = $formattedScopes
        }
        else {
            Write-Verbose "[$functionName] Creating new delegatedCredentials property"
            $delegatedCredentials.refresh_token = $refreshToken.refresh_token
            $delegatedCredentials.scope = $formattedScopes
            $config | Add-Member -MemberType NoteProperty -Name 'delegatedCredentials' -Value $delegatedCredentials
        }

        Write-Verbose "[$functionName] Saving refresh token to config file: $configFilePath"
        $config | ConvertTo-Json -Depth $maxJSONDepth | Set-Content -Path $configFilePath -Force

        Write-Verbose "[$functionName] Refresh token saved to config file"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Refresh token operation completed successfully" -LogLevel "Information"
    }

    function Get-RefreshToken {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [object]$accessTokenObject,
            [Parameter(Mandatory = $true)]
            [string]$clientId,
            [string]$clientSecret,
            [Parameter(Mandatory = $true)]
            [string]$tenantId,
            [string[]]$scopes,
            [Parameter(Mandatory = $true)]
            [string]$domain,
            [ValidateSet('file', 'memory')]
            [string]$cacheType,
            [string]$cacheTokenFile,
            [string]$cacheFolder,
            [string]$configFilePath
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Attempting to use refresh token to get a new access token" -LogLevel "Verbose"
        try {
            # First test if the refresh token is valid
            $isValid, $tokenResponse = Test-RefreshTokenValidity -refreshToken $accessTokenObject.refresh_token -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain
            if (-not $isValid) {
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Refresh token is invalid or expired. Cannot proceed with token refresh." -LogLevel "Warning"
                return $null
            }
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Refresh token is valid. Proceeding to get new access token." -LogLevel "Information"
            $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain -refreshToken $tokenResponse.refresh_token
            # Cache the access token based on cache type
            Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
            # Only save the refresh token if it's different from the one we already have
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Checking whether to save the refresh token..." -LogLevel "Verbose"
            if ($tokenResponse.refresh_token -and $tokenResponse.refresh_token -ne $accessTokenObject.refresh_token) {
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "Saving new refresh token as it differs from the existing one." -LogLevel "Verbose"
                Save-RefreshTokenToConfig -refreshToken $tokenResponse -configFilePath $configFilePath
            }
            else {
                Write-Log -LogFile $LogFile -Module "$functionName" -Message "No need to save refresh token as it hasn't changed." -LogLevel "Verbose"
            }
            return Format-TokenOutput -token $tokenResponse.access_token -secureString $SecureString
        }
        catch {
            Write-Host "Failed to use refresh token: $_. Will request new authorization."
            Write-Host "Refresh token might be expired or revoked, proceeding with new authorization."
            Write-Log -LogFile $LogFile -Module "$functionName" -Message "Failed to use refresh token: $_. Will request new authorization." -LogLevel "Error"
            return $null
        }
    }

    function LaunchBrowser {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [ValidateNotNullOrEmpty()]
            [string]$url,
            [ValidateSet("Chrome", "Edge", "Firefox", "Default")]
            [string]$browser
        )

        $functionName = $MyInvocation.MyCommand.Name
        #print log of incoming parameters.
        Write-Verbose "[$functionName] Launching browser with URL: $url"
        Write-Verbose "[$functionName] Browser preference: $browser"
        Write-Verbose "[$functionName] Private session: $private"
        if ($null -eq $browser -or $browser -eq '') {
            Write-Verbose "[$functionName] No preferred browser set in settings, using default browser"
            $browser = 'Default'
        }
        switch ($Browser) {
            'Edge' {
                Write-Verbose "[$functionName] Opening Edge browser for authentication"
                if ($privateSession) {
                    Write-Verbose "[$functionName] Private session detected.  Opening $browser in private mode"
                    $urlParams = @{
                        FilePath     = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                        ArgumentList = "--inprivate", $url
                    }
                }
                else {
                    $urlParams = @{
                        FilePath     = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                        ArgumentList = $url
                    }
                }
            }
            'Chrome' {
                Write-Verbose "[$functionName] Opening Chrome browser for authentication"
                if ($privateSession) {
                    Write-Verbose "[$functionName] Private session detected.  Opening $preferredBrowser in private mode"
                    $urlParams = @{
                        FilePath     = "C:\Program Files\Google\Chrome\Application\chrome.exe"
                        ArgumentList = "--incognito", $url
                    }
                }
                else {
                    $urlParams = @{
                        FilePath     = "C:\Program Files\Google\Chrome\Application\chrome.exe"
                        ArgumentList = $url
                    }
                }
            }
            'Firefox' {
                Write-Verbose "[$functionName] Opening Firefox browser for authentication"
                if ($privateSession) {
                    Write-Verbose "[$functionName] Private session detected.  Opening $preferredBrowser  in private mode"
                    $urlParams = @{
                        FilePath     = "C:\Program Files\Mozilla Firefox\firefox.exe"
                        ArgumentList = "-private-window", $url
                    }
                }
                else {
                    $urlParams = @{
                        FilePath     = "C:\Program Files\Mozilla Firefox\firefox.exe"
                        ArgumentList = $url
                    }
                }
            }
            default {
                Write-Verbose "[$functionName] Opening default browser for authentication"
                $urlParams = @{
                    FilePath = $url
                }
            }
        }
        Write-Verbose "[$functionName] Launching $browser with URL: $url"
        try {
            # Start the browser with the specified URL
            Start-Process @urlParams
            Write-Verbose "[$functionName] Browser launched successfully."
        }
        catch {
            Write-Error "Failed to launch browser: $_"
            return $false
        }
        return $true
    }

    function Get-DelegatedToken {
        [CmdletBinding()]
        param(
            [string]$tenantId,
            [string]$clientId,
            [string]$clientSecret,
            [string[]]$scopes,
            [string]$domain,
            [string]$cacheType,
            [string]$cacheTokenFile,
            [string]$cacheFolder,
            [string]$configFilePath,
            [string]$configRefreshToken,
            [string]$AuthType,
            [switch]$NoSaveRefreshToken,
            [switch]$ForcedRenewal
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting Get-DelegatedToken function."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Received parameters: tenantId=$tenantId, clientId=$clientId, scopes=$scopes, domain=$domain, cacheType=$cacheType, cacheTokenFile=$cacheTokenFile, cacheFolder=$cacheFolder, configFilePath=$configFilePath, configRefreshToken=$configRefreshToken, AuthType=$AuthType, NoSaveRefreshToken=$NoSaveRefreshToken, ForcedRenewal=$ForcedRenewal"
        Write-Verbose "[$functionName] Received parameters: tenantId=$tenantId, clientId=$clientId, scopes=$scopes, domain=$domain, cacheType=$cacheType, cacheTokenFile=$cacheTokenFile, cacheFolder=$cacheFolder, configFilePath=$configFilePath, configRefreshToken=$configRefreshToken, AuthType=$AuthType, NoSaveRefreshToken=$NoSaveRefreshToken, ForcedRenewal=$ForcedRenewal"
        if ($clientSecret) {
            Write-Log -LogFile $LogFile -Module $functionName -Message "Client secret was provided."
            Write-Verbose "[$functionName] Client secret was provided."
        }
        elseif ($certificateThumbprint) {
            Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate thumbprint was provided."
            Write-Verbose "[$functionName] Certificate thumbprint was provided."
        }
        # Handle null or empty scopes by providing a sensible default
        if (-not $scopes -or $scopes -eq "") {
            Write-Verbose "[$functionName] Scopes parameter is null or empty. Using default Graph API scopes."
            Write-Log -LogFile $LogFile -Module $functionName -Message "Scopes parameter is null or empty. Using default Graph API scopes." -LogLevel Warning
            $scopes = "offline_access openid Device.ReadWrite.All DeviceManagementApps.Read.All DeviceManagementConfiguration.ReadWrite.All DeviceManagementManagedDevices.PrivilegedOperations.All DeviceManagementManagedDevices.ReadWrite.All DeviceManagementServiceConfig.ReadWrite.All"
            Write-Host "Using default scopes as none were provided: $scopes" -ForegroundColor Yellow
        }

        Write-Verbose "[$functionName] Starting with scopes: '$scopes'"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Starting with scopes: '$scopes'"

        # First check if we have a valid refresh token in config
        if ($configRefreshToken) {
            Write-Verbose "[$functionName] Found refresh token in config. Testing its validity before requesting new authorization..."
            Write-Log -LogFile $LogFile -Module $functionName -Message "Found refresh token in config. Testing its validity before requesting new authorization..."
            $isValid, $tokenResponse = Test-RefreshTokenValidity -refreshToken $configRefreshToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $scopes -domain $domain -AuthType $AuthType
            if ($isValid) {
                Write-Host "Existing refresh token is valid. Using it without requesting a new authorization."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Existing refresh token is valid. Using it without requesting a new authorization."
                $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain -refreshToken $configRefreshToken
                # Cache the access token
                Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
                # We don't need to save the refresh token again since it's the same one
                Write-Verbose "[$functionName] Using existing refresh token that is still valid."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using existing refresh token that is still valid."
                return Format-TokenOutput -token $tokenResponse.access_token -secureString $SecureString
            }
            else {
                Write-Verbose "[$functionName] Existing refresh token is invalid. Will proceed with new authorization."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Existing refresh token is invalid. Will proceed with new authorization."
                if ($ForcedRenewal) {
                    Write-Host "Forcing new refresh token - proceeding with new authentication flow." -ForegroundColor Yellow
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Forcing new refresh token - proceeding with new authentication flow." -LogLevel Warning
                }
            }
        }
        else {
            Write-Verbose "[$functionName] No existing refresh token found in config."
            Write-Log -LogFile $LogFile -Module $functionName -Message "No existing refresh token found in config."
            if ($ForcedRenewal) {
                Write-Host "No existing refresh token found or refresh token was cleared - proceeding with new authentication flow." -ForegroundColor Yellow
                Write-Log -LogFile $LogFile -Module $functionName -Message "No existing refresh token found or refresh token was cleared - proceeding with new authentication flow."
            }
            else {
                Write-Verbose "[$functionName] Proceeding with new authentication flow."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Proceeding with new authentication flow."
            }
        }

        # Generate a random state string
        Write-Verbose "[$functionName] Generating random state string."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Generating random state string."
        $state = [System.Guid]::NewGuid().ToString()
        $scopesFormatted = FormatScopes -scopes $scopes
        $encodedScopes = [uri]::EscapeDataString($scopesFormatted)

        $automaticFlowSuccess = $false
        switch ($AuthType) {
            PublicAuthFlow {
                Write-Verbose "[$functionName] Using device auth flow."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using device auth flow."
                $deviceCodeRequestUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode"
                $deviceCodeRequestBody = @{
                    client_id = $clientId
                    scope     = $scopesFormatted
                }
                Write-Verbose "[$functionName] Device code request URL: $deviceCodeRequestUrl"
                Write-Verbose "[$functionName] Original scopes parameter: '$scopes'"
                Write-Verbose "[$functionName] Formatted scopes: '$scopesFormatted'"
                Write-Verbose "[$functionName] Requesting device code with body: $($deviceCodeRequestBody | ConvertTo-Json -Depth $maxJSONDepth)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Device code request URL: $deviceCodeRequestUrl"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Original scopes parameter: '$scopes'"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Requesting device code with body: $($deviceCodeRequestBody | ConvertTo-Json -Depth $maxJSONDepth)"
                try {
                    $deviceCodeResponse = Invoke-RestMethod -Method POST -Uri $deviceCodeRequestUrl -Body $deviceCodeRequestBody
                    Write-Verbose "[$functionName] Device code response: $($deviceCodeResponse | ConvertTo-Json -Depth $maxJSONDepth)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Device code response: $($deviceCodeResponse | ConvertTo-Json -Depth $maxJSONDepth)"
                }
                catch {
                    Write-Error "Error requesting device code: $($_.Exception.Message)"
                    Write-Error "Response: $($_.Exception.Response.GetResponseStream() | ForEach-Object { New-Object System.IO.StreamReader($_) } | ForEach-Object { $_.ReadToEnd() })"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Error requesting device code: $($_.Exception.Message)" -LogLevel Error
                    return $null
                }
                Write-Host ""
                Write-Host $deviceCodeResponse.message
                #extract the code and copy it to the clipboard.
                $regex = "(?<=enter the code )([A-Z0-9]+)"
                if ($deviceCodeResponse.message -match $regex) {
                    $code = $matches[1]
                    Write-Verbose "[$functionName] Extracted code from device code response: $code"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Extracted code from device code response: $code"
                    #now copy the code to the clipboard.
                    Set-Clipboard -Value $code
                    Write-Host "The code has been copied to the clipboard for you to paste."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "The code has been copied to the clipboard for you to paste."
                }
                else {
                    Write-Verbose "[$functionName] No code found in device code response message."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "No code found in device code response message."
                }
                if ([string]::IsNullOrWhiteSpace($preferredBrowser)) {
                    Write-Verbose "[$functionName] No preferred browser set in settings, using default browser."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "No preferred browser set in settings, using default browser."
                    $preferredBrowser = 'Default'
                    $displayMessage = "If you choose 'Yes', your default browser will be used for authentication."
                }
                else {
                    Write-Verbose "[$functionName] Using preferred browser : $preferredBrowser"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Using preferred browser: $preferredBrowser  "
                    $displayMessage = "If you choose 'Yes', $preferredBrowser will be used for authentication."
                    if ($privateSession) {
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Using private session for browser."
                        $displayMessage += " A private session (incognito) will be used."
                    }
                }
                Write-Verbose "[$functionName] Preferred browser for authentication: $preferredBrowser"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Preferred browser for authentication: $preferredBrowser"
                $authUrl = "https://microsoft.com/devicelogin"
                Write-Verbose "[$functionName] Authentication URL: $authUrl"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Authentication URL: $authUrl"
                Write-Host "Would you like to open a browser to the authentication page?"
                Write-Host "`n$displayMessage"
                $userChoice = Read-Host "Type 'Yes' to open the browser, or 'No' to continue without opening a browser `n (you will need to manually open your browser to $authUrl)"
                while ($userChoice -notin @('Yes', 'No')) {
                    Write-Host "Invalid choice. Please type 'Yes' or 'No'."
                    #beep
                    [console]::beep(1000, 500)
                    $userChoice = Read-Host "Type 'Yes' to open the browser, or 'No' to continue without opening a browser"
                }
                if ($userChoice -eq 'Yes') {
                    Write-Verbose "[$functionName] User chose to open browser for authentication."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "User chose to open browser for authentication."
                    $browserOpened = LaunchBrowser -url $authUrl -browser $preferredBrowser
                    if (-not $browserOpened) {
                        Write-Error "Failed to open browser for authentication. Please open it manually."
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to open browser for authentication."
                    }
                }
                else {
                    Write-Verbose "[$functionName] User chose not to open browser, will continue with manual authentication."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "User chose not to open browser, will continue with manual authentication."
                    Write-Host "Please open your browser to $authUrl and paste the code: $code to sign-in"
                }
                Write-Host "Waiting for authentication..."
                Write-Host ""
                # --- Poll for Access Token ---
                Write-Verbose "[$functionName] Polling for access token using device code."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Polling for access token using device code."
                $tokenRequestUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
                $tokenRequestBody = @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $clientId
                    device_code = $deviceCodeResponse.device_code
                }
                $accessToken = $null
                $timeoutSeconds = $deviceCodeResponse.expires_in # Typically 15 minutes
                Write-Verbose "[$functionName] Token request URL: $tokenRequestUrl"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Token request URL: $tokenRequestUrl"
                Write-Verbose "[$functionName] Token request body: $($tokenRequestBody | ConvertTo-Json -Depth $maxJSONDepth)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Token request body: $($tokenRequestBody | ConvertTo-Json -Depth $maxJSONDepth)"
                Write-Verbose "Timeout for polling: $timeoutSeconds seconds"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Timeout for polling: $timeoutSeconds seconds"
                $intervalSeconds = $deviceCodeResponse.interval # Typically 5 seconds
                Write-Verbose "Polling interval: $intervalSeconds seconds"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Polling interval: $intervalSeconds seconds"
                $startTime = Get-Date
                Write-Verbose "[$functionName] Start time for polling: $startTime"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Start time for polling: $startTime"
                while ((Get-Date -UFormat %s) -lt ($startTime.AddSeconds($timeoutSeconds) | Get-Date -UFormat %s)) {
                    Write-Verbose "[$functionName] Polling for access token..."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Polling for access token..."
                    Start-Sleep -Seconds $intervalSeconds
                    try {
                        $tokenResponse = Invoke-RestMethod -Method POST -Uri $tokenRequestUrl -Body $tokenRequestBody -ErrorAction SilentlyContinue
                        Write-Verbose "[$functionName] Polling attempt successful."
                        Write-Verbose "[$functionName] Token response: $($tokenResponse | ConvertTo-Json -Depth $maxJSONDepth)"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Polling attempt successful. Token response: $($tokenResponse | ConvertTo-Json -Depth $maxJSONDepth)"
                        if ($tokenResponse.access_token) {
                            $accessToken = $tokenResponse.access_token
                            Write-Host "Authentication successful. Access token acquired."
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Authentication successful. Access token acquired."
                            $automaticFlowSuccess = $true
                            break
                        }
                        elseif ($tokenResponse.error -ne "authorization_pending") {
                            Write-Error "Error polling for token: $($tokenResponse.error_description)"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Error polling for token: $($tokenResponse.error_description)"
                            return $null
                        }
                        else {
                            Write-Verbose "[$functionName] Authorization still pending, continuing to poll..."
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Authorization still pending, continuing to poll..."
                        }
                    }
                    catch {
                        # Check if this is the expected "authorization_pending" error (400 Bad Request)
                        $isAuthPending = $false

                        # Check the HTTP status code first
                        if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 400) {
                            Write-Verbose "[$functionName] Received 400 Bad Request during polling - checking if authorization is pending..."
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Received 400 Bad Request during polling - checking if authorization is pending..."
                            # Multiple ways to detect authorization_pending:
                            # 1. Check the exception message for common patterns
                            $exceptionMessage = $_.Exception.Message
                            if ($exceptionMessage -like "*authorization_pending*" -or
                                $exceptionMessage -like "*Bad Request*" -or
                                $exceptionMessage -like "*400*") {
                                $isAuthPending = $true
                                Write-Verbose "[$functionName] Detected authorization_pending from exception message pattern"
                                Write-Log -LogFile $LogFile -Module $functionName -Message "Detected authorization_pending from exception message pattern"
                            }

                            # 2. Try to parse the response body if available
                            if (-not $isAuthPending) {
                                try {
                                    $errorResponse = $_.Exception.Response.GetResponseStream()
                                    if ($errorResponse -and $errorResponse.CanRead) {
                                        $streamReader = New-Object System.IO.StreamReader($errorResponse)
                                        $errorMessage = $streamReader.ReadToEnd()
                                        $streamReader.Close()

                                        if ($errorMessage) {
                                            $errorJson = $errorMessage | ConvertFrom-Json
                                            if ($errorJson.error -eq "authorization_pending") {
                                                $isAuthPending = $true
                                                Write-Verbose "[$functionName] Confirmed authorization_pending from response body"
                                                Write-Log -LogFile $LogFile -Module $functionName -Message "Confirmed authorization_pending from response body"
                                            }
                                        }
                                    }
                                }
                                catch {
                                    Write-Verbose "[$functionName] Could not parse error response, but assuming authorization_pending for 400 status"
                                    Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse error response, assuming authorization_pending for 400 status"
                                    # For 400 errors during OAuth device flow polling, assume it's authorization_pending
                                    $isAuthPending = $true
                                }
                            }

                            if ($isAuthPending) {
                                Write-Verbose "[$functionName] Authorization still pending (from catch block), continuing to poll..."
                                Write-Log -LogFile $LogFile -Module $functionName -Message "Authorization still pending (from catch block), continuing to poll..."
                            }
                        }

                        # Only show warning for unexpected errors, not for authorization_pending
                        if (-not $isAuthPending) {
                            Write-Warning "Polling attempt failed: $($_.Exception.Message)"
                            Write-Verbose "[$functionName] Unexpected error during polling: $($_.Exception | Out-String)"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Unexpected error during polling: $($_.Exception.Message)"
                        }
                    }
                    Write-Host -NoNewline "."
                }
                if (-not $accessToken) {
                    Write-Error "Authentication timed out or failed."
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Authentication timed out or failed."
                    return $null
                }
            }
            'interactive' {
                Write-Verbose "[$functionName] Using interactive authentication flow."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using interactive authentication flow."
                $redirectUri = "http://localhost:8080/"
                Write-Verbose "[$functionName] Redirect URI: $redirectUri"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Redirect URI: $redirectUri"
                $encodedRedirectUri = [uri]::EscapeDataString($redirectUri)
                Write-Verbose "[$functionName] Encoded Redirect URI: $encodedRedirectUri"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Encoded Redirect URI: $encodedRedirectUri"
                Write-Verbose "[$functionName] Attempting automatic HTTP listener flow"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Attempting automatic HTTP listener flow"
                try {
                    Write-Verbose "[$functionName] Starting HTTP listener at $redirectUri"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Starting HTTP listener at $redirectUri"
                    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?client_id=$clientId&response_type=code&redirect_uri=$encodedRedirectUri&response_mode=query&scope=$encodedScopes&state=$state"
                    Write-Verbose "[$functionName] Authorization URL: $authUrl"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Authorization URL: $authUrl"
                    Write-Host "Opening browser for user authentication and consent..."
                    if (-not (LaunchBrowser -url $authUrl -browser $preferredBrowser)) {
                        Write-Error "Failed to launch browser. Please open the URL manually: $authUrl"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to launch browser for authentication" -LogLevel Error
                        return $null
                    }
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Browser launched successfully for authentication"
                    $listenerResult = Start-HttpListener -redirectUri $redirectUri
                    if ($listenerResult.Success) {
                        Write-Verbose "[$functionName] HTTP listener successfully captured the authorization code"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP listener successfully captured the authorization code"
                        $code = $listenerResult.Code
                        $automaticFlowSuccess = $true
                    }
                    else {
                        Write-Warning "HTTP listener failed to capture the authorization code: $($listenerResult.ErrorMessage)"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP listener failed to capture the authorization code: $($listenerResult.ErrorMessage)" -LogLevel Warning
                        Write-Verbose "[$functionName] Will fall back to manual code input"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Will fall back to manual code input"
                        $automaticFlowSuccess = $false
                    }
                }
                catch {
                    Write-Warning "Error in automatic HTTP listener flow: $_"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Error in automatic HTTP listener flow: $_" -LogLevel Warning
                    Write-Verbose "[$functionName] Will fall back to manual code input"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Will fall back to manual code input due to exception"
                    $automaticFlowSuccess = $false
                }
            }
            'Private' {
                Write-Verbose "[$functionName] Using non-interactive mode (manual code input)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using non-interactive mode (manual code input)"
                $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
                Write-Verbose "[$functionName] Redirect URI: $redirectUri"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Redirect URI: $redirectUri"
                $encodedRedirectUri = [uri]::EscapeDataString($redirectUri)
                Write-Verbose "[$functionName] Encoded Redirect URI: $encodedRedirectUri"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Encoded Redirect URI: $encodedRedirectUri"
            }
            default {
                Write-Error "Invalid AuthType specified. Use 'interactive' or 'device'."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Invalid AuthType specified: $AuthType" -LogLevel Error
                return $null
            }
        }

        # Fall back to manual code input if automatic flow failed
        if (-not $automaticFlowSuccess) {
            Write-Log -LogFile $LogFile -Module $functionName -Message "Automatic flow was not successful or auth is set to private, falling back to manual code input"
            Write-Verbose "[$functionName] Automatic flow was not successful or auth is set to private, falling back to manual code input"
            # Step 1: Open the authorization URL
            $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?client_id=$clientId&response_type=code&redirect_uri=$encodedRedirectUri&response_mode=query&scope=$encodedScopes&state=$state"
            Write-Verbose "[$functionName] Authorization URL: $authUrl"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Authorization URL: $authUrl"
            Write-Host "Opening browser for user authentication and consent..."
            if (-not (LaunchBrowser -url $authUrl -browser $preferredBrowser)) {
                Write-Error "Failed to launch browser. Please open the URL manually: $authUrl"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to launch browser for manual authentication" -LogLevel Error
                return $null
            }
            Write-Log -LogFile $LogFile -Module $functionName -Message "Browser launched for manual authentication"
            Write-Host "After granting consent, copy the 'code' parameter from the redirected URL and paste it below."
            $code = Read-Host "Enter the authorization code"
            Write-Verbose "[$functionName] Received authorization code input from user"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Received authorization code input from user"
            #The $code string above contains the intire URL. Extract the code from the URL.
            if ($code -match '.*code=([^&]+).*') {
                $code = $Matches[1]
                Write-Verbose "[$functionName] Extracted code from URL: $($code.Substring(0, [Math]::Min(10, $code.Length)))..."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Successfully extracted authorization code from URL"
            }
            else {
                Write-Warning "Could not extract code parameter from the provided URL"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Could not extract code parameter from provided URL" -LogLevel Warning
                Write-Verbose "[$functionName] Input received: $($code.Substring(0, [Math]::Min(30, $code.Length)))..."
            }
        }
        # Regardless of how we got the code, exchange it for a token
        if ($code -and $AuthType -ne 'PublicAuthFlow') {
            Write-Verbose "[$functionName] Exchanging authorization code for access token"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Exchanging authorization code for access token"
            $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
            $tokenRequestBody = @{
                client_id     = $clientId
                client_secret = $clientSecret
                code          = $code
                redirect_uri  = $redirectUri
                grant_type    = "authorization_code"
                scope         = $scopesFormatted
            }
            Write-Verbose "[$functionName] Token request parameters:"
            Write-Verbose "[$functionName]   Endpoint: $tokenEndpoint"
            Write-Verbose "[$functionName]   Client ID: $clientId"
            Write-Verbose "[$functionName]   Redirect URI: $redirectUri"
            Write-Verbose "[$functionName]   Grant Type: authorization_code"
            Write-Verbose "[$functionName]   Scopes: $scopesFormatted"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token request endpoint: $tokenEndpoint"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token request grant type: authorization_code"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token request scopes: $scopesFormatted"
            try {
                Write-Verbose "[$functionName] Sending token request to $tokenEndpoint"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Sending token request to endpoint"
                $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $tokenRequestBody -ErrorVariable tokenError
                Write-Verbose "[$functionName] Access token received successfully"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Access token received successfully"
            }
            catch {
                Write-Error "Failed to get delegated access token: $_"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to get delegated access token: $_" -LogLevel Error
                if ($tokenError) {
                    Write-Verbose "[$functionName] Error details: $($tokenError | Out-String)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Token error details captured" -LogLevel Error
                    if ($_.ErrorDetails.Message) {
                        Write-Verbose "[$functionName] Error message details: $($_.ErrorDetails.Message)"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Error message details: $($_.ErrorDetails.Message)" -LogLevel Error
                        try {
                            $errorJson = $_.ErrorDetails.Message | ConvertFrom-Json
                            Write-Verbose "[$functionName] Error JSON: $($errorJson | ConvertTo-Json -Depth $maxJSONDepth)"
                            Write-Verbose "[$functionName] Error code: $($errorJson.error)"
                            Write-Verbose "[$functionName] Error description: $($errorJson.error_description)"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Error code: $($errorJson.error), Description: $($errorJson.error_description)" -LogLevel Error
                        }
                        catch {
                            Write-Verbose "[$functionName] Could not convert error details to JSON: $_"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Could not convert error details to JSON" -LogLevel Error
                        }
                    }
                }
                if ($_.Exception.Response) {
                    Write-Verbose "[$functionName] Status code: $($_.Exception.Response.StatusCode)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP Status code: $($_.Exception.Response.StatusCode)" -LogLevel Error
                    # Try to get more information from the response
                    try {
                        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                        $responseBody = $reader.ReadToEnd()
                        $reader.Close()
                        Write-Verbose "[$functionName] Response body: $responseBody"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Response body: $responseBody" -LogLevel Error
                        try {
                            $responseJson = $responseBody | ConvertFrom-Json
                            Write-Verbose "[$functionName] Error code: $($responseJson.error)"
                            Write-Verbose "[$functionName] Error description: $($responseJson.error_description)"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Parsed error - Code: $($responseJson.error), Description: $($responseJson.error_description)" -LogLevel Error
                        }
                        catch {
                            Write-Verbose "[$functionName] Could not parse response body as JSON"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse response body as JSON" -LogLevel Error
                        }
                    }
                    catch {
                        Write-Verbose "[$functionName] Could not read response stream: $_"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Could not read response stream: $_" -LogLevel Error
                    }
                }
                return $null
            }
        }
        # Log the token response properties (without exposing the actual token)
        if ($tokenResponse) {
            Write-Verbose "[$functionName] Token response contains the following properties:"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Processing token response"
            foreach ($prop in $tokenResponse.PSObject.Properties.Name) {
                if ($prop -eq "access_token" -or $prop -eq "refresh_token" -or $prop -eq "id_token") {
                    $tokenLength = $tokenResponse.$prop.Length
                    Write-Verbose "[$functionName]   $($prop): [Token of length $tokenLength]"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Token response contains $($prop) of length $tokenLength"
                }
                else {
                    Write-Verbose "[$functionName]   $($prop): $($tokenResponse.$prop)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Token response property: $($prop) = $($tokenResponse.$prop)"
                }
            }

            $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain
            Write-Log -LogFile $LogFile -Module $functionName -Message "Retrieved token from response"
            # Cache the access token based on cache type
            Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token saved to cache (type: $cacheType)"
            # Save the refresh token to config file regardless of cache type
            if ($tokenResponse.refresh_token) {
                if (-not $NoSaveRefreshToken) {
                    Save-RefreshTokenToConfig -refreshToken $tokenResponse -configFilePath $configFilePath
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Refresh token saved to config file: $configFilePath"
                    if ($ForcedRenewal) {
                        Write-Verbose "[$functionName] Forced refresh token renewed and saved to config file $configFilePath"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Forced refresh token renewed and saved to config file $configFilePath"
                    }
                }
                else {
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Refresh token not saved due to NoSaveRefreshToken setting"
                    if ($ForcedRenewal) {
                        Write-Host "New refresh token has been obtained but was not saved due to NoSaveRefreshToken setting." -ForegroundColor Yellow
                        Write-Verbose "[$functionName] Forced refresh token renewal completed but not saved per NoSaveRefreshToken setting."
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Forced refresh token renewal completed but not saved per NoSaveRefreshToken setting" -LogLevel Warning
                    }
                }
            }
            Write-Log -LogFile $LogFile -Module $functionName -Message "Returning formatted token output"
            return Format-TokenOutput -token $tokenResponse.access_token -secureString $SecureString
        }
        else {
            Write-Verbose "[$functionName] Token response is null. Authorization code exchange failed."
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token response is null. Authorization code exchange failed." -LogLevel Error
            return $null
        }
    }

    function Get-ClientCredentialsToken {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$tenantId,
            [Parameter(Mandatory = $true)]
            [string]$clientId,
            [Parameter(Mandatory = $false)]
            [string]$clientSecret,
            [Parameter(Mandatory = $false)]
            [string]$certificateThumbprint,
            [Parameter(Mandatory = $false)]
            [string]$domain,
            [Parameter(Mandatory = $false)]
            [string]$cacheType,
            [Parameter(Mandatory = $false)]
            [string]$cacheTokenFile,
            [Parameter(Mandatory = $false)]
            [string]$cacheFolder,
            [Parameter(Mandatory = $false)]
            [switch]$secureString
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Starting Get-ClientCredentialsToken function"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Starting Get-ClientCredentialsToken - tenantId=$tenantId, clientId=$clientId, domain=$domain, cacheType=$cacheType"
        Write-Verbose "[$functionName] Using non-delegated access (client credentials flow)"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Using non-delegated access (client credentials flow)"

        # Validate that we have at least one authentication method
        if (-not $clientSecret -and -not $certificateThumbprint) {
            $errorMsg = "Either clientSecret or certificateThumbprint must be provided"
            Write-Error "[$functionName] $errorMsg"
            Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error
            return $null
        }

        # Determine authentication strategy
        $useCertificate = $false
        $useClientSecret = $false
        $tryBothWithFallback = $false

        # Check for non-empty certificate and client secret (not just presence)
        $hasCertificate = -not [string]::IsNullOrWhiteSpace($certificateThumbprint)
        $hasClientSecret = -not [string]::IsNullOrWhiteSpace($clientSecret)

        if ($hasCertificate -and $hasClientSecret) {
            Write-Verbose "[$functionName] Both certificate and client secret provided - will try certificate first with fallback to secret"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Both certificate and client secret provided - will try certificate first with fallback to secret" -LogLevel Warning
            $tryBothWithFallback = $true
            $useCertificate = $true
        }
        elseif ($hasCertificate) {
            Write-Verbose "[$functionName] Using certificate-based authentication (certificate-only mode)"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Using certificate-based authentication (certificate-only mode) with thumbprint: $certificateThumbprint"
            $useCertificate = $true
        }
        else {
            Write-Verbose "[$functionName] Using client secret authentication"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Using client secret authentication"
            $useClientSecret = $true
        }

        $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        $scope = 'https://graph.microsoft.com/.default'
        Write-Verbose "[$functionName] Token endpoint: $tokenEndpoint"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Token endpoint: $tokenEndpoint"
        Write-Verbose "[$functionName] Scope: $scope"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Scope: $scope"

        # Certificate authentication attempt
        if ($useCertificate) {
            Write-Verbose "[$functionName] Attempting certificate-based authentication"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Attempting certificate-based authentication"

            try {
                # Get certificate from store
                Write-Verbose "[$functionName] Retrieving certificate with thumbprint: $certificateThumbprint"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Retrieving certificate from store with thumbprint: $certificateThumbprint"

                $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -Recurse |
                Where-Object { $_.Thumbprint -eq $certificateThumbprint } |
                Select-Object -First 1

                if (-not $certificate) {
                    $errorMsg = "Certificate with thumbprint $certificateThumbprint not found in certificate stores"
                    Write-Verbose "[$functionName] $errorMsg"
                    Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                    if ($tryBothWithFallback) {
                        Write-Verbose "[$functionName] Certificate authentication failed - falling back to client secret"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed - falling back to client secret" -LogLevel Warning
                        $useCertificate = $false
                        $useClientSecret = $true
                    }
                    else {
                        return $null
                    }
                }

                if ($certificate) {
                    Write-Verbose "[$functionName] Certificate found: Subject=$($certificate.Subject), NotAfter=$($certificate.NotAfter)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate found: Subject=$($certificate.Subject), NotAfter=$($certificate.NotAfter)"

                    # Check if certificate has expired
                    if ($certificate.NotAfter -lt (Get-Date)) {
                        $errorMsg = "Certificate has expired on $($certificate.NotAfter)"
                        Write-Verbose "[$functionName] $errorMsg"
                        Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                        if ($tryBothWithFallback) {
                            Write-Verbose "[$functionName] Certificate expired - falling back to client secret"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate expired - falling back to client secret" -LogLevel Warning
                            $useCertificate = $false
                            $useClientSecret = $true
                            $certificate = $null
                        }
                        else {
                            return $null
                        }
                    }

                    # Check if certificate has a private key
                    if ($certificate -and -not $certificate.HasPrivateKey) {
                        $errorMsg = "Certificate does not have a private key"
                        Write-Verbose "[$functionName] $errorMsg"
                        Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                        if ($tryBothWithFallback) {
                            Write-Verbose "[$functionName] Certificate has no private key - falling back to client secret"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate has no private key - falling back to client secret" -LogLevel Warning
                            $useCertificate = $false
                            $useClientSecret = $true
                            $certificate = $null
                        }
                        else {
                            return $null
                        }
                    }
                }

                if ($certificate) {
                    # Create JWT client assertion
                    Write-Verbose "[$functionName] Creating JWT client assertion"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Creating JWT client assertion"

                    $now = [Math]::Floor([decimal](Get-Date (Get-Date).ToUniversalTime() -UFormat "%s"))
                    $exp = $now + 600 # Token valid for 10 minutes

                    # Create JWT header
                    $jwtHeader = @{
                        alg = "RS256"
                        typ = "JWT"
                        x5t = [Convert]::ToBase64String($certificate.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
                    } | ConvertTo-Json -Compress

                    # Create JWT payload
                    $jwtPayload = @{
                        aud = $tokenEndpoint
                        iss = $clientId
                        sub = $clientId
                        jti = [guid]::NewGuid().ToString()
                        nbf = $now
                        exp = $exp
                    } | ConvertTo-Json -Compress

                    # Base64Url encode header and payload
                    $jwtHeaderEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)) -replace '\+', '-' -replace '/', '_' -replace '='
                    $jwtPayloadEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)) -replace '\+', '-' -replace '/', '_' -replace '='

                    # Create signature
                    $jwtToSign = "$jwtHeaderEncoded.$jwtPayloadEncoded"
                    $jwtBytes = [System.Text.Encoding]::UTF8.GetBytes($jwtToSign)

                    Write-Verbose "[$functionName] Signing JWT with certificate private key"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Signing JWT with certificate private key"
                    $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                    if (-not $privateKey) {
                        $errorMsg = "Failed to access certificate private key"
                        Write-Warning "[$functionName] $errorMsg"
                        Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                        if ($tryBothWithFallback) {
                            Write-Warning "[$functionName] Cannot access private key - falling back to client secret"
                            Write-Log -LogFile $LogFile -Module $functionName -Message "Cannot access private key - falling back to client secret" -LogLevel Warning
                            $useCertificate = $false
                            $useClientSecret = $true
                            $certificate = $null
                            # Skip to client secret authentication
                            throw "PrivateKeyAccessFailed"
                        }
                        else {
                            return $null
                        }
                    }

                    if ($privateKey -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
                        $signature = $privateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                    }
                    else {
                        # For CNG keys
                        $signature = $privateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                    }

                    $jwtSignatureEncoded = [Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='
                    $clientAssertion = "$jwtToSign.$jwtSignatureEncoded"

                    Write-Verbose "[$functionName] JWT client assertion created successfully"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "JWT client assertion created successfully"

                    # Build request body with certificate authentication
                    $body = @{
                        client_id             = $clientId
                        scope                 = $scope
                        client_assertion      = $clientAssertion
                        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                        grant_type            = 'client_credentials'
                    }

                    Write-Verbose "[$functionName] Sending token request with certificate authentication"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Sending token request with certificate authentication"

                    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

                    Write-Verbose "[$functionName] Access token received successfully via certificate authentication"
                    Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Access token received successfully via certificate authentication. Expires in: $($tokenResponse.expires_in) seconds"

                    $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain
                    Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Token cached successfully (type: $cacheType)"

                    return Format-TokenOutput -token $tokenResponse.access_token -secureString $secureString
                }
            }
            catch {
                Write-Error "[$functionName] Certificate authentication failed: $_"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed: $_" -LogLevel Error
                if ($_.Exception.Response) {
                    Write-Verbose "[$functionName] Status code: $($_.Exception.Response.StatusCode)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP Status code: $($_.Exception.Response.StatusCode)" -LogLevel Error
                    try {
                        $errorResponse = $_.Exception.Response.GetResponseStream()
                        $streamReader = New-Object System.IO.StreamReader($errorResponse)
                        $errorMessage = $streamReader.ReadToEnd()
                        $streamReader.Close()
                        Write-Verbose "[$functionName] Server Response: $errorMessage"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Server Response: $errorMessage" -LogLevel Error

                        $errorJson = $errorMessage | ConvertFrom-Json
                        Write-Verbose "[$functionName] Error code: $($errorJson.error)"
                        Write-Verbose "[$functionName] Error description: $($errorJson.error_description)"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Error code: $($errorJson.error), Description: $($errorJson.error_description)" -LogLevel Error
                    }
                    catch {
                        Write-Verbose "[$functionName] Could not parse error response: $_"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse error response" -LogLevel Error
                    }
                }

                if ($tryBothWithFallback) {
                    Write-Warning "[$functionName] Certificate authentication failed - falling back to client secret"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed - falling back to client secret" -LogLevel Warning
                    $useCertificate = $false
                    $useClientSecret = $true
                }
                else {
                    return $null
                }
            }
        }

        # Client secret authentication attempt
        if ($useClientSecret) {
            Write-Verbose "[$functionName] Attempting client secret authentication"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Attempting client secret authentication"

            $body = @{
                client_id     = $clientId
                scope         = $scope
                client_secret = $clientSecret
                grant_type    = 'client_credentials'
            }

            Write-Verbose "[$functionName] Token request body: client_id=$clientId, scope=$scope, grant_type=client_credentials"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token request body prepared: client_id=$clientId, scope=$scope, grant_type=client_credentials"

            try {
                Write-Verbose "[$functionName] Sending request to token endpoint: $tokenEndpoint"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Sending token request with client secret authentication"

                $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

                Write-Verbose "[$functionName] Access token received successfully via client secret authentication"
                Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Access token received successfully via client secret authentication. Expires in: $($tokenResponse.expires_in) seconds"

                $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain
                Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
                Write-Log -LogFile $LogFile -Module $functionName -Message "Token cached successfully (type: $cacheType)"

                return Format-TokenOutput -token $tokenResponse.access_token -secureString $secureString
            }
            catch {
                Write-Error "[$functionName] Failed to get access token via client secret: $_"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to get access token via client secret: $_" -LogLevel Error

                if ($_.Exception.Response) {
                    Write-Verbose "[$functionName] Status code: $($_.Exception.Response.StatusCode)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP Status code: $($_.Exception.Response.StatusCode)" -LogLevel Error

                    try {
                        $errorResponse = $_.Exception.Response.GetResponseStream()
                        $streamReader = New-Object System.IO.StreamReader($errorResponse)
                        $errorMessage = $streamReader.ReadToEnd()
                        $streamReader.Close()
                        Write-Verbose "[$functionName] Server Response: $errorMessage"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Server Response: $errorMessage" -LogLevel Error

                        $errorJson = $errorMessage | ConvertFrom-Json
                        Write-Verbose "[$functionName] Error code: $($errorJson.error)"
                        Write-Verbose "[$functionName] Error description: $($errorJson.error_description)"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Error code: $($errorJson.error), Description: $($errorJson.error_description)" -LogLevel Error
                    }
                    catch {
                        Write-Verbose "[$functionName] Could not parse error response: $_"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse error response" -LogLevel Error
                    }
                }

                return $null
            }
        }
    }
    #endregion helper functions

    $functionName = $MyInvocation.MyCommand.Name
    #region Process config files
    Write-Log -LogFile $LogFile -Module $functionName -Message "Starting Graph access token retrieval" -LogLevel "Verbose"
    # Read and process configuration file
    if (-not $configFile) {
        Write-Error "Config file not found. Please provide a valid config file."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Config file not found. Please provide a valid config file." -LogLevel "Verbose"
        return $null
    }

    $config = Get-Content -Path $configFile -Raw | ConvertFrom-Json
    $configRefreshToken = $null

    # Read the refresh token if it exists
    if ($configRefreshToken) {
        Write-Verbose "[$functionName] Found refresh token in encrypted config."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Found refresh token in encrypted config" -LogLevel "Verbose"
    }
    else {
        Write-Verbose "[$functionName] No refresh token found in config."
        Write-Log -LogFile $LogFile -Module $functionName -Message "No refresh token found in config" -LogLevel "Verbose"
    }

    # Handle ForceNewRefreshToken parameter
    if ($ForceNewRefreshToken -and $delegated) {
        Write-Host "Force new refresh token requested. Invalidating existing refresh token." -ForegroundColor Yellow
        Write-Verbose "[$functionName] ForceNewRefreshToken requested. Clearing existing refresh token from memory."
        # Clear the configRefreshToken variable so a new one will be obtained
        $configRefreshToken = $null
        Write-Verbose "[$functionName] Cleared configRefreshToken variable. New refresh token will be obtained and saved to config."
    }

    #get the tenant Id
    $tenantId = $config.tenantId
    if ($tenantId) {
        Write-Verbose "[$functionName] Tenant ID found in config: $tenantId"
    }
    else {
        Write-Error "Tenant ID not found in config file."
        return $null
    }

    #get the domain
    $domain = $config.domain
    if ($domain) {
        Write-Verbose "[$functionName] Domain found in config: $domain"
    }
    else {
        Write-Error "Domain not found in config file."
        return $null
    }
    #get the client id
    $clientId = $config.appId
    if ($clientId) {
        Write-Verbose "[$functionName] Client ID found in config: $clientId"
    }
    else {
        Write-Error "Client ID not found in config file."
        return $null
    }

    #get the client secret
    $clientSecret = $config.appSecret
    if ($clientSecret) {
        Write-Verbose "[$functionName] Client Secret found in config."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Client Secret found in config" -LogLevel "Verbose"
    }
    else {
        Write-Verbose "[$functionName] Client Secret not found in config file."
        Write-Verbose "Checking whether the authtype is public flow which does not require client secret."
        if ($AuthType -ne 'PublicAuthFlow') {
            Write-Verbose "[$functionName] Will check for certificate thumbprint as alternative authentication method."
        }
        else {
            Write-Verbose "[$functionName] Public Auth Flow does not require Client Secret."
        }
    }

    # Check for certificate thumbprint (try both "certificateThumbprint" and "thumbprint" for compatibility)
    $certificateThumbprint = $config.thumbprint
    if ($certificateThumbprint) {
        Write-Verbose "[$functionName] Certificate thumbprint found in config: $certificateThumbprint"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate thumbprint found in config: $certificateThumbprint" -LogLevel "Verbose"
    }
    else {
        Write-Verbose "[$functionName] Certificate thumbprint not found in config."
    }

    # Validate that we have at least one authentication method for non-delegated flows
    if (-not $delegated -and -not $clientSecret -and -not $certificateThumbprint) {
        Write-Error "[$functionName] Either client secret or certificate thumbprint must be provided for non-delegated authentication."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Either client secret or certificate thumbprint must be provided for non-delegated authentication" -LogLevel "Error"
        return $null
    }
    if ($delegated ) {
        Write-Verbose "[$functionName] Delegated access selected. Checking for scope."
        if ($null -eq $scope) {
            Write-Verbose "[$functionName] No scope provided in parameters. Checking config file for scope."
            $Scope = $config.Scope
            Write-Host "Scope: $Scope"
            if ($Scope) {
                Write-Verbose "[$functionName] Found scope in config file."
            }
            else {
                Write-Error "No scope provided."
                return $null
            }
        }
        else {
            Write-Verbose "[$functionName] Scope provided in parameters: $Scope"
        }
    }
    else {
        Write-Verbose "[$functionName] Non-delegated access selected. No scope required."
        Write-Verbose "[$functionName] Using default scope: $Scope"
    }
    #endregion Process config files

    #region Log parameters
    Write-Verbose "[$functionName] Received parameters:"
    Write-Verbose "[$functionName] Configuration File: $configFile"
    Write-Verbose "[$functionName] Renewal Lead Time: $renewalLeadTime"
    Write-Verbose "[$functionName] Secure String: $SecureString"
    Write-Verbose "[$functionName] Force New Token: $ForceNewToken"
    Write-Verbose "[$functionName] Force New Refresh Token: $ForceNewRefreshToken"
    Write-Verbose "[$functionName] Use Public Auth Flow: $UsePublicAuthFlow"
    Write-Verbose "[$functionName] Interactive: $Interactive"
    Write-Verbose "[$functionName] Cache Type: $CacheType"
    Write-Verbose "[$functionName] Domain: $domain"
    Write-Verbose "[$functionName] delegated: $delegated"
    Write-Verbose "[$functionName] Scopes: $Scope"
    Write-Verbose "[$functionName] Config has refresh token: $($null -ne $configRefreshToken)"
    #endregion Log parameters

    # Set up cache paths
    $cacheFolder = Split-Path $configFile
    $cacheTokenFile = Join-Path $cacheFolder "accessToken.json"

    #region Try to get token from cache if not forcing new token
    $accessToken = $null
    if (-not $ForceNewToken) {
        Write-Verbose "[$functionName] Attempting to retrieve token from cache."
        Write-Log -logFile $logFile -Module $functionName -Message "Attempting to retrieve token from cache."
        $accessToken = Get-TokenFromCache -cacheType $CacheType -domain $domain -renewalLeadTime $renewalLeadTime `
            -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId -scopes $Scope `
            -delegated $delegated -cacheFolder $cacheFolder -cacheTokenFile $cacheTokenFile `
            -secureString $SecureString -configFilePath $configFile -configRefreshToken $configRefreshToken
        if ($accessToken) {
            Write-Verbose "[$functionName] Successfully retrieved valid token from cache."
            Write-Log -logFile $logFile -Module $functionName -Message "Successfully retrieved valid token from cache."
            return $accessToken
        }
    }
    else {
        Write-Host "Force new token requested. Ignoring cache."
        Write-Log -logFile $logFile -Module $functionName -Message "Force new token requested. Ignoring cache."
    }
    #endregion Try to get token from cache if not forcing new token

    #region Authentication flow
    if ($delegated) {
        Write-Verbose "[$functionName] delegated authentication flow selected."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Using delegated authentication flow"
        $params = @{
            tenantId           = $tenantId
            clientId           = $clientId
            scopes             = $Scope
            domain             = $domain
            cacheType          = $CacheType
            AuthType           = $AuthType
            cacheTokenFile     = $cacheTokenFile
            cacheFolder        = $cacheFolder
            configFilePath     = $configFile
            configRefreshToken = $configRefreshToken
        }

        switch ($AuthType) {
            PublicAuthFlow {
                Write-Verbose "[$functionName] Using public authentication flow for delegated token."
                Write-Log -logFile $logFile -Module $functionName -Message "Using public authentication flow for delegated token."
            }
            Interactive {
                Write-Verbose "[$functionName] Using interactive authentication flow for delegated token."
                Write-Log -logFile $logFile -Module $functionName -Message "Using interactive authentication flow for delegated token."
                $params += @{
                    clientSecret = $clientSecret
                }
            }
            Private {
                Write-Verbose "[$functionName] Using private authentication flow for delegated token."
                Write-Log -logFile $logFile -Module $functionName -Message "Using private authentication flow for delegated token."
                $params += @{
                    clientSecret = $clientSecret
                }
            }
        }
        if ($NoSaveRefreshToken) {
            Write-Verbose "[$functionName] No save refresh token option selected. Not saving refresh token."
            Write-Log -logFile $logFile -Module $functionName -Message "No save refresh token option selected. Not saving refresh token."
            $params += @{
                NoSaveRefreshToken = $NoSaveRefreshToken
            }
        }
        if ($ForceNewRefreshToken) {
            Write-Verbose "[$functionName] Force new refresh token requested. This will force a new authentication flow."
            Write-Log -logFile $logFile -Module $functionName -Message "Force new refresh token requested. This will force a new authentication flow."
            # Add a marker to indicate this was a forced refresh token renewal
            $params += @{
                ForcedRenewal = $true
            }
        }
        return Get-DelegatedToken @params
    }
    else {
        # Non-delegated (client credentials) authentication
        Write-Verbose "[$functionName] Using non-delegated (client credentials) authentication flow."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Using non-delegated (client credentials) authentication flow" -LogLevel "Verbose"

        if ($tenantId -and $clientId -and ($clientSecret -or $certificateThumbprint)) {
            Write-Verbose "[$functionName] Preparing parameters for client credentials token retrieval."
            Write-Log -logFile $logFile -Module $functionName -Message "Preparing parameters for client credentials token retrieval."
            $params = @{
                tenantId       = $tenantId
                clientId       = $clientId
                domain         = $domain
                cacheType      = $CacheType
                cacheTokenFile = $cacheTokenFile
                cacheFolder    = $cacheFolder
            }

            if ($clientSecret) {
                Write-Verbose "[$functionName] Using client secret for authentication."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using client secret for authentication" -LogLevel "Verbose"
                $params['clientSecret'] = $clientSecret
            }

            if ($certificateThumbprint) {
                Write-Verbose "[$functionName] Using certificate thumbprint for authentication: $certificateThumbprint"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Using certificate thumbprint for authentication: $certificateThumbprint" -LogLevel "Verbose"
                $params['certificateThumbprint'] = $certificateThumbprint
            }

            if ($clientSecret -and $certificateThumbprint) {
                Write-Verbose "[$functionName] Both client secret and certificate available - will try certificate first with fallback to secret."
                Write-Log -LogFile $LogFile -Module $functionName -Message "Both client secret and certificate available - will try certificate first with fallback to secret" -LogLevel "Warning"
            }

            if ($SecureString) {
                Write-Verbose "[$functionName] Secure string option selected. Returning token as SecureString."
                Write-Log -logFile $logFile -Module $functionName -Message "Secure string option selected. Returning token as SecureString."
                $params['secureString'] = $true
            }

            return Get-ClientCredentialsToken @params
        }
        else {
            Write-Error "[$functionName] Missing required authentication parameters (tenantId, clientId, and either clientSecret or certificateThumbprint)"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Missing required authentication parameters" -LogLevel "Error"
            return $null
        }
    }
    #endregion Authentication flow
}

