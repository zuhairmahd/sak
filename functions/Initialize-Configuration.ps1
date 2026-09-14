function Initialize-Configuration {
    [CmdletBinding()]
    param(
        [string]$configFile
    )

    $functionName = $MyInvocation.MyCommand.Name
    $returnObject = @{
        $success   = $false
        publicFlow = $false
        message    = $null
    }


    Write-Log -LogFile $LogFile -Module $functionName -Message "Starting Graph access token retrieval" -LogLevel "Verbose"
    # Read and process configuration file
    if (-not $configFile) {
        $returnObject.message = "Config file not found. Please provide a valid config file."
        Write-Error $returnObject.message
        Write-Log -LogFile $LogFile -Module $functionName -Message $returnObject.message -LogLevel "Verbose"
        return $returnObject
    }

    $config = Get-Content -Path $configFile -Raw | ConvertFrom-Json
    # Read the refresh token if it exists
    $configRefreshToken = $config.delegatedCredentials.refresh_token
    if ($configRefreshToken) {
        Write-Verbose "[$functionName] Found refresh token in config."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Found refresh token in config" -LogLevel "Verbose"
        $returnObject.refreshToken = $configRefreshToken
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
        $returnObject.refreshToken = $null
        Write-Verbose "[$functionName] Cleared configRefreshToken variable. New refresh token will be obtained and saved to config."
    }

    #get the tenant Id
    $tenantId = $config.tenantId
    if ($tenantId) {
        Write-Verbose "[$functionName] Tenant ID found in config: $tenantId"
        $returnObject.tenantId = $tenantId
    }
    else {
        $returnObject.message = "Tenant ID not found in config file."
        Write-Error $returnObject.message
    }

    #get the domain
    $domain = $config.domain
    if ($domain) {
        Write-Verbose "[$functionName] Domain found in config: $domain"
        $returnObject.domain = $domain
    }
    else {
        $returnObject.message = "Domain not found in config file."
        Write-Verbose "[$functionName] Domain not found in config file."
    }
    #get the client id
    $clientId = $config.appId
    if ($clientId) {
        Write-Verbose "[$functionName] Client ID found in config: $clientId"
        $returnObject.clientId = $clientId
    }
    else {
        $returnObject.message = "Client ID not found in config file."
        Write-Error $returnObject.message
    }

    #get the client secret
    $clientSecret = $config.appSecret
    if ($clientSecret) {
        Write-Verbose "[$functionName] Client Secret found in config."
        Write-Log -LogFile $LogFile -Module $functionName -Message "Client Secret found in config" -LogLevel "Verbose"
        $returnObject.clientSecret = $clientSecret
    }
    else {
        Write-Verbose "[$functionName] Client Secret not found in config file."
        Write-Verbose "Checking whether the authtype is public flow which does not require client secret."
        if ($AuthType -ne 'PublicAuthFlow') {
            Write-Verbose "[$functionName] Will check for certificate thumbprint as alternative authentication method."
        }
        else {
            Write-Verbose "[$functionName] Public Auth Flow does not require Client Secret."
            $returnObject.PublicFlow = $true
        }
    }

    # Check for certificate thumbprint
    $certificateThumbprint = $config.thumbprint
    if ($certificateThumbprint) {
        Write-Verbose "[$functionName] Certificate thumbprint found in config: $certificateThumbprint"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate thumbprint found in config: $certificateThumbprint" -LogLevel "Verbose"
        $returnObject.certificateThumbprint = $certificateThumbprint
    }
    else {
        Write-Verbose "[$functionName] Certificate thumbprint not found in config."
    }

    # Validate that we have at least one authentication method for non-delegated flows
    if (-not $delegated -and -not $clientSecret -and -not $certificateThumbprint) {
        $returnObject.message = "Either client secret or certificate thumbprint must be provided for non-delegated authentication."
        Write-Error $returnObject.message
        Write-Log -LogFile $LogFile -Module $functionName -Message $returnObject.message -LogLevel "Error"
        return $returnObject
    }
    $returnObject.success = $true
    if ($delegated ) {
        Write-Verbose "[$functionName] Delegated access selected. Checking for scope."
        if ($null -eq $scope) {
            Write-Verbose "[$functionName] No scope provided in parameters. Checking config file for scope."
            $Scope = $config.auth.Scope
            Write-Host "Scope: $Scope"
            if ($Scope) {
                Write-Verbose "[$functionName] Found scope in config file."
                $returnObject.scope = $Scope
            }
            else {
                $returnObject.message = "No scope provided."
                Write-Error $returnObject.message
                return $returnObject
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
}