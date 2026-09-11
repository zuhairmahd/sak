function Request-AdditionalScopes()
{
    <#
    .SYNOPSIS
    Requests additional Microsoft Graph API scopes when current token has insufficient permissions.

    .DESCRIPTION
    This function handles the process of requesting additional scopes when a token is found to be
    missing required permissions. For delegated authentication, it re-initiates the auth flow with
    the combined scope set (current + missing). For application authentication, it reports that
    admin consent is required. The function provides user-friendly prompts and error handling.

    .PARAMETER MissingScopes
    Array of scope strings that are missing from the current token.

    .PARAMETER AuthConfiguration
    Hashtable containing authentication configuration including Delegated flag and current scopes.
    This parameter is mandatory.

    .PARAMETER CurrentScopes
    Array of scopes currently available in the token.

    .PARAMETER AuthParams
    Hashtable containing authentication parameters (tenantId, clientId, etc) for token acquisition.
    This parameter is mandatory.

    .OUTPUTS
    System.Collections.Hashtable
    Returns hashtable with properties:
    - Success: Boolean indicating if scope request succeeded
    - NewAccessToken: The new access token with additional scopes (if successful)
    - ErrorMessage: Error message if operation failed
    - UserCancelled: Boolean indicating if user cancelled the operation

    .EXAMPLE
    $result = Request-AdditionalScopes -MissingScopes @("User.ReadWrite.All") -AuthConfiguration $authConfig -CurrentScopes $current -AuthParams $params

    .NOTES
    For delegated auth: re-authenticates with combined scope set.
    For application auth: returns error indicating admin consent needed.
    Handles empty missing scopes gracefully (returns success).
    Provides clear user prompts for additional permission requests.
    Compatible with PowerShell 5.1.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [array]$MissingScopes = @(),
        [Parameter(Mandatory = $true)]
        [hashtable]$AuthConfiguration,
        [Parameter(Mandatory = $false)]
        [array]$CurrentScopes = @(),
        [Parameter(Mandatory = $true)]
        [hashtable]$AuthParams
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting additional scope request process"
Write-Log -LogFile $LogFile -Module $functionName -Message "Starting additional scope request for $($MissingScopes.Count) missing scopes" -LogLevel "Verbose"
    Write-Verbose "[$functionName] Starting additional scope request for $($MissingScopes.Count) missing scopes (Write-Log not available)"

    $result = @{
        Success        = $false
        NewAccessToken = $null
        ErrorMessage   = ""
        UserCancelled  = $false
    }

    # Handle empty missing scopes - nothing to request
    if ($MissingScopes.Count -eq 0)
    {
        Write-Log -LogFile $LogFile -Module $functionName -Message "No missing scopes - nothing to request" -LogLevel "Information"
        $result.Success = $true
        $result.ErrorMessage = "No additional scopes needed."
        return $result
    }

    try
    {
        # Check if this is delegated authentication
        $isDelegated = $AuthConfiguration.Delegated -eq $true

        if (-not $isDelegated)
        {
            # For application authentication, we cannot request additional scopes at runtime
            Write-Host "`nApplication authentication is being used." -ForegroundColor Yellow
            Write-Host "Missing application permissions cannot be requested at runtime." -ForegroundColor Yellow
            Write-Host "`nThe following application permissions need to be added by an administrator:" -ForegroundColor Cyan
            Write-Log -logFile $LogFile -module $functionName -message "The following application permissions need to be added by an administrator:" -loglevel "Warning"
            foreach ($scope in $MissingScopes)
            {
                Write-Host "  - $($scope.Scope)" -ForegroundColor White
                Write-Host "    Reason: $($scope.Reason)" -ForegroundColor Gray
                if ($scope.Endpoints -and $scope.Endpoints.Count -gt 0)
                {
                    Write-Host "    Affects endpoints: $($scope.Endpoints -join ', ')" -ForegroundColor Gray
                }
                Write-Host ""
            }

            Write-Host "To add these permissions:" -ForegroundColor Cyan
            Write-Host "1. Go to the Microsoft Entra admin center (https://entra.microsoft.com)" -ForegroundColor White
            Write-Host "2. Navigate to 'Identity' > 'Applications' > 'App registrations'" -ForegroundColor White
            Write-Host "3. Find and select your application" -ForegroundColor White
            Write-Host "4. Go to 'API permissions'" -ForegroundColor White
            Write-Host "5. Click 'Add a permission' > 'Microsoft Graph' > 'Application permissions'" -ForegroundColor White
            Write-Host "6. Add the missing permissions listed above" -ForegroundColor White
            Write-Host "7. Click 'Grant admin consent for [your organization]'" -ForegroundColor White
            Write-Host ""

            $result.ErrorMessage = "Application permissions require administrator consent and cannot be requested at runtime."
            Write-Log -LogFile $LogFile -Module $functionName -Message "Application authentication detected. Cannot request additional scopes at runtime." -LogLevel "Warning"
            Write-Verbose "[$functionName] Application authentication detected. Cannot request additional scopes at runtime."
            return $result
        }

        # For delegated authentication, we can request additional scopes
        Write-Host "`nDelegated authentication is being used." -ForegroundColor Green
        Write-Host "Additional scopes can be requested by re-authenticating." -ForegroundColor Green
        Write-Host "`nThe following delegated permissions are missing:" -ForegroundColor Cyan

        foreach ($scope in $MissingScopes)
        {
            Write-Host "  - $($scope.Scope)" -ForegroundColor White
            Write-Host "    Reason: $($scope.Reason)" -ForegroundColor Gray
            if ($scope.Endpoints -and $scope.Endpoints.Count -gt 0)
            {
                Write-Host "    Affects endpoints: $($scope.Endpoints -join ', ')" -ForegroundColor Gray
            }
            Write-Host ""
        }

        # Ask user if they want to request additional scopes
        Write-Host "Would you like to re-authenticate to request these additional permissions?" -ForegroundColor Yellow
        Write-Host "This will open a browser window for authentication and consent." -ForegroundColor Gray
        Write-Host ""
        $userChoice = Read-Host "Enter 'Yes' to re-authenticate with additional scopes, or 'No' to continue with limited functionality"

        while ($userChoice -notin @('Yes', 'No', 'Y', 'N'))
        {
            Write-Host "Invalid choice. Please enter 'Yes' or 'No'." -ForegroundColor Red
            [console]::beep(1000, 500)
            $userChoice = Read-Host "Enter 'Yes' to re-authenticate, or 'No' to continue with limited functionality"
        }

        if ($userChoice -in @('No', 'N'))
        {
            Write-Host "Continuing with current permissions. Some functionality may be limited." -ForegroundColor Yellow
            $result.UserCancelled = $true
            Write-Log -LogFile $LogFile -Module $functionName -Message "User chose not to re-authenticate for additional scopes" -LogLevel "Information"
            Write-Verbose "[$functionName] User chose not to re-authenticate for additional scopes"
            return $result
        }

        # User chose to re-authenticate - combine current scopes with missing scopes
        Write-Host "Preparing to re-authenticate with additional scopes..." -ForegroundColor Green

        # Build the new scope list
        $newScopes = @()
        $newScopes += $CurrentScopes | Where-Object { $_ -and $_.Trim() }
        $newScopes += $MissingScopes | ForEach-Object { $_.Scope } | Where-Object { $_ -and $_.Trim() }

        # Remove duplicates and ensure we have the basic scopes
        $finalScopes = @('offline_access', 'openid') + ($newScopes | Sort-Object -Unique)
        $finalScopesString = $finalScopes -join ' '

        Write-Verbose "[$functionName] New scope list: $finalScopesString"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Re-authenticating with scopes: $finalScopesString" -LogLevel "Information"
        # Create new auth parameters with additional scopes
        $newAuthParams = $AuthParams.Clone()
        $newAuthParams.Scope = $finalScopes
        $newAuthParams.ForceNewToken = $true

        Write-Host "Re-authenticating with additional scopes..." -ForegroundColor Cyan

        # Request new token with additional scopes
        $newAccessToken = GetGraphAccessToken @newAuthParams

        if ($newAccessToken)
        {
            Write-Host "Successfully re-authenticated with additional scopes!" -ForegroundColor Green
            $result.Success = $true
            $result.NewAccessToken = $newAccessToken
            Write-Log -LogFile $LogFile -Module $functionName -Message "Successfully obtained new access token with additional scopes" -LogLevel "Information"
            Write-Verbose "[$functionName] Successfully obtained new access token with additional scopes"
        }
        else
        {
            Write-Host "Failed to obtain new access token with additional scopes." -ForegroundColor Red
            $result.ErrorMessage = "Failed to obtain new access token during re-authentication."
            Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to obtain new access token with additional scopes" -LogLevel "Error"
        }

        return $result
    }
    catch
    {
        Write-Error "[$functionName] Error during additional scope request: $($_.Exception.Message)"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Error during additional scope request: $($_.Exception.Message)" -LogLevel "Error"
        $result.Success = $false
        $result.ErrorMessage = "Error during additional scope request: $($_.Exception.Message)"
        return $result
    }
}