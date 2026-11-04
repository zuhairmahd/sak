[CmdletBinding()]
param(
    [switch]$continue,
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

#region import functions.
. $PSScriptRoot\functions\Find-FolderPath.ps1
. $PSScriptRoot\functions\Test-PowerShellSyntax.ps1
$functionsFolder = Find-FolderPath -Path "$psscriptRoot" -FolderName "functions"
if (Test-Path $functionsFolder) {
    Write-Verbose "[$scriptName] Importing functions from $functionsFolder"
    $functions = Get-ChildItem -Path "$functionsFolder\*.ps1" -File
    foreach ($function in $functions) {
        Write-Verbose " [$scriptName] Importing function $function"
        $syntaxCheck = Test-PowerShellSyntax -File $function
        if ($syntaxCheck.HasErrors) {
            Write-Host "Syntax errors found in $($function.FullName). Skipping import." -ForegroundColor Red
            write-log -logFile $logFile -Module $scriptName -Message "Syntax errors found in $($function.FullName). Skipping import." -LogLevel "Error"
            continue
        }
        . $function.FullName
    }
}
else {
    Write-Host 'Cannot find the functions folder. Exiting script.' -ForegroundColor Red
    exit 1
}
#endregion import functions.

#region define variables
$scriptName = $MyInvocation.MyCommand.Name
$logFile = Join-Path -Path $env:TEMP\sak -ChildPath "logs\$($scriptName)_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Continue = if ($continue) { $true } else { $false }
$menuItems = @(
    @{
        name        = "GetAccessToken"
        description = "Get an access token you can use in this session"
    },
    @{
        name        = "AddUserToGroup"
        description = "Add a user to a specified group"
    },
    @{
        name        = "CallGraphEndpoint"
        description = "Call a Graph API endpoint"
    },      
    @{
        name        = "Continue"
        description = "Continue executing the rest of the script"
    }
)
#endregion define variables

#region define configuration parameters
$params = @{
    configFile = $configFile
}
$auth = Get-Content -Path $paramsFile -Raw -Force | ConvertFrom-Json
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

if (-not $continue) {
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "Welcome to SAK, the Swiss Army Knife for Intune Administrators!" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan
    #if no commandline parameters are passed, display the menu.
    if (-not $PSBoundParameters.Keys.Count) {
        write-log -logFile $LogFile -Module $scriptName -Message "No parameters provided. Displaying menu for user selection." -LogLevel "Information"
        $userChoice = Show-NumericMenu -choices $menuItems -banner "Select an action to perform:" -RequireEnter
        switch ($userChoice) {
            "CallGraphEndpoint" {
                $endpoint = Get-UserInput -message "Enter the Graph API endpoint" -inputType "string"
                $method = Get-UserInput -message "Enter the HTTP method (GET, POST, PATCH, DELETE)" -inputType "string"
                $body = Get-UserInput -message "Enter the request body (if applicable)" -inputType "string"
                $accessToken = Get-GraphAccessToken @params
                if (-not $endpoint) {
                    throw "Graph API endpoint is required."
                }
                if (-not $method) {
                    throw "HTTP method is required."
                }
                if (-not $accessToken) {
                    throw "Access token could not be acquired."
                }
                $global:response = Invoke-GraphRequest -Endpoint $endpoint -Method $method -Body $body -AccessToken $accessToken
                if ($global:response.statusCode -ge 200 -and $global:response.statusCode -lt 300) {
                    Write-Host "API call successful.`n The value has been returned to `$$global:response`." -ForegroundColor Green                          
                }
                else {
                    Write-Host "API call failed with status code $($global:response.statusCode).`n The value has been returned to `$$global:response`." -ForegroundColor Red
                }
            }                               
            "GetAccessToken" {
                $global:accessToken = Get-GraphAccessToken @params
                Write-Host "Access token acquired successfully." -ForegroundColor Green
            }
            "AddUserToGroup" {
                [array]$userPrincipalNames = Get-UserInput -message "Enter the User Principal Name (UPN) of the user to add" -inputType "array"
                $groupDisplayName = Get-UserInput -message "Enter the display name of the group" -inputType "string"
                $accessToken = Get-GraphAccessToken @params
                #sanity check: make sure we have all parameters
                if (-not $userPrincipalNames -or ($userPrincipalNames.Count -eq 0)) {
                    throw "User Principal Name is required."
                }
                if (-not $groupDisplayName) {
                    throw "Group Display Name is required."
                }
                if (-not $accessToken) {
                    throw "Access token could not be acquired."
                }
                $response = Add-EntraGroupMemberByName -UserPrincipalNames $userPrincipalNames -GroupDisplayName $groupDisplayName -AccessToken $accessToken
                $summaryColor = if ($response.success) { 'Green' } else { 'Yellow' }
                Write-Host "Results: $($response.addedCount) added, $($response.alreadyMemberCount) already member, $($response.failedCount) failed of $($response.totalRequested) requested." -ForegroundColor $summaryColor
                foreach ($result in $response.results) {
                    $resultColor = switch ($result.status) {
                        'Added' { 'Green' }
                        'AlreadyMember' { 'Yellow' }
                        default { 'Red' }
                    }
                    Write-Host "  [$($result.status)] $($result.upn): $($result.message)" -ForegroundColor $resultColor
                }
            }
            "Continue" {
                $continue = $true
            }
            default {
                Write-Host "Exiting script."
                write-log -logFile $logFile -FinishLogging
                write-log -logFile $LogFile -Module $scriptName -Message "Exiting script." -LogLevel "Warning"
                exit 1
            }
        }
    }
    if (-not $continue) {
        exit 0
    }
}


# $templateId = "66df8dce-0166-4b82-92f7-1f74e3ca17a3_5"
# $uri = "deviceManagement/configurationPolicyTemplates/$templateId/settingTemplates"
# $uri = "deviceManagement/templateSettings/$baseId"
$uri = "deviceManagement/windowsAutopilotDeviceIdentities"
$extraParameters = "top=1000"
# $filter = "templateType eq 'securityBaseline'"
$consistencyLevel = $true
# $APIVersion = "beta"
# $method = ""
# $body = ""
# $search = ""


$accessToken = Get-GraphAccessToken @params
$apiParams = Get-ApiParams -accessToken $accessToken -uri $uri -extraParameters $extraParameters -consistencyLevel $consistencyLevel -method $method -APIVersion $APIVersion -Search $search -Body $body -headers $headers -filter $filter
if ($accessToken) {
    Write-Host "Got access token"
    $global:response = Invoke-GraphAPI @apiParams
    if ($response.statusCode -in 200..299) {
        Write-Host "API call succeeded."
        $numberOfEntriesReturned = if ($response.value) { $response.value.Count } else { 0 }            
        Write-Host "Number of entries returned: $numberOfEntriesReturned"
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

