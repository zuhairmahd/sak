function Show-DirectoryObjectList {
    <#
    .SYNOPSIS
        Unified function to display and select from a list of users or groups.

    .DESCRIPTION
        This function consolidates the logic from DisplayUserList and DisplayGroupList into a single
        unified implementation. It handles empty lists, single item confirmation, list truncation,
        menu creation, and consistent exit handling for both users and groups.

        Returns the appropriate identifier based on EntityType:
        - User: userPrincipalName (string)
        - Group: displayName (string)

    .PARAMETER EntityList
        Array of user or group objects from Graph API.

    .PARAMETER EntityType
        Specifies whether displaying 'User' or 'Group' entities.

    .PARAMETER MaxDisplay
        Maximum number of items to display. Defaults to 10.

    .OUTPUTS
        Returns:
        - String identifier (userPrincipalName for users, displayName for groups)
        - Integer 0 for exit
        - Return value message for navigation (back, cancel, etc.)

    .EXAMPLE
        # Display user list
        $selection = Show-DirectoryObjectList -EntityList $users -EntityType User -MaxDisplay 10
        if ($selection -eq 0) { exit }

    .EXAMPLE
        # Display group list
        $selection = Show-DirectoryObjectList -EntityList $groups -EntityType Group -MaxDisplay 5
        Write-Host "Selected group: $selection"

    .NOTES
        Author: Windows Autopilot Management Tool Team
        Version: 1.0.0
        This function replaces DisplayUserList and DisplayGroupList to eliminate code duplication.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$EntityList,
        [Parameter(Mandatory = $true)]
        [ValidateSet("User", "Group")]
        [string]$EntityType,
        [Parameter(Mandatory = $false)]
        [int]$MaxDisplay = 10,
        [Parameter(Mandatory = $false)]
        [switch]$NoPrompt
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Displaying $EntityType list"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Displaying $EntityType list with $($EntityList.Count) items" -LogLevel "Verbose"

    # Handle empty list
    if ($EntityList.Count -eq 0) {
        Write-Host "No ${EntityType}s found." -ForegroundColor Yellow
        return $null
    }

    # Handle single item - prompt for confirmation unless NoPrompt is specified
    if ($EntityList.Count -eq 1) {
        # Return appropriate identifier
        $identifier = if ($EntityType -eq "User") {
            $EntityList[0].userPrincipalName
        }
        else {
            $EntityList[0].displayName
        }

        # If NoPrompt is specified, auto-accept the single item
        if ($NoPrompt) {
            Write-Verbose "[$functionName] NoPrompt specified - auto-accepting single ${EntityType}: $identifier"
            Write-Log -LogFile $LogFile -Module $functionName -Message "NoPrompt - auto-accepting ${EntityType}: $identifier" -LogLevel "Information"
            return $identifier
        }

        Write-Verbose "[$functionName] Only one $EntityType found, prompting for confirmation"

        # Display item based on type
        if ($EntityType -eq "User") {
            Write-Host "$($EntityList[0].displayName): ($($EntityList[0].userPrincipalName))"
        }
        else {
            Write-Host "$($EntityList[0].displayName)"
        }

        # Get Y/N confirmation
        $choice = Read-Host "(Y/N)"
        while ($choice -notin @('Y', 'N', 'y', 'n')) {
            Write-Host "Please enter Y or N." -ForegroundColor Red
            [console]::beep(1000, 500)
            $choice = Read-Host "(Y/N)"
        }

        if ($choice -in @('Y', 'y')) {
            Write-Verbose "[$functionName] $EntityType confirmed: $identifier"
            Write-Log -LogFile $LogFile -Module $functionName -Message "$EntityType confirmed: $identifier" -LogLevel "Information"
            return $identifier
        }
        else {
            Write-Verbose "[$functionName] $EntityType not confirmed, returning userCanceledMessage"
            return "User canceled"
        }
    }

    # Create selection menu
    $menuName = if ($EntityType -eq "User") { "userMenu" } else { "groupMenu" }

    $menu = @()
    Write-Verbose "[$functionName] Creating $EntityType menu with $($EntityList.Count) items"

    # Add each entity as a menu item
    foreach ($entity in $EntityList) {
        # Create display name based on type
        if ($EntityType -eq "User") {
            $menuItemName = "$($entity.displayName): ($($entity.userPrincipalName))"
            $identifier = $entity.userPrincipalName
        }
        else {
            $menuItemName = "$($entity.displayName)"
            $identifier = $entity.displayName
        }

        Write-Verbose "[$functionName] Adding menu item: $menuItemName"
        $menuObject = @{
            Name        = $identifier
            Description = $menuItemName
        }
        $menu += $menuObject
    }

    $selectedEntity = Show-NumericMenu -choices $menu -banner "Select a matching $EntityType to use:" -RequireEnter
    if ($null -eq $selectedEntity) {
        Write-Verbose "[$functionName] ShowMenu returned null - treating as exit command"
        Write-Log -LogFile $LogFile -Module $functionName -Message "User selected exit from $EntityType list" -LogLevel "Information"
        return 0
    }

    # Extract identifier string from the PSCustomObject Show-NumericMenu returns
    if ($selectedEntity -is [PSCustomObject]) {
        $selectedEntity = $selectedEntity.Name
    }

    if ($selectedEntity -is [string]) {
        Write-Verbose "[$functionName] Selected $EntityType`: $selectedEntity"
        Write-Log -LogFile $LogFile -Module $functionName -Message "User selected $EntityType`: $selectedEntity" -LogLevel "Information"
    }

    return $selectedEntity
}

function Get-EntraDirectoryObject {
    <#
    .SYNOPSIS
        Unified function to retrieve users or groups from Entra ID with optional fuzzy matching.

    .DESCRIPTION
        This function consolidates the logic from GetEntraUser and getEntraGroup into a single
        unified implementation. It supports exact and fuzzy matching for both users and groups,

        The function returns a tuple: (EntityInfo, IsFuzzyMatch)
        - EntityInfo: The Graph API response with matching entities
        - IsFuzzyMatch: Boolean indicating if fuzzy search was used

    .PARAMETER EntityType
        Specifies whether to search for 'User' or 'Group'. Required.

    .PARAMETER EntityName
        The name to search for (userPrincipalName for users, displayName for groups).

    .PARAMETER AccessToken
        The access token for authenticating with Microsoft Graph API.

    .PARAMETER FindSimilar
        If specified, performs fuzzy matching when exact match fails.

    .OUTPUTS
        Returns a tuple: (EntityInfo, IsFuzzyMatch)
        - EntityInfo: PSCustomObject with '@odata.context' and 'value' properties
        - IsFuzzyMatch: Boolean ($true for fuzzy, $false for exact)

    .EXAMPLE
        # Search for user with exact match
        $result = Get-EntraDirectoryObject -EntityType User -EntityName "john.doe@contoso.com" -AccessToken $token
        if ($result[1] -eq $false) {
            Write-Host "Exact match found: $($result[0].value[0].displayName)"
        }

    .EXAMPLE
        # Search for group with fuzzy matching
        $result = Get-EntraDirectoryObject -EntityType Group -EntityName "Marketing" -AccessToken $token -FindSimilar
        Write-Host "Found $($result[0].value.Count) matches (Fuzzy: $($result[1]))"

    .NOTES
        Author: Windows Autopilot Management Tool Team
        Version: 1.0.0
        This function replaces GetEntraUser and getEntraGroup to eliminate code duplication.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("User", "Group")]
        [string]$EntityType,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$EntityName,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$AccessToken,
        [Parameter(Mandatory = $false)]
        [switch]$FindSimilar
    )

    $functionName = $MyInvocation.MyCommand.Name
    $substringSearch = $false
    Write-Verbose "[$functionName] Starting function to get $EntityType from Entra ID"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Starting $EntityType search for: $EntityName" -LogLevel "Verbose"

    # Validate access token
    if ([string]::IsNullOrWhiteSpace($AccessToken)) {
        Write-Verbose "[$functionName] AccessToken is null or empty. Cannot proceed."
        Write-Log -LogFile $LogFile -Module $functionName -Message "AccessToken is required but was not provided" -LogLevel "Error"
        return $null
    }

    # Step 1: Attempt exact match
    Write-Verbose "[$functionName] Attempting exact match for $EntityType`: $EntityName"

    if ($EntityType -eq "User") {
        # User exact match by userPrincipalName
        $resourcePath = "users/$EntityName"
        $extraParameters = "select=givenName,surName,displayName,userPrincipalName,mail,id"
        $entityInfo = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -extraParameters $extraParameters
    }
    else {
        $resourcePath = "groups"
        $filter = "displayName eq '$EntityName'"
        $extraParameters = "select=displayName,id,groupTypes"
        $entityInfo = Invoke-GraphAPI -AccessToken $AccessToken -ResourcePath $resourcePath -Filter $filter -ExtraParameters $extraParameters
    }

    Write-Verbose "[$functionName] Exact match API response: $($entityInfo | Out-String)"

    # Check if exact match succeeded
    $exactMatchFound = $false
    if ($EntityType -eq "User") {
        # Invoke-GraphAPI may return an integer code OR an error object with .error/.statusCode
        $exactMatchFound = ($entityInfo -notin 400, 401, 403, 404 -and $null -eq $entityInfo.error)
    }
    else {
        # For groups, filter returns object with value array
        $exactMatchFound = ($entityInfo -notin 400, 401, 403, 404 -and $null -eq $entityInfo.error -and $entityInfo.value -and $entityInfo.value.Count -gt 0)
    }

    if ($exactMatchFound) {
        Write-Verbose "[$functionName] Exact match found for $EntityType`: $EntityName"

        # Create consistent response format
        if ($EntityType -eq "User") {
            $exactMatchResponse = [PSCustomObject]@{
                '@odata.context' = $null
                value            = @($entityInfo)  # Wrap single user in array
            }
            Write-Verbose "[$functionName] User found: $($entityInfo.displayName) ($($entityInfo.userPrincipalName))"
        }
        else {
            $exactMatchResponse = $entityInfo  # Already in correct format
            Write-Verbose "[$functionName] Group found: $($entityInfo.value[0].displayName)"
        }

        $result = $exactMatchResponse, $substringSearch
        return $result
    }

    # Handle error cases - show error message only if FindSimilar is not enabled
    $exactMatchError = if ($entityInfo -in 400, 401, 403, 404) { $entityInfo } elseif ($null -ne $entityInfo.error) { $entityInfo.statusCode } else { $null }
    if ($null -ne $exactMatchError -and -not $FindSimilar) {
        Write-Verbose "[$functionName] Exact match failed with error code: $exactMatchError"
        Write-Log -LogFile $LogFile -Module $functionName -Message "$EntityType lookup failed with error code: $exactMatchError" -LogLevel "Error"

        # Display appropriate error message
        switch ($exactMatchError) {
            400 {
                Write-Host "Bad request. Please check the $EntityType name format." -ForegroundColor Red
            }
            401 {
                Write-Host "Unauthorized. Please check your access token." -ForegroundColor Red
            }
            403 {
                Write-Host "Forbidden. You do not have permission to access this $EntityType." -ForegroundColor Red
            }
            404 {
                Write-Host "$EntityType '$EntityName' not found in Entra ID." -ForegroundColor Red
            }
        }

        return $null
    }

    # Step 2: Perform fuzzy search if FindSimilar is enabled
    if ($FindSimilar) {
        Write-Verbose "[$functionName] No exact match found. Performing fuzzy search for $EntityType"
        $substringSearch = $true

        if ($EntityType -eq "User") {
            # User fuzzy search: use $search for partial UPN/displayName matching
            $searchTerm = $EntityName -replace '@.*$', ''  # Remove domain for better matching
            Write-Verbose "[$functionName] Cleaned search term: $searchTerm"

            $resourcePath = "users"
            $extraParameters = "select=givenName,surName,displayName,userPrincipalName,mail,id"

            # Try $search first — supports partial UPN and displayName matching
            $searchClause = '"displayName:{0}" OR "userPrincipalName:{0}"' -f $searchTerm
            $encodedSearch = [uri]::EscapeDataString($searchClause)
            $searchExtra = "search=$encodedSearch&$extraParameters"

            Write-Verbose "[$functionName] Trying `$search with clause: $searchClause"
            $fallbackResults = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -extraParameters $searchExtra -consistencyLevel

            # Fallback to givenName/surname filter if $search fails
            if ($fallbackResults -in 400, 401, 403, 404 -or $null -ne $fallbackResults.error -or (-not $fallbackResults.value) -or $fallbackResults.value.Count -eq 0) {
                $filterExpression = "startswith(givenName, '$searchTerm') or startsWith(surname, '$searchTerm')"
                Write-Verbose "[$functionName] `$search failed. Trying filter: $filterExpression"
                $fallbackResults = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -filter $filterExpression -extraParameters $extraParameters
            }
        }
        else {
            # Group fuzzy search: try $search first, then filters
            $searchTerm = $EntityName
            $resourcePath = "groups"
            $extraParameters = "select=displayName,id,groupTypes"

            # Try $search with advanced query
            $searchClause = '"displayName:{0}"' -f $searchTerm
            $encodedSearch = [uri]::EscapeDataString($searchClause)
            $searchExtra = "search=$encodedSearch&$extraParameters"

            Write-Verbose "[$functionName] Trying `$search with clause: $searchClause"
            $fallbackResults = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -extraParameters $searchExtra -consistencyLevel

            # Fallback to filters if $search fails
            if ($fallbackResults -in 400, 401, 403, 404 -or (-not $fallbackResults.value) -or $fallbackResults.value.Count -eq 0) {
                Write-Verbose "[$functionName] `$search failed. Trying startsWith filter"
                $filterExpression = "startsWith(displayName, '$searchTerm')"
                $fallbackResults = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -filter $filterExpression -extraParameters $extraParameters

                # If startsWith fails, try contains with advanced query
                if ($fallbackResults -in 400, 401, 403, 404 -or (-not $fallbackResults.value) -or $fallbackResults.value.Count -eq 0) {
                    Write-Verbose "[$functionName] startsWith failed. Trying contains filter with advanced query"
                    $advancedFilterExpression = "contains(displayName, '$searchTerm')"
                    $advancedExtraParameters = "$extraParameters&count=true"
                    $fallbackResults = Invoke-GraphAPI -accessToken $AccessToken -ResourcePath $resourcePath -filter $advancedFilterExpression -extraParameters $advancedExtraParameters -consistencyLevel
                }
            }
        }

        # Process fuzzy search results
        if ($fallbackResults.statusCode -notin 400, 401, 403, 404 -and $null -eq $fallbackResults.error -and $fallbackResults.value -and $fallbackResults.value.Count -gt 0) {
            Write-Verbose "[$functionName] Fuzzy search returned $($fallbackResults.value.Count) results"

            # Apply exclusion patterns
            $exclusionPatterns = if ($EntityType -eq "User") {
                $settings.userPatternsToExclude
            }
            else {
                $settings.groupPatternsToExclude
            }
            $filteredResults = @()

            foreach ($item in $fallbackResults.value) {
                # Check for duplicates
                $uniqueKey = if ($EntityType -eq "User") {
                    $item.userPrincipalName
                }
                else {
                    $item.displayName
                }

                if ($filteredResults.userPrincipalName -contains $uniqueKey -or $filteredResults.displayName -contains $uniqueKey) {
                    Write-Verbose "[$functionName] Skipping duplicate: $uniqueKey"
                    continue
                }

                # Check exclusion patterns
                $excludeItem = $false
                if ($exclusionPatterns -and $exclusionPatterns.Count -gt 0) {
                    foreach ($pattern in $exclusionPatterns) {
                        $matchField = if ($EntityType -eq "User") {
                            $item.userPrincipalName
                        }
                        else {
                            $item.displayName
                        }
                        if ($matchField -match $pattern -or $item.displayName -match $pattern) {
                            Write-Verbose "[$functionName] Excluding $EntityType`: $matchField (matched pattern: $pattern)"
                            $excludeItem = $true
                            break
                        }
                    }
                }

                if (-not $excludeItem) {
                    Write-Verbose "[$functionName] Including $EntityType`: $uniqueKey"
                    $filteredResults += $item
                }
            }

            # Create filtered response
            $sortProperty = if ($EntityType -eq "User") {
                "userPrincipalName"
            }
            else {
                "displayName"
            }
            $filteredResponse = [PSCustomObject]@{
                '@odata.context' = $fallbackResults.'@odata.context'
                value            = @($filteredResults | Sort-Object -Property $sortProperty -Unique)
            }

            Write-Verbose "[$functionName] Filtered to $($filteredResponse.value.Count) unique $EntityType entities after exclusions"

            $result = $filteredResponse, $substringSearch
            return $result
        }
        else {
            Write-Verbose "[$functionName] Fuzzy search failed (Error code: $fallbackResults)"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Fuzzy search failed for $EntityType`: $EntityName" -LogLevel "Warning"
        }
    }

    # No matches found
    Write-Verbose "[$functionName] No matches found for $EntityType`: $EntityName"
    Write-Log -LogFile $LogFile -Module $functionName -Message "No matches found for $EntityType`: $EntityName" -LogLevel "Warning"
    return $null
}

function Add-EntraGroupMemberByName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$UserPrincipalNames,
        [Parameter(Mandatory = $true)]
        [string]$GroupDisplayName,
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,
        [switch]$batchMode
    )

    $functionName = $MyInvocation.MyCommand.Name
    $results = [System.Collections.ArrayList]@()
    $returnObject = @{
        success            = $false
        totalRequested     = $UserPrincipalNames.Count
        addedCount         = 0
        alreadyMemberCount = 0
        failedCount        = 0
        results            = $results
    }

    try {
        # 1. Resolve group; interactive mode prompts when fuzzy match returns multiple results
        $groupInfo = if ($batchMode) {
            Get-EntraDirectoryObject -EntityType Group -EntityName $GroupDisplayName -AccessToken $AccessToken
        }
        else {
            Get-EntraDirectoryObject -EntityType Group -EntityName $GroupDisplayName -AccessToken $AccessToken -FindSimilar
        }
        if ($null -eq $groupInfo) {
            [void]$results.Add(@{ upn = '*'; status = 'Failed'; message = "Group '$GroupDisplayName' was not found in Entra ID." })
            $returnObject.failedCount = $UserPrincipalNames.Count
            return $returnObject
        }
        $selectedGroupName = if (-not $batchMode -and $groupInfo[1]) {
            Show-DirectoryObjectList -EntityList $groupInfo[0].value -EntityType Group
        }
        else {
            $groupInfo[0].value[0].displayName
        }
        $targetGroup = $groupInfo[0].value | Where-Object { $_.displayName -eq $selectedGroupName } | Select-Object -First 1
        if ($null -eq $targetGroup) {
            [void]$results.Add(@{ upn = '*'; status = 'Failed'; message = "Group selection was canceled." })
            $returnObject.failedCount = $UserPrincipalNames.Count
            return $returnObject
        }
        Write-Verbose "[$functionName] Resolved Group: $($targetGroup.displayName) ($($targetGroup.id))"

        # 2. Guard: dynamic groups cannot have members added manually
        if ($targetGroup.groupTypes -contains "DynamicMembership") {
            foreach ($upn in $UserPrincipalNames) {
                [void]$results.Add(@{ upn = $upn; status = 'Failed'; message = "Group '$GroupDisplayName' is a Dynamic Group. Members cannot be added manually." })
            }
            $returnObject.failedCount = $UserPrincipalNames.Count
            return $returnObject
        }

        # 3. Resolve each user; interactive mode prompts when fuzzy match returns multiple results
        $resolvedUsers = @{}
        foreach ($upn in $UserPrincipalNames) {
            $userInfo = if ($batchMode) {
                Get-EntraDirectoryObject -EntityType User -EntityName $upn -AccessToken $AccessToken
            }
            else {
                Get-EntraDirectoryObject -EntityType User -EntityName $upn -AccessToken $AccessToken -FindSimilar
            }
            if ($null -eq $userInfo) {
                Write-Verbose "[$functionName] User not found: $upn"
                [void]$results.Add(@{ upn = $upn; status = 'NotFound'; message = "User '$upn' was not found in Entra ID." })
                $returnObject.failedCount++
                continue
            }
            $selectedUPN = if (-not $batchMode -and $userInfo[1]) {
                Show-DirectoryObjectList -EntityList $userInfo[0].value -EntityType User
            }
            else {
                $userInfo[0].value[0].userPrincipalName
            }
            $targetUser = $userInfo[0].value | Where-Object { $_.userPrincipalName -eq $selectedUPN } | Select-Object -First 1
            if ($null -eq $targetUser) {
                Write-Verbose "[$functionName] User selection canceled for: $upn"
                [void]$results.Add(@{ upn = $upn; status = 'Failed'; message = "User selection was canceled for '$upn'." })
                $returnObject.failedCount++
                continue
            }
            Write-Verbose "[$functionName] Resolved User: $($targetUser.displayName) ($($targetUser.id))"
            $resolvedUsers[$upn] = $targetUser.id
        }

        # 4. Add each resolved user — sequential POST gives per-user error isolation
        $addMemberUri = "groups/$($targetGroup.id)/members/`$ref"
        foreach ($upn in $resolvedUsers.Keys) {
            $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($resolvedUsers[$upn])" } | ConvertTo-Json -Compress
            $addResult = Invoke-GraphAPI -ResourcePath $addMemberUri -Method Post -Body $body -AccessToken $AccessToken

            if ($null -eq $addResult.error) {
                Write-Verbose "[$functionName] Added: $upn"
                [void]$results.Add(@{ upn = $upn; status = 'Added'; message = "Successfully added '$upn' to '$($targetGroup.displayName)'." })
                $returnObject.addedCount++
            }
            elseif ($addResult.error.message -match 'already exist') {
                Write-Verbose "[$functionName] Already a member: $upn"
                [void]$results.Add(@{ upn = $upn; status = 'AlreadyMember'; message = "User '$upn' is already a member of '$($targetGroup.displayName)'." })
                $returnObject.alreadyMemberCount++
            }
            else {
                $errMsg = if ($addResult.error.message) { $addResult.error.message } else { "Unknown error (status $($addResult.statusCode))" }
                Write-Verbose "[$functionName] Failed to add '$upn': $errMsg"
                [void]$results.Add(@{ upn = $upn; status = 'Failed'; message = $errMsg })
                $returnObject.failedCount++
            }
        }
    }
    catch {
        Write-Error "[$functionName] Unexpected error: $_"
        $returnObject.failedCount = $UserPrincipalNames.Count - $returnObject.addedCount - $returnObject.alreadyMemberCount
    }

    $returnObject.success = ($returnObject.failedCount -eq 0)
    return $returnObject
}

