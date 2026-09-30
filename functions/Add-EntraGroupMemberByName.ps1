function Add-EntraGroupMemberByName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string]$GroupDisplayName,
        [Parameter(Mandatory = $true)]
        [string]$AccessToken
    )

    $functionName = $MyInvocation.MyCommand.Name
    $returnObject = @{
        success = $false
        message = ""
    }
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    try {
        # 1. Resolve User Object ID (Handles exact UPN or Mail matches)
        $userFilter = [uri]::EscapeDataString("userPrincipalName eq '$UserPrincipalName' or mail eq '$UserPrincipalName'")
        $userUri = "https://graph.microsoft.com/v1.0/users?`$filter=$userFilter&`$select=id,userPrincipalName,displayName"

        $userResponse = Invoke-RestMethod -Method Get -Uri $userUri -Headers $headers
        if (-not $userResponse.value -or $userResponse.value.Count -eq 0) {
            $returnObject.message = "User '$UserPrincipalName' was not found in Entra ID."
            throw "User '$UserPrincipalName' was not found in Entra ID."
        }
        $targetUser = $userResponse.value[0]
        Write-Verbose "[$functionName] Resolved User: $($targetUser.displayName) ($($targetUser.id))"

        # 2. Resolve Group Object ID
        $groupFilter = [uri]::EscapeDataString("displayName eq '$GroupDisplayName'")
        $groupUri = "https://graph.microsoft.com/v1.0/groups?`$filter=$groupFilter&`$select=id,displayName,groupTypes"

        $groupResponse = Invoke-RestMethod -Method Get -Uri $groupUri -Headers $headers
        if (-not $groupResponse.value -or $groupResponse.value.Count -eq 0) {
            $returnObject.message = "Group '$GroupDisplayName' was not found in Entra ID."
            throw "Group '$GroupDisplayName' was not found in Entra ID."
        }
        if ($groupResponse.value.Count -gt 1) {
            $returnObject.message = "Multiple groups found with the name '$GroupDisplayName'. Please target via Group ID to prevent ambiguity."
            throw "Multiple groups found with the name '$GroupDisplayName'. Please target via Group ID to prevent ambiguity."
        }
        $targetGroup = $groupResponse.value[0]
        Write-Verbose "[$functionName]              Resolved Group: $($targetGroup.displayName) ($($targetGroup.id))"

        # 3. Guard: Dynamic groups cannot be manually modified
        if ($targetGroup.groupTypes -contains "DynamicMembership") {
            $returnObject.message = "The group '$GroupDisplayName' is a Dynamic Group. Members cannot be added manually."
            throw "The group '$GroupDisplayName' is a Dynamic Group. Members cannot be added manually."
        }

        # 4. Check if the user is already a member
        $checkMemberUri = "https://graph.microsoft.com/v1.0/groups/$($targetGroup.id)/members/`$ref?`$filter=id eq '$($targetUser.id)'"
        $membershipCheck = Invoke-RestMethod -Method Get -Uri $checkMemberUri -Headers $headers
        if ($membershipCheck.value.Count -gt 0) {
            Write-Host "User '$($targetUser.userPrincipalName)' is already a member of '$($targetGroup.displayName)'." -ForegroundColor Yellow
            $returnObject.success = $true
            $returnObject.message = "User '$($targetUser.userPrincipalName)' is already a member of '$($targetGroup.displayName)'."
            return $returnObject
        }

        # 5. Add user to the group
        $addMemberUri = "https://graph.microsoft.com/v1.0/groups/$($targetGroup.id)/members/`$ref"
        $body = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($targetUser.id)"
        } | ConvertTo-Json -Compress

        $response = Invoke-RestMethod -Method Post -Uri $addMemberUri -Headers $headers -Body$body
        if ($response -in 200..299) {
            Write-Host "Successfully added '$($targetUser.userPrincipalName)' to '$($targetGroup.displayName)'." -ForegroundColor Green
            $returnObject.success = $true
            $returnObject.message = "Successfully added '$($targetUser.userPrincipalName)' to '$($targetGroup.displayName)'."
        }
        else {
            Write-Error "Failed to add '$($targetUser.userPrincipalName)' to '$($targetGroup.displayName)'."
            $returnObject.success = $false
            $returnObject.message = "Failed to add '$($targetUser.userPrincipalName)' to '$($targetGroup.displayName).`n Got response: $response'        "
        }
    }
    catch {
        Write-Error "Failed to add member to group: $_"
        $returnObject.success = $false
        $returnObject.message = "Failed to add member to group: $_"
    }
    return $returnObject
}