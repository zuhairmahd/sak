function Get-RequiredScopesForEndpoints()
{
    <#
    .SYNOPSIS
        Maps Graph API endpoint URIs to their required permission scopes.

    .DESCRIPTION
        Extracts the endpoint-to-scope mapping logic from the deprecated HasScope function.
        Normalizes URIs by replacing dynamic segments (GUIDs, IDs, UPNs) with {id} placeholders
        and matches them against the configured scope definitions to determine which permissions
        are required for the given endpoints.

        This function focuses solely on endpoint-to-scope mapping, while Test-ScopeAvailability
        handles the actual token validation.

    .PARAMETER ResourcePaths
        Array of Graph API endpoint paths to check. Can be relative paths (e.g., 'users')
        or full URLs (e.g., 'https://graph.microsoft.com/v1.0/users/{id}').

    .PARAMETER RequiredScopes
        Array of scope definition objects from settings.psd1. Each object should have:
        - Scope: The permission scope name (e.g., 'User.Read.All')
        - Endpoints: Array of endpoint patterns this scope applies to
        - Reason: Description of why this scope is needed (optional)

    .OUTPUTS
        Array of scope definition objects that apply to the given endpoints.
        Returns empty array if no scopes are required (public endpoints).

    .EXAMPLE
        $endpointScopes = Get-RequiredScopesForEndpoints -ResourcePaths @('users', 'groups') `
            -RequiredScopes $settings.requiredScopes

        # Returns scope objects for User.Read.All and Group.Read.All

    .EXAMPLE
        $endpointScopes = Get-RequiredScopesForEndpoints `
            -ResourcePaths @('https://graph.microsoft.com/v1.0/users/12345-guid') `
            -RequiredScopes $settings.requiredScopes

        # Normalizes GUID to {id} pattern and matches against configured endpoints

    .NOTES
        This function was extracted from HasScope.ps1 as part of the scope validation refactoring.
        It preserves the URI normalization and pattern matching logic while simplifying the
        overall architecture by separating endpoint mapping from token validation.

        URI Normalization Patterns:
        - GUIDs: 8-4-4-4-12 format → {id}
        - Numeric IDs: Any number → {id}
        - UPNs: email-like patterns → {id}
        - Serial numbers: Alphanumeric 7-20 chars → {id}
        - Generic long IDs: Alphanumeric/special 10+ chars → {id}

    .LINK
        Test-ScopeAvailability - Validates token has required scopes
        Test-ScopeHierarchy - Checks hierarchical permission relationships
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ResourcePaths,

        [Parameter(Mandatory = $true)]
        [array]$RequiredScopes
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting endpoint-to-scope mapping for $($ResourcePaths.Count) resource paths"
    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Mapping $($ResourcePaths.Count) endpoints to required scopes" -LogLevel "Verbose"

    $allMatchingScopes = @()

    foreach ($uri in $ResourcePaths)
    {
        Write-Verbose "[$functionName] Processing resource path: $uri"
        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Processing resource path: $uri" -LogLevel "Verbose"

        # Normalize $uri to relative path for endpoint comparison
        if ($uri -match 'https?://[^/]+/(v1\.0|beta)/(.+)')
        {
            $relativeUri = $Matches[2]
            Write-Verbose "[$functionName] Extracted relative URI from full URL: $relativeUri"
        }
        else
        {
            $relativeUri = $uri
            Write-Verbose "[$functionName] Using URI as-is: $relativeUri"
        }

        # Strip query parameters if present
        if ($relativeUri -match '^([^?]+)')
        {
            $relativeUri = $Matches[1]
            Write-Verbose "[$functionName] Removed query parameters: $relativeUri"
        }

        # Strip leading slash if present for consistent comparison
        if ($relativeUri.StartsWith('/'))
        {
            $relativeUri = $relativeUri.Substring(1)
            Write-Verbose "[$functionName] Removed leading slash: $relativeUri"
        }

        # Normalize dynamic parameters by replacing GUIDs, IDs, and other dynamic segments
        $normalizedUri = $relativeUri

        # Replace GUID patterns (8-4-4-4-12 format) with {id}
        $normalizedUri = $normalizedUri -replace '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}', '{id}'

        # Replace other common ID patterns
        # Replace serial numbers (alphanumeric, typically 7-20 characters)
        $normalizedUri = $normalizedUri -replace '/[A-Za-z0-9]{7,20}(?=/|$)', '/{id}'

        # Replace numeric IDs
        $normalizedUri = $normalizedUri -replace '/\d+(?=/|$)', '/{id}'

        # Replace UPN patterns (email-like identifiers)
        $normalizedUri = $normalizedUri -replace '/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?=/|$)', '/{id}'

        # Handle special cases for common Microsoft Graph patterns
        # Replace any remaining dynamic segments that look like IDs but might not match above patterns
        $normalizedUri = $normalizedUri -replace '/[A-Za-z0-9_.-]{10,}(?=/|$)', '/{id}'

        Write-Verbose "[$functionName] Original URI: $relativeUri"
        Write-Verbose "[$functionName] Normalized URI for pattern matching: $normalizedUri"

        # Find required scopes for this endpoint
        # Normalize endpoints for comparison - handle leading slash variations and dynamic parameters
        [array]$matchingScopes = $RequiredScopes | Where-Object {
            $normalizedEndpoints = $_.endpoints | ForEach-Object {
                $endpoint = $_
                # Remove leading slash if present for comparison
                if ($endpoint.StartsWith('/'))
                {
                    $endpoint = $endpoint.Substring(1)
                }
                $endpoint
            }

            Write-Verbose "[$functionName]   Checking scope '$($_.Scope)' with normalized endpoints: $($normalizedEndpoints -join ', ')"

            # Check for exact matches first (for both original and normalized URIs)
            $exactMatch = $normalizedEndpoints -contains $relativeUri -or $normalizedEndpoints -contains $normalizedUri

            # Check for pattern matches using wildcards
            $patternMatch = $false
            foreach ($endpoint in $normalizedEndpoints)
            {
                # Convert endpoint patterns to PowerShell wildcard patterns
                $wildcardPattern = $endpoint.Replace('{id}', '*')

                # Test against both original and normalized URIs
                if ($relativeUri -like $wildcardPattern -or $normalizedUri -like $wildcardPattern)
                {
                    $patternMatch = $true
                    Write-Verbose "[$functionName]     Pattern match found: '$relativeUri' or '$normalizedUri' matches pattern '$wildcardPattern'"
                    break
                }
            }

            $exactMatch -or $patternMatch
        }

        Write-Verbose "[$functionName] Number of matching scopes for '$uri': $($matchingScopes.Count)"
        Write-Log -LogFile $LogFile -Module "$functionName" -Message "Found $($matchingScopes.Count) matching scopes for endpoint: $relativeUri" -LogLevel "Verbose"

        # Add matching scopes to the result set (avoiding duplicates)
        foreach ($scope in $matchingScopes)
        {
            $scopeName = if ($scope.Scope)
            {
                $scope.Scope 
            }
            else
            {
                $scope.scope 
            }

            # Check if we already added this scope
            $alreadyAdded = $allMatchingScopes | Where-Object {
                $existingName = if ($_.Scope)
                {
                    $_.Scope 
                }
                else
                {
                    $_.scope 
                }
                $existingName -eq $scopeName
            }

            if (-not $alreadyAdded)
            {
                $allMatchingScopes += $scope
                Write-Verbose "[$functionName] Added scope: $scopeName"
            }
            else
            {
                Write-Verbose "[$functionName] Scope already added (skipping duplicate): $scopeName"
            }
        }
    }

    Write-Verbose "[$functionName] Total unique scopes required: $($allMatchingScopes.Count)"
    Write-Log -LogFile $LogFile -Module "$functionName" -Message "Mapped $($ResourcePaths.Count) endpoints to $($allMatchingScopes.Count) unique scopes" -LogLevel "Information"

    # Force return as array (PowerShell unwraps single-item arrays to scalars)
    return , @($allMatchingScopes)
}
