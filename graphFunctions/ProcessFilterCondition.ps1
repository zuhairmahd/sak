function ProcessFilterCondition()
{
    <#
    .SYNOPSIS
    Processes and encodes OData filter conditions for Microsoft Graph API queries.

    .DESCRIPTION
    This function parses OData filter condition strings and properly encodes values for use in
    Microsoft Graph API URLs. It handles both function-based filters (startswith, contains, endswith)
    and standard comparison operators (eq, ne, gt, lt, ge, le). The function properly encodes
    special characters and handles null/empty string values without encoding.

    .PARAMETER condition
    The filter condition string to process (e.g., "displayName eq 'Test'", "startswith(mail,'admin')").

    .OUTPUTS
    System.String
    Returns the processed and URL-encoded filter condition string ready for Graph API use.

    .EXAMPLE
    $filter = ProcessFilterCondition -condition "displayName eq 'John Doe'"
    # Returns: displayName eq 'John%20Doe'
    
    $filter = ProcessFilterCondition -condition "startswith(userPrincipalName,'admin')"
    # Returns: startswith(userPrincipalName,'admin')

    .NOTES
    Supported function-based filters:
    - startswith(property, value)
    - contains(property, value)
    - endswith(property, value)
    
    Supported comparison operators:
    - eq (equal)
    - ne (not equal)
    - gt (greater than)
    - lt (less than)
    - ge (greater than or equal)
    - le (less than or equal)
    
    Special handling for:
    - null values (not encoded)
    - Empty strings '' or "" (not encoded)
    - URL special characters (properly encoded)
    
    Removes quotes from values during processing.
    Compatible with PowerShell 5.1.
    #>
    [CmdletBinding()]
    param(
        [string]$condition
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Processing filter condition: $condition"
    # Check if this is a function-based filter (contains, startswith, endswith)
    Write-Verbose "[$functionName] Checking for function-based filter..."
    if ($condition -match '(startswith|contains|endswith)\s*\(([^,]+),\s*([^)]+)\)')
    {
        Write-Verbose "[$functionName] Found function-based filter: $($Matches[1])"
        $filterOperator = $Matches[1]
        Write-Verbose "[$functionName] Filter Operator: $filterOperator"
        $filterKey = $Matches[2].Trim()
        Write-Verbose "[$functionName] Filter Key: $filterKey"
        $filterValue = $Matches[3].Trim()
        Write-Verbose "[$functionName] Filter Value: $filterValue"
        # Remove quotes if present in the value
        Write-Verbose "[$functionName] Removing quotes from filter value..."
        $filterValue = $filterValue -replace "^'|'$", ""
        Write-Verbose "[$functionName] Filter Value after removing double quotes: $filterValue"
        $filterValue = $filterValue -replace '^"|"$', ""
        Write-Verbose "[$functionName] Filter Value after removing single quotes: $filterValue"
        Write-Verbose "[$functionName] Filter Key after removing quotes: $FilterKey"
        Write-Verbose "[$functionName] Filter Value: $FilterValue"
        Write-Verbose "[$functionName] Filter Operator: $FilterOperator"
        $encodedFilterValue = [uri]::EscapeDataString($FilterValue)
        Write-Verbose "[$functionName] Encoded Filter Value: $encodedFilterValue"
        # Rebuild the function call with encoded value
        $returnFilter = "$filterOperator($filterKey,'$encodedFilterValue')"
        Write-Verbose "[$functionName] Returning filter: $returnFilter"
        return $returnFilter
    }
    # Check for standard comparison operators
    elseif ($condition -match '([^\s]+)\s+(eq|ne|gt|lt|ge|le)\s+(.+)')
    {
        Write-Verbose "[$functionName] Not a function based filter. Checking for standard comparison operators..."
        $filterKey = $Matches[1].Trim()
        $filterOperator = $Matches[2].Trim()
        $filterValue = $Matches[3].Trim()
        Write-Verbose "[$functionName] Filter Key: $FilterKey"
        Write-Verbose "[$functionName] Filter Operator: $FilterOperator"
        Write-Verbose "[$functionName] Filter Value: $FilterValue"
        # Special handling for null and empty string
        Write-Verbose "[$functionName] Checking for null or empty string..."
        if ($filterValue -eq "null" -or $filterValue -eq "''" -or $filterValue -eq '""')
        {
            Write-Verbose "[$functionName] Filter value is null or empty string."
            Write-Verbose "[$functionName] Returning filter without encoding: $filterKey $filterOperator $filterValue"
            # Don't encode null or empty string values
            return "$filterKey $filterOperator $filterValue"
        }
        else
        {
            # Remove quotes if present
            Write-Verbose "[$functionName] Checking for quotes and removing from value if present..."
            Write-Verbose "[$functionName] Value before processing: $filterValue"
            $filterValue = $filterValue -replace "^'|'$", ""
            Write-Verbose "[$functionName] Value after removing double quotes: $filterValue"
            $filterValue = $filterValue -replace '^"|"$', ""
            Write-Verbose "[$functionName] Value after removing single quotes: $filterValue"
            Write-Verbose "[$functionName] Filter Key: $FilterKey"
            Write-Verbose "[$functionName] Filter Value: $FilterValue"
            $encodedFilterValue = [uri]::EscapeDataString($FilterValue)
            Write-Verbose "[$functionName] Encoded Filter Value: $encodedFilterValue"
            # Add quotes back for the encoded value
            $returnFilter = "$filterKey $filterOperator '$encodedFilterValue'"
            Write-Verbose "[$functionName] Returning filter: $returnFilter"
            return $returnFilter
        }
    }
    else
    {
        Write-Verbose "[$functionName] Unrecognized filter condition format: $condition"
        return $condition
    }
}

