function Invoke-GraphAPI {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$accessToken,
        [Parameter(Mandatory = $true)]
        [object]$ResourcePath,  # Can be string or string array for batch processing
        [string]$APIVersion = 'beta',
        [string]$method = 'get',
        [string]$Filter = $null,
        [string]$Search = $null,
        [string]$ExtraParameters = $null,
        $headers,
        [string]$body = $null,
        [switch]$consistencyLevel,
        [switch]$secureString
    )

    function ProcessFilterCondition {
        [CmdletBinding()]
        param(
            [string]$condition
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Processing filter condition: $condition"
        # Check if this is a function-based filter (contains, startswith, endswith)
        Write-Verbose "[$functionName] Checking for function-based filter..."
        if ($condition -match '(startswith|contains|endswith)\s*\(([^,]+),\s*([^)]+)\)') {
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
        elseif ($condition -match '([^\s]+)\s+(eq|ne|gt|lt|ge|le)\s+(.+)') {
            Write-Verbose "[$functionName] Not a function based filter. Checking for standard comparison operators..."
            $filterKey = $Matches[1].Trim()
            $filterOperator = $Matches[2].Trim()
            $filterValue = $Matches[3].Trim()
            Write-Verbose "[$functionName] Filter Key: $FilterKey"
            Write-Verbose "[$functionName] Filter Operator: $FilterOperator"
            Write-Verbose "[$functionName] Filter Value: $FilterValue"
            # Special handling for null and empty string
            Write-Verbose "[$functionName] Checking for null or empty string..."
            if ($filterValue -eq "null" -or $filterValue -eq "''" -or $filterValue -eq '""') {
                Write-Verbose "[$functionName] Filter value is null or empty string."
                Write-Verbose "[$functionName] Returning filter without encoding: $filterKey $filterOperator $filterValue"
                # Don't encode null or empty string values
                return "$filterKey $filterOperator $filterValue"
            }
            else {
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
        else {
            Write-Verbose "[$functionName] Unrecognized filter condition format: $condition"
            return $condition
        }
    }

    #region variables and logs
    $functionName = $MyInvocation.MyCommand.Name
    if ($accessToken) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Access token provided." -LogLevel "Information"
        Write-Verbose "[$functionName] Access token provided."
    }
    else {
        Write-Verbose "[$functionName] Access token not provided. Please provide a valid access token."
        Write-Log -LogFile $logFile -Module $functionName -Message "Access token not provided." -LogLevel "Error"
        return
    }
    Write-Log -LogFile $logFile -Module $functionName -Message "Resource Path: $ResourcePath" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Method: $method" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Filter: $filter" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Search: $Search" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Extra Parameters: $ExtraParameters" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Version: $APIVersion" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Consistency Level: $consistencyLevel" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "Body: $body" -LogLevel "Information"
    Write-Log -LogFile $logFile -Module $functionName -Message "SecureString: $secureString" -LogLevel "Information"
    #write-verbose all the above.
    Write-Verbose "[$functionName] Resource Path: $ResourcePath"
    Write-Verbose "[$functionName] Method: $method"
    Write-Verbose "[$functionName] Filter: $filter"
    Write-Verbose "[$functionName] Search: $Search"
    Write-Verbose "[$functionName] Extra Parameters: $ExtraParameters"
    Write-Verbose "[$functionName] API Version: $APIVersion"
    Write-Verbose "[$functionName] Consistency Level: $consistencyLevel"
    Write-Verbose "[$functionName] Body: $body"
    Write-Verbose "[$functionName] SecureString: $secureString"

    # Check if ResourcePath is an array
    $isArrayInput = $ResourcePath -is [array]
    Write-Verbose "[$functionName] isArrayInput: $isArrayInput"
    Write-Log -logFile $logFile -Module $functionName -Message "Function called with ResourcePath type: $($ResourcePath.GetType().FullName)" -LogLevel "Information"
    # Handle single-item array
    if ($isArrayInput -and $ResourcePath.Count -eq 1) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Single-item array detected, processing as single request" -LogLevel "Verbose"
        Write-Verbose "[$functionName] Single-item array detected, processing as single request"
        $ResourcePath = $ResourcePath[0]
        $isArrayInput = $false
    }
    # Check if batch processing is requested (array with multiple items)
    $isBatchRequest = $isArrayInput -and $ResourcePath.Count -gt 1
    $batchThreshold = 1
    Write-Verbose "[$functionName] isBatchRequest: $isBatchRequest with a threshold of $batchThreshold"
    Write-Log -logFile $logFile -Module $functionName -Message "isBatchRequest: $isBatchRequest with a threshold of $batchThreshold" -LogLevel "Information"
    if ($isBatchRequest -and $ResourcePath.Count -ge $batchThreshold) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Batch request detected: $($ResourcePath.Count) resources" -LogLevel "Information"
        Write-Verbose "[$functionName] Batch request detected: $($ResourcePath.Count) resources"
        # Attempt to use native Graph API $batch endpoint
        # Graph API supports up to 20 requests per batch
        $maxBatchSize = 20
        $allResults = @()
        $successCount = 0
        $failureCount = 0
        # Split requests into batches of max 20
        $batches = @()
        for ($i = 0; $i -lt $ResourcePath.Count; $i += $maxBatchSize) {
            $batchSize = [Math]::Min($maxBatchSize, $ResourcePath.Count - $i)
            $batches += , @($ResourcePath[$i..($i + $batchSize - 1)])
        }
        Write-Log -LogFile $logFile -Module $functionName -Message "Processing $($ResourcePath.Count) requests in $($batches.Count) batch(es)" -LogLevel "Information"
        Write-Verbose "[$functionName] Processing $($ResourcePath.Count) requests in $($batches.Count) batch(es)"
        $batchIndex = 0
        foreach ($batch in $batches) {
            # Build batch request body according to Graph API spec
            $batchRequests = @()
            $requestId = 1
            foreach ($path in $batch) {
                # Build full URL for the request
                $requestUrl = "/$path"
                # Handle filters, search, and extra parameters in the URL
                $queryParams = @()
                if ($Filter) {
                    $queryParams += "`$filter=$([uri]::EscapeUriString($Filter))"
                }
                if ($Search) {
                    $queryParams += "`$search=$([uri]::EscapeUriString($Search))"
                }
                if ($ExtraParameters) {
                    $queryParams += $ExtraParameters
                }
                if ($queryParams.Count -gt 0) {
                    $requestUrl += "?" + ($queryParams -join "&")
                }
                # Build request object
                $batchRequest = @{
                    id     = $requestId.ToString()
                    method = $method.ToUpper()
                    url    = $requestUrl
                }
                # Add headers if needed
                if ($consistencyLevel) {
                    $batchRequest['headers'] = @{
                        'ConsistencyLevel' = 'eventual'
                    }
                }
                # Add body if provided
                if ($body) {
                    $batchRequest['body'] = $body | ConvertFrom-Json
                    if (-not $batchRequest.ContainsKey('headers')) {
                        $batchRequest['headers'] = @{}
                    }
                    $batchRequest['headers']['Content-Type'] = 'application/json'
                }
                $batchRequests += $batchRequest
                $requestId++
            }
            # Create batch request body
            $batchBody = @{
                requests = $batchRequests
            } | ConvertTo-Json -Depth 10
            Write-Log -LogFile $logFile -Module $functionName -Message "Sending batch with $($batchRequests.Count) requests to `$batch endpoint" -LogLevel "Verbose"
            # Send batch request to Graph API
            try {
                $batchHeaders = @{
                    'Authorization' = "Bearer $accessToken"
                    'Content-Type'  = 'application/json'
                }
                $batchUri = "https://graph.microsoft.com/$APIVersion/`$batch"
                $batchResponse = Invoke-RestMethod -Uri $batchUri -Method Post -Headers $batchHeaders -Body $batchBody -UseBasicParsing
                # Process batch responses
                # Renumber response IDs to be globally unique across all batches
                $globalIdOffset = $batchIndex * $maxBatchSize
                foreach ($response in $batchResponse.responses) {
                    # Adjust the response ID to be globally unique (1-240 instead of 1-20 per batch)
                    $globalId = ([int]$response.id) + $globalIdOffset
                    $response.id = $globalId

                    if ($response.status -ge 200 -and $response.status -lt 300) {
                        # Preserve the entire response object so downstream code can match by id
                        $allResults += $response
                        $successCount++
                        Write-Log -LogFile $logFile -Module $functionName -Message "Batch request $($response.id) succeeded (status: $($response.status))" -LogLevel "Verbose"
                    }
                    else {
                        # Include failed responses so downstream code can handle them properly
                        $allResults += $response
                        $failureCount++
                        $errorMsg = if ($response.body.error) { $response.body.error.message } else { "Unknown error" }
                        Write-Log -LogFile $logFile -Module $functionName -Message "Batch request $($response.id) failed (status: $($response.status)): $errorMsg" -LogLevel "Warning"
                    }
                }
                $batchIndex++
            }
            catch {
                Write-Log -LogFile $logFile -Module $functionName -Message "Batch endpoint failed: $($_.Exception.Message). Falling back to sequential processing." -LogLevel "Warning"
                # Final fallback: process each resource path individually
                foreach ($path in $batch) {
                    Write-Log -LogFile $logFile -Module $functionName -Message "Processing resource sequentially: $path" -LogLevel "Verbose"
                    # Recursive call with single resource path
                    $result = CallGraphAPI -accessToken $accessToken -ResourcePath $path -APIVersion $APIVersion `
                        -method $method -Filter $Filter -Search $Search -ExtraParameters $ExtraParameters `
                        -body $body -consistencyLevel:$consistencyLevel -secureString:$secureString
                    if ($null -eq $result -or $result.error -ne $null) {
                        $failureCount++
                        Write-Log -LogFile $logFile -Module $functionName -Message "Failed to process resource: $path (Status: $result)" -LogLevel "Warning"
                    }
                    else {
                        $allResults += $result
                        $successCount++
                    }
                }
            }
        }
        Write-Log -LogFile $logFile -Module $functionName -Message "Batch processing completed: $successCount successful, $failureCount failed" -LogLevel "Information"
        Write-Verbose "[$functionName] Batch processing completed: $successCount successful, $failureCount failed"
        # Return combined results
        return @{
            value          = $allResults
            batchProcessed = $true
            batchMethod    = if ($useBatchProcessor) { "GraphCore" } else { "NativeBatch" }
            successCount   = $successCount
            failureCount   = $failureCount
            totalCount     = $ResourcePath.Count
        }
    }

    # Single request processing (original behavior continues below)
    $uri = "https://graph.microsoft.com/$APIVersion/$ResourcePath"
    $statusCode = $null
    Write-Log -LogFile $logFile -Module $functionName -Message "Uri: $uri" -LogLevel "Information"
    Write-Verbose "[$functionName] Uri: $uri"
    #endregion

    #region Encode filter and add headers
    if ($Filter) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Processing filter string: $Filter" -LogLevel "Verbose"
        Write-Log -LogFile $logFile -Module $functionName -Message "Splitting filter by logical operators while preserving operators." -LogLevel "Information"
        Write-Verbose "[$functionName] Splitting filter by logical operators while preserving operators."
        $filterParts = [System.Collections.ArrayList]::new()
        $logicalOperators = [System.Collections.ArrayList]::new()
        # Pattern to match a logical operator with surrounding spaces
        $pattern = '\s+(and|or)\s+'
        $lastIndex = 0
        # Find all logical operators and their positions
        $logicalOperaterMatches = [regex]::Matches($Filter, $pattern)
        Write-Log -LogFile $logFile -Module $functionName -Message "Found $($logicalOperaterMatches.Count) logical operators." -LogLevel "Verbose"
        Write-Verbose "[$functionName] Found $($logicalOperaterMatches.Count) logical operators."
        # If no logical operators, process as a single condition
        if ($logicalOperaterMatches.Count -eq 0) {
            Write-Log -LogFile $logFile -Module $functionName -Message "No logical operators found. Processing as a single filter condition." -LogLevel "Verbose"
            Write-Verbose "[$functionName] No logical operators found. Processing as a single filter condition."
            $processedFilter = ProcessFilterCondition -condition $Filter
            Write-Log -LogFile $logFile -Module $functionName -Message "Processed single filter condition: $processedFilter" -LogLevel "Information"
            Write-Verbose "[$functionName] Processed single filter condition: $processedFilter"
            $encodedFilter = $processedFilter
            Write-Log -LogFile $logFile -Module $functionName -Message "Encoded filter: $encodedFilter" -LogLevel "Information"
            Write-Verbose "[$functionName] Encoded filter: $encodedFilter"
        }
        else {
            # Process each part of the filter
            Write-Log -LogFile $logFile -Module $functionName -Message "Logical operators found. Processing filter as multiple conditions." -LogLevel "Verbose"
            Write-Verbose "[$functionName] Logical operators found. Processing filter as multiple conditions."
            foreach ($logicalOperatorMatch in $logicalOperaterMatches) {
                Write-Log -LogFile $logFile -Module $functionName -Message "Processing filter condition before logical operator: $($Filter.Substring($lastIndex, $logicalOperatorMatch.Index - $lastIndex))" -LogLevel "Debug"
                Write-Verbose "[$functionName] Processing filter condition before logical operator: $($Filter.Substring($lastIndex, $logicalOperatorMatch.Index - $lastIndex))"
                $condition = $Filter.Substring($lastIndex, $logicalOperatorMatch.Index - $lastIndex)
                Write-Log -LogFile $logFile -Module $functionName -Message "Condition to process: $condition" -LogLevel "Information"
                Write-Verbose "[$functionName] Condition to process: $condition"
                [void]$filterParts.Add((ProcessFilterCondition -condition $condition))
                Write-Log -LogFile $logFile -Module $functionName -Message "Processed filter condition: $($filterParts[$filterParts.Count - 1])" -LogLevel "Information"
                Write-Verbose "[$functionName] Processed filter condition: $($filterParts[$filterParts.Count - 1])"
                # Store the logical operator (and, or)
                [void]$logicalOperators.Add($logicalOperatorMatch.Value.Trim())
                $lastIndex = $logicalOperatorMatch.Index + $logicalOperatorMatch.Length
                Write-Log -LogFile $logFile -Module $functionName -Message "Logical operators so far: $($logicalOperators -join ', ')" -LogLevel "Information"
                Write-Verbose "[$functionName] Logical operators so far: $($logicalOperators -join ', ')"
            }
            # Don't forget the last part after the last logical operator
            if ($lastIndex -lt $Filter.Length) {
                Write-Log -LogFile $logFile -Module $functionName -Message "Processing filter condition after the last logical operator." -LogLevel "Verbose"
                Write-Verbose "[$functionName] Processing filter condition after the last logical operator."
                $condition = $Filter.Substring($lastIndex)
                [void]$filterParts.Add((ProcessFilterCondition -condition $condition))
                Write-Log -LogFile $logFile -Module $functionName -Message "Processed filter condition: $($filterParts[$filterParts.Count - 1])" -LogLevel "Information"
                Write-Verbose "[$functionName] Processed filter condition: $($filterParts[$filterParts.Count - 1])"
            }
            # Rebuild the filter string with processed parts and original logical operators
            Write-Log -LogFile $logFile -Module $functionName -Message "Rebuilding the filter string with processed parts and logical operators." -LogLevel "Information"
            Write-Verbose "[$functionName] Rebuilding the filter string with processed parts and logical operators."
            $encodedFilter = $filterParts[0]
            for ($i = 0; $i -lt $logicalOperators.Count; $i++) {
                $encodedFilter += " $($logicalOperators[$i]) $($filterParts[$i+1])"
                Write-Log -LogFile $logFile -Module $functionName -Message "Adding logical operator: $($logicalOperators[$i])" -LogLevel "Information"
                Write-Verbose "[$functionName] Adding logical operator: $($logicalOperators[$i])"
            }
            Write-Log -LogFile $logFile -Module $functionName -Message "Processed complex filter: $encodedFilter" -LogLevel "Information"
            Write-Verbose "[$functionName] Processed complex filter: $encodedFilter"
        }
        $encodedUri = "$uri`?`$filter=$([uri]::EscapeUriString($encodedFilter))"
        Write-Log -LogFile $logFile -Module $functionName -Message "Uri after applying filters: $encodedUri" -LogLevel "Information"
        Write-Verbose "[$functionName] Uri after applying filters: $encodedUri"
    }
    else {
        Write-Log -LogFile $logFile -Module $functionName -Message "No filter provided." -LogLevel "Information"
        Write-Verbose "[$functionName] No filter provided."
        $encodedUri = $uri
    }

    # Handle search parameter
    if ($Search) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Processing search parameter: $Search" -LogLevel "Verbose"
        Write-Verbose "[$functionName] Processing search parameter: $Search"
        # URL encode the search string
        $encodedSearch = [uri]::EscapeUriString($Search)
        Write-Log -LogFile $logFile -Module $functionName -Message "Encoded search: $encodedSearch" -LogLevel "Information"
        Write-Verbose "[$functionName] Encoded search: $encodedSearch"
        # Add search parameter to URI
        if ($encodedUri.Contains("?")) {
            $encodedUri = "$encodedUri&`$search=$encodedSearch"
        }
        else {
            $encodedUri = "$encodedUri`?`$search=$encodedSearch"
        }
        Write-Log -LogFile $logFile -Module $functionName -Message "Uri after applying search: $encodedUri" -LogLevel "Information"
        Write-Verbose "[$functionName] Uri after applying search: $encodedUri"
    }
    else {
        Write-Log -LogFile $logFile -Module $functionName -Message "No search parameter provided." -LogLevel "Information"
        Write-Verbose "[$functionName] No search parameter provided."
    }

    if ($extraParameters) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Extra parameters provided." -LogLevel "Information"
        Write-Log -LogFile $logFile -Module $functionName -Message "Splitting the extra parameters by ampersand to get individual key-value pairs." -LogLevel "Information"
        Write-Verbose "[$functionName] Extra parameters provided."
        # Initialize the parameter list
        $paramsList = @()
        # Split by ampersand to get individual key-value pairs
        $keyValuePairs = $extraParameters -split '&'
        Write-Log -LogFile $logFile -Module $functionName -Message "Found $($keyValuePairs.Count) key-value pairs." -LogLevel "Verbose"
        Write-Verbose "[$functionName] Found $($keyValuePairs.Count) key-value pairs."
        foreach ($pair in $keyValuePairs) {
            Write-Log -LogFile $logFile -Module $functionName -Message "Processing key-value pair: $pair" -LogLevel "Verbose"
            Write-Verbose "[$functionName] Processing key-value pair: $pair"
            # Split each pair by equals sign to separate key and value
            $keyAndValue = $pair -split '=', 2
            if ($keyAndValue.Count -eq 2) {
                $key = $keyAndValue[0].Trim()
                $value = $keyAndValue[1].Trim()
                Write-Log -LogFile $logFile -Module $functionName -Message "Key: $key" -LogLevel "Information"
                Write-Log -LogFile $logFile -Module $functionName -Message "Value: $value" -LogLevel "Information"
                Write-Verbose "[$functionName] Key: $key"
                Write-Verbose "[$functionName] Value: $value"
                # Add the $ prefix to the key for OData parameters
                $formattedKey = "`$$key"
                Write-Log -LogFile $logFile -Module $functionName -Message "Formatted Key with $ prefix: $formattedKey" -LogLevel "Information"
                Write-Verbose "[$functionName] Formatted Key with $ prefix: $formattedKey"
                # Add the formatted parameter to the list
                $paramsList += "$formattedKey=$value"
            }
            else {
                Write-Warning "Invalid parameter format: $pair - skipping"
                Write-Log -LogFile $logFile -Module $functionName -Message "Invalid parameter format: $pair - skipping" -LogLevel "Warning"
            }
        }
        Write-Log -LogFile $logFile -Module $functionName -Message "Final parameter list:" -LogLevel "Information"
        Write-Verbose "[$functionName] Final parameter list:"
        $paramsList | ForEach-Object { Write-Verbose "[$functionName] $_" }
        # Join the parameters with & to create a complete query string
        $queryString = $paramsList -join '&'
        Write-Log -LogFile $logFile -Module $functionName -Message "Final query string: $queryString" -LogLevel "Information"
        Write-Verbose "[$functionName] Final query string: $queryString"
        # Append the extra parameters to the URI
        if ($filter -or $Search) {
            Write-Log -LogFile $logFile -Module $functionName -Message "Adding extra parameters to the uri along with existing parameters." -LogLevel "Information"
            Write-Verbose "[$functionName] Adding extra parameters to the uri along with existing parameters."
            $encodedUri = "$encodedUri`&$queryString"
        }
        else {
            Write-Log -LogFile $logFile -Module $functionName -Message "No filter or search provided. Adding extra parameters to the uri." -LogLevel "Information"
            Write-Verbose "[$functionName] No filter or search provided. Adding extra parameters to the uri."
            $encodedUri = "$encodedUri`?$queryString"
        }
    }
    else {
        Write-Log -LogFile $logFile -Module $functionName -Message "No extra parameters provided." -LogLevel "Information"
        Write-Verbose "[$functionName] No extra parameters provided."
    }
    # Build default headers with Authorization and Content-Type
    if ($consistencyLevel) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Adding consistency level to the headers." -LogLevel "Information"
        Write-Verbose "[$functionName] Adding consistency level to the headers."
        $defaultHeaders = @{
            Authorization    = "Bearer $accessToken"
            'Content-Type'   = 'application/json'
            ConsistencyLevel = 'Eventual'
        }
    }
    else {
        Write-Log -LogFile $logFile -Module $functionName -Message "No consistency level provided." -LogLevel "Information"
        Write-Verbose "[$functionName] No consistency level provided."
        $defaultHeaders = @{
            Authorization  = "Bearer $accessToken"
            'Content-Type' = 'application/json'
        }
    }

    # Merge custom headers if provided (custom headers take precedence)
    if ($headers) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Custom headers provided. Merging with default headers." -LogLevel "Information"
        Write-Verbose "[$functionName] Custom headers provided. Merging with default headers."
        foreach ($key in $headers.Keys) {
            $defaultHeaders[$key] = $headers[$key]
            Write-Log -LogFile $logFile -Module $functionName -Message "Added/Overridden header: $key" -LogLevel "Information"
            Write-Verbose "[$functionName] Added/Overridden header: $key"
        }
    }
    #endregion

    #region prepare the call
    # Create parameter hashtable for splatting
    Write-Log -LogFile $logFile -Module $functionName -Message "Preparing parameters for Invoke-RestMethod call." -LogLevel "Information"
    Write-Verbose "[$functionName] Preparing parameters for Invoke-RestMethod call."
    $restParams = @{
        Method          = $method
        Uri             = $encodedUri
        Headers         = $defaultHeaders
        UseBasicParsing = $true
    }
    #add headers parameter if it was passed
    if ($headers) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Headers provided. Adding to the request." -LogLevel "Information"
        Write-Verbose "[$functionName] Headers provided. Adding to the request."
        $restParams['Headers'] = $headers
    }
    # Only add Body parameter if it exists
    if ($body) {
        Write-Log -LogFile $logFile -Module $functionName -Message "Body parameter provided. Adding to the request." -LogLevel "Information"
        Write-Verbose "[$functionName] Body parameter provided. Adding to the request."
        $restParams['Body'] = $body
    }
    Write-Log -LogFile $logFile -Module $functionName -Message "Making the following call to Microsoft Graph:" -LogLevel "Information"
    Write-Verbose "[$functionName] Making the following call to Microsoft Graph:"
    Write-Log -LogFile $logFile -Module $functionName -Message "URI: $encodedUri." -LogLevel "Information"
    Write-Verbose "[$functionName] URI: $encodedUri"
    Write-Log -LogFile $logFile -Module $functionName -Message "Method: $method." -LogLevel "Information"
    Write-Verbose "[$functionName] Method: $method"
    #endregion
    try {
        $response = Invoke-RestMethod @restParams
        Write-Log -LogFile $logFile -Module $functionName -Message "NextLink: $($response.'@odata.nextLink')" -LogLevel "Information"
        Write-Verbose "[$functionName] NextLink: $($response.'@odata.nextLink')"
        Write-Log -LogFile $logFile -Module $functionName -Message "Response count: $($response.value.count)" -LogLevel "Information"
        Write-Verbose "[$functionName] Response count: $($response.value.count)"
        if ($response.'@odata.nextLink') {
            Write-Log -LogFile $logFile -Module $functionName -Message "NextLink found. Fetching additional pages." -LogLevel "Verbose"
            Write-Verbose "[$functionName] NextLink found. Fetching additional pages."
            # Initialize an array to hold all items
            $allItems = @()
            $allItems += $response.value
            $nextLink = $response.'@odata.nextLink'
            while ($nextLink) {
                $nextGroup = Invoke-RestMethod -Method $method -Uri $nextLink -Headers $defaultHeaders -UseBasicParsing
                Write-Log -LogFile $logFile -Module $functionName -Message "Fetched next page with $($nextGroup.value.Count) items." -LogLevel "Information"
                Write-Verbose "[$functionName] Fetched next page with $($nextGroup.value.Count) items."
                if ($nextGroup.value) {
                    Write-Log -LogFile $logFile -Module $functionName -Message "Adding items from next page to the collection." -LogLevel "Information"
                    Write-Verbose "[$functionName] Adding items from next page to the collection."
                    $allItems += $nextGroup.value
                }
                $nextLink = $nextGroup.'@odata.nextLink'
            }
            # Optionally, reconstruct a response object if needed
            $response.value = $allItems
            Write-Log -LogFile $logFile -Module $functionName -Message "All items collected. Total count: $($Response.value.Count)" -LogLevel "Information"
            Write-Verbose "[$functionName] All items collected. Total count: $($Response.value.Count)"
        }
        else {
            Write-Log -LogFile $logFile -Module $functionName -Message "No nextLink found. Single page response received." -LogLevel "Verbose"
            Write-Verbose "[$functionName] No nextLink found. Single page response received."
        }
        Write-Log -LogFile $logFile -Module $functionName -Message "The call was successful." -LogLevel "Information"
        Write-Verbose "[$functionName] The call was successful."
        #Add a status code
        $statusCode = if ($response -and $response.statusCode) { $response.statusCode } else { 200 }
        $response | Add-Member -NotePropertyName 'statusCode' -NotePropertyValue $statusCode -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log -LogFile $logFile -Module $functionName -Message "Graph API call failed: $($PSItem.Exception.Message)" -LogLevel "Error"

        # Extract HTTP status code (PS5.1: .statuscode enum; PS7: same; non-HTTP: parse from message)
        if ($null -eq $PSItem.Exception.statusCode) {
            $statusCode = [regex]::Match($PSItem.Exception.Message, '\d+').Value
        }
        else {
            try { $statusCode = $PSItem.Exception.statuscode.value__ }
            catch { $statusCode = [int]$PSItem.Exception.statuscode }
        }

        # Extract response body (PS5.1: response stream; PS7: ErrorDetails.Message)
        $responseBodyRaw = $null
        $resp = $PSItem.Exception.Response
        if ($null -ne $resp -and $resp -is [System.Net.HttpWebResponse]) {
            try {
                $stream = $resp.GetResponseStream()
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $responseBodyRaw = $reader.ReadToEnd()
                    $reader.Close()
                }
            }
            catch {}
        }
        if (-not $responseBodyRaw -and $PSItem.ErrorDetails) {
            $responseBodyRaw = $PSItem.ErrorDetails.Message
        }

        $response = $responseBodyRaw | ConvertFrom-Json -ErrorAction SilentlyContinue
        if (-not $response) {
            $response = [PSCustomObject]@{ error = [PSCustomObject]@{ message = $PSItem.Exception.Message } }
        }
        $response | Add-Member -NotePropertyName 'statusCode' -NotePropertyValue $statusCode -Force -ErrorAction SilentlyContinue
        return $response
    }
    Write-Log -Message "Response: $($response)" -LogFile $logFile -Module $functionName -LogLevel Information -CMTraceFormat:$false -ErrorAction SilentlyContinue
    Write-Log -Message "Response value: $($response.value)" -LogFile $logFile -Module $functionName -LogLevel Information -CMTraceFormat:$false -ErrorAction SilentlyContinue
    return $response
}

