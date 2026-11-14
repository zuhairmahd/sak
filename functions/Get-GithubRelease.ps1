function Get-GithubRelease {
    <#
    .SYNOPSIS
        Downloads a release asset from a public or private GitHub repository.

    .DESCRIPTION
        Looks up a release (latest or by tag), filters its assets with a wildcard pattern and downloads
        the selected asset. When several assets match, the user chooses one from a menu, unless
        -NonInteractive is set, in which case the function fails instead of guessing.
        For private repositories, specify -Private together with -AccessToken. The token is sent as a
        Bearer header and is never written to the log or the console.

    .PARAMETER owner
        GitHub user or organisation that owns the repository.

    .PARAMETER repo
        Repository name.

    .PARAMETER tag
        Release tag to download from. Defaults to "latest".

    .PARAMETER pattern
        Wildcard pattern (-like) matched against asset names. A pattern with no wildcard characters
        (*, ?, [) is treated as "contains", so "msi" is the same as "*msi*". Defaults to "*linux_amd64*".

    .PARAMETER OutputDirectory
        Folder to save the asset in. Created if missing. Defaults to the current directory.

    .PARAMETER NonInteractive
        Never prompt. If more than one asset matches, return -1 without downloading.

    .PARAMETER Private
        The repository is private. Requires -AccessToken.

    .PARAMETER AccessToken
        GitHub personal access token as a SecureString. Required with -Private, optional otherwise
        (raises the API rate limit).

    .OUTPUTS
        System.Int32
        200 = downloaded, 0 = no match or user cancelled, -1 = several matches in non-interactive mode,
        -2 = failure without an HTTP status, any other value = HTTP status code of the failure.

    .EXAMPLE
        Get-GithubRelease -owner PowerShell -repo PowerShell -pattern "*win-x64.msi"

    .EXAMPLE
        $token = Read-Host -AsSecureString
        Get-GithubRelease -owner contoso -repo tools -Private -AccessToken $token -NonInteractive -pattern "*.zip"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[A-Za-z0-9](?:[A-Za-z0-9-]{0,38})$')]
        [string]$owner,
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[A-Za-z0-9._-]{1,100}$')]
        [string]$repo,
        [ValidateNotNullOrEmpty()]
        [string]$tag = "latest",
        [ValidateNotNullOrEmpty()]
        [string]$pattern = "*linux_amd64*",
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory = $PWD.Path,
        [switch]$NonInteractive,
        [switch]$Private,
        [securestring]$AccessToken
    )

    # Streams a response to disk and reports progress. Removes the partial file if the copy fails.
    function Save-StreamWithProgress {
        param(
            [Parameter(Mandatory = $true)][System.IO.Stream]$Source,
            [Parameter(Mandatory = $true)][string]$Path,
            [long]$TotalBytes,
            [string]$Activity = 'Downloading'
        )
        $fileStream = [System.IO.File]::Create($Path)
        $succeeded = $false
        try {
            $buffer = New-Object byte[] 81920
            $copied = 0L
            $timer = [System.Diagnostics.Stopwatch]::StartNew()
            $lastUpdate = -1000
            while (($read = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fileStream.Write($buffer, 0, $read)
                $copied += $read
                if ($timer.ElapsedMilliseconds - $lastUpdate -ge 200) {
                    $lastUpdate = $timer.ElapsedMilliseconds
                    $status = if ($TotalBytes -gt 0) { '{0:N1} MB of {1:N1} MB' -f ($copied / 1MB), ($TotalBytes / 1MB) } else { '{0:N1} MB' -f ($copied / 1MB) }
                    if ($TotalBytes -gt 0) {
                        Write-Progress -Activity $Activity -Status $status -PercentComplete ([math]::Min(100, [int](($copied * 100) / $TotalBytes)))
                    }
                    else {
                        Write-Progress -Activity $Activity -Status $status
                    }
                }
            }
            $succeeded = $true
        }
        finally {
            $fileStream.Dispose()
            Write-Progress -Activity $Activity -Completed
            if (-not $succeeded) { Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue }
        }
    }

    # Downloads a URL (following redirects) to disk with progress and returns the HTTP status code.
    function Save-UriWithProgress {
        param(
            [Parameter(Mandatory = $true)][string]$Uri,
            [Parameter(Mandatory = $true)][string]$Path,
            [Parameter(Mandatory = $true)][string]$UserAgent,
            [long]$TotalBytes,
            [string]$Activity = 'Downloading'
        )
        $request = [System.Net.HttpWebRequest]::Create($Uri)
        $request.UserAgent = $UserAgent
        $webResponse = $request.GetResponse()
        try {
            if ($TotalBytes -le 0) { $TotalBytes = $webResponse.ContentLength }
            $stream = $webResponse.GetResponseStream()
            try { Save-StreamWithProgress -Source $stream -Path $Path -TotalBytes $TotalBytes -Activity $Activity }
            finally { $stream.Dispose() }
            return [int]$webResponse.StatusCode
        }
        finally { $webResponse.Close() }
    }

    $functionName = $MyInvocation.MyCommand.Name
    $statusCode = 0
    $menu = @()
    $hasToken = $AccessToken -and $AccessToken.Length -gt 0
    if ($Private -and -not $hasToken) {
        Write-Log -logFile $LogFile -Module $functionName -Message "A private repository was specified but no access token was provided." -LogLevel "Error"
        throw "An access token is required for private repositories. Provide -AccessToken."
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    # GitHub rejects API requests that have no User-Agent header.
    $headers = @{
        'User-Agent' = 'sak-Get-GithubRelease'
        'Accept'     = 'application/vnd.github+json'
    }
    if ($hasToken) {
        $headers['Authorization'] = "Bearer $([System.Net.NetworkCredential]::new('', $AccessToken).Password)"
    }
    $encodedTag = [uri]::EscapeDataString($tag)
    $uri = if ($tag -eq 'latest') { "https://api.github.com/repos/$owner/$repo/releases/latest" } else { "https://api.github.com/repos/$owner/$repo/releases/tags/$encodedTag" }
    Write-Verbose "[$functionName] Starting for $owner/$repo from $uri (private: $Private, token supplied: $hasToken)"
    Write-Log -logFile $LogFile -Module $functionName -Message "Starting Get-GithubRelease for $owner/$repo from $uri (private: $Private, token supplied: $hasToken)" -LogLevel "INFORMATION"
    try {
        $release = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        # A pattern without wildcard characters is treated as "contains" (e.g. "msi" matches "PowerShell-7.7.0-win-x64.msi").
        $likePattern = if ($pattern -match '[*?\[]') { $pattern } else { "*$pattern*" }
        $assets = @($release.assets | Where-Object { $_.name -like $likePattern })
        Write-Log -logFile $LogFile -Module $functionName -Message "Found $($assets.Count) assets matching '$likePattern': $($assets.name -join ', ')" -LogLevel "INFORMATION"
        Write-Host "Found $($assets.Count) matching assets."
        if ($assets.Count -eq 0) {
            Write-Log -logFile $LogFile -Module $functionName -Message "No matching assets found." -LogLevel "Warning"
            Write-Host "No matching assets found."
            return $statusCode
        }
        elseif ($assets.Count -eq 1) {
            $asset = $assets[0]
            Write-Log -logFile $LogFile -Module $functionName -Message "Single matching asset found: $($asset.name)" -LogLevel "INFORMATION"
        }
        else {
            if ($NonInteractive) {
                Write-Host "More than one asset matches '$pattern'. Refine the pattern to match exactly one:" -ForegroundColor Yellow
                $assets | ForEach-Object { Write-Host "  - $($_.name)" }
                Write-Log -logFile $LogFile -Module $functionName -Message "Non-interactive mode and $($assets.Count) assets match '$pattern'. Nothing downloaded." -LogLevel "Warning"
                return -1
            }
            $assets | ForEach-Object { $menu += $_.name }
            $selection = Show-NumericMenu -choices $menu -banner "Select an asset to download:"
            if ($selection -eq 0) {
                return $statusCode
            }
            $asset = @($assets | Where-Object { $_.name -eq $selection })[0]
        }
        if (-not (Test-Path -Path $OutputDirectory -PathType Container)) {
            New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
        }
        $outputPath = Join-Path $OutputDirectory $asset.name
        Write-Log -logFile $LogFile -Module $functionName -Message "Downloading $($asset.name) to $outputPath" -LogLevel "INFORMATION"
        Write-Host "Downloading $($asset.name) to $outputPath"
        if ($Private) {
            # browser_download_url returns 404 for private assets, so use the API asset URL.
            # The API answers with a redirect to a pre-signed URL; follow it manually so the token is not forwarded to the storage host.
            # The API may also answer HTTP 200 with the asset body streamed directly, in which case save that body.
            $downloadUri = $null
            $apiStatus = 0
            $apiRequest = [System.Net.HttpWebRequest]::Create($asset.url)
            $apiRequest.AllowAutoRedirect = $false
            $apiRequest.UserAgent = $headers['User-Agent']
            $apiRequest.Accept = 'application/octet-stream'
            $apiRequest.Headers.Add('Authorization', $headers['Authorization'])
            $apiResponse = $apiRequest.GetResponse()
            try {
                $apiStatus = [int]$apiResponse.StatusCode
                if ($apiStatus -eq 200) {
                    $responseStream = $apiResponse.GetResponseStream()
                    try { Save-StreamWithProgress -Source $responseStream -Path $outputPath -TotalBytes ([long]$asset.size) -Activity "Downloading $($asset.name)" }
                    finally { $responseStream.Dispose() }
                }
                else {
                    $downloadUri = $apiResponse.Headers['Location']
                }
            }
            finally { $apiResponse.Close() }
            if ($apiStatus -eq 200) {
                $response = [pscustomobject]@{ StatusCode = 200 }
            }
            else {
                if ([string]::IsNullOrWhiteSpace($downloadUri)) { throw "GitHub did not return a download location for $($asset.name) (HTTP $apiStatus)." }
                $downloadStatus = Save-UriWithProgress -Uri $downloadUri -Path $outputPath -UserAgent $headers['User-Agent'] -TotalBytes ([long]$asset.size) -Activity "Downloading $($asset.name)"
                $response = [pscustomobject]@{ StatusCode = $downloadStatus }
            }
        }
        else {
            $downloadStatus = Save-UriWithProgress -Uri $asset.browser_download_url -Path $outputPath -UserAgent $headers['User-Agent'] -TotalBytes ([long]$asset.size) -Activity "Downloading $($asset.name)"
            $response = [pscustomobject]@{ StatusCode = $downloadStatus }
        }
        if ($response.StatusCode -ne 200) {
            $statusCode = $response.StatusCode
            throw "Failed to download $($asset.name). Status code: $($response.StatusCode)"
        }
        Write-Log -logFile $LogFile -Module $functionName -Message "Successfully downloaded $($asset.name)" -LogLevel "INFORMATION"
        Write-Host "Successfully downloaded: $($asset.name)"
        return $response.StatusCode
    }
    catch {
        $httpCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        Write-Host "An error occurred: $_"
        Write-Log -logFile $LogFile -Module $functionName -Message "An error occurred (HTTP $httpCode): $_" -LogLevel "Error"
        if ($Private -and $httpCode -in 401, 403, 404) {
            Write-Host "Access was denied or the repository/release was not found. Check the token's scope (repo / Contents: read) and the repository name." -ForegroundColor Yellow
            Write-Log -logFile $LogFile -Module $functionName -Message "Private repository access failed with HTTP $httpCode. Check token scope and repository name." -LogLevel "Warning"
        }
        if ($httpCode) { return $httpCode }
        if ($statusCode) { return $statusCode }
        return -2
    }
}
