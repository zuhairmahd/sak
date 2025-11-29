function Get-FileFromURL()
{
    [CmdletBinding()]
    param (
        [string]$url,
        [string]$destination
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Downloading file from URL: $url to destination: $destination"
    write-log -LogFile $logFile -module $functionName -message "Downloading file from URL: $url to destination: $destination"
    $destinationFolder = Split-Path -Path $destination -Parent
    Write-Verbose "[$functionName] Checking if the destination folder $destinationFolder exists"
    write-log -LogFile $logFile -module $functionName -message "Checking if the destination folder $destinationFolder exists"
    if (-not (Test-Path -Path $destinationFolder))
    {
        Write-Verbose "[$functionName] Creating destination folder: $destinationFolder"
        write-log -LogFile $logFile -module $functionName -message "Creating destination folder: $destinationFolder"
        try 
        {
            New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null
            Write-Verbose "[$functionName] Destination folder created successfully: $destinationFolder"
            write-log -LogFile $logFile -module $functionName -message "Destination folder created successfully: $destinationFolder"
        }
        catch
        {
            Write-Error "[$functionName] Failed to create destination folder: $destinationFolder. Error: $_"
            write-log -LogFile $logFile -module $functionName -message "Failed to create destination folder: $destinationFolder. Error: $_"
            return $false
        }
    }   
    try
    {
        Invoke-WebRequest -Uri $url -OutFile $destination -ErrorAction Stop
        Write-Host "File downloaded successfully from $url to $destination"
        write-log -LogFile $logFile -module $functionName -message "File downloaded successfully from $url to $destination"
        return $true
    }
    catch
    {
        Write-Host "Failed to download file from $url to $destination. Error: $_"
        write-log -LogFile $logFile -module $functionName -message "Failed to download file from $url to $destination. Error: $_"
        return $false
    }
}
