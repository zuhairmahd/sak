function Get-FileDetectionRule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallLocation
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -logFile $logFile -module $functionName -Message "Starting Get-FileDetectionRule for $InstallLocation" -LogLevel "Information"
    $criteria = [hashtable]@{
        FileExists   = "File exists"
        FileNotExist = "File does not exist"
        String       = "string comparison"
        Version      = "version is greater than or equal to"
    }
    $returnObject = @{
        success  = $false
        path     = $InstallLocation
        fileName = $null
        version  = $null
        criteria = $null
    }
    $exeFiles = Get-ChildItem -Path $installLocation -Filter *.exe -File -ErrorAction SilentlyContinue
    Write-Log -logFile $logFile -module $functionName -Message "Found $($exeFiles.Count) executable file(s) in $installLocation" -LogLevel "Information"
    if ($exeFiles.Count -eq 0) {
        Write-Log -logFile $logFile -module $functionName -Message "No executable files found in $installLocation" -LogLevel "Warning"
        return $returnObject
    }
    elseif ($exeFiles.Count -eq 1) {
        $exeFile = $exeFiles[0]
        Write-Log -logFile $logFile -module $functionName -Message "Only one executable found: $($exeFile.FullName)" -LogLevel "Information"
    }
    elseif ($exeFiles.Count -gt 1) {
        foreach ($file in $exeFiles) {
            Write-Log -logFile $logFile -module $functionName -Message "Executable found: $($file.FullName)" -LogLevel "Information"
            $exeFiles = $exeFiles | Where-Object { $_.BaseName -notmatch "uninst*" }
            Write-Log -logFile $logFile -module $functionName -Message "Filtered executable files: $($exeFiles.Count)" -LogLevel "Information"
            if ((Split-Path -Leaf $installLocation) -match $file.BaseName) {
                $exeFile = $file
                Write-Log -logFile $logFile -module $functionName -Message "Using executable matching the parent folder name: $($exeFile.FullName)" -LogLevel "Information"
                break
            }
        }
    }
    if (-not $exeFile) {
        Write-Log -logFile $logFile -module $functionName -Message "No suitable executable found in $installLocation" -LogLevel "Warning"
        return $returnObject
    }
    $returnObject.success = $true
    Write-Log -logFile $logFile -module $functionName -Message "Executable selected: $($exeFile.FullName)" -LogLevel "Information"
    $returnObject.fileName = $exeFile.Name
    $fileVersionInfo = Get-Item -Path $exeFile.FullName | Select-Object -ExpandProperty VersionInfo
    if ($fileVersionInfo.ProductVersion -eq $fileVersionInfo.FileVersion) {
        Write-Log -logFile $logFile -module $functionName -Message "ProductVersion and FileVersion are equal: $($fileVersionInfo.ProductVersion)" -LogLevel "Information"
        $returnObject.version = $fileVersionInfo.ProductVersion
        $returnObject.criteria = $criteria['Version'] + " " + $returnObject.version
    }
    else {
        Write-Log -logFile $logFile -module $functionName -Message "ProductVersion and FileVersion are not equal. Using FileVersion: $($fileVersionInfo.FileVersion)" -LogLevel "Information"
        $returnObject.version = $fileVersionInfo.FileVersion
        $returnObject.criteria = $criteria['String']
    }
    return $returnObject
}