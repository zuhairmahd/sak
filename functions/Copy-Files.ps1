function Copy-Files()
{
    <#
    .SYNOPSIS
        Copy an array of files to a destination folder, optionally recreating source subfolder structure.

    .DESCRIPTION
        Copy-Files accepts a list of absolute file paths and copies them to the Destination folder.
        By default it preserves each file's relative subfolder hierarchy beneath Destination, creating
        missing directories as required. Use -noSubFolders to copy all files directly into Destination.
        Wildcard entries containing *.* are expanded (their parent folder is enumerated and all files are added).
        Returns a rich PSCustomObject describing the outcome (success/failure counts, lists, messages).

    .PARAMETER FilesToCopy
        One or more full file paths. Entries containing *.* are expanded to all files in the parent folder.

    .PARAMETER Destination
        Target root folder. Missing folders (or recreated subfolders) are created as needed.

    .PARAMETER noSubFolders
        Switch. If specified, subfolder structure is ignored and all files are copied directly into Destination.

    .EXAMPLE
        Copy-Files -FilesToCopy @("C:\Source\file1.txt","C:\Source\Sub\file2.log") -Destination "D:\Backup"
        Copies file1.txt to D:\Backup\ and file2.log to D:\Backup\Sub\ (creating Sub if necessary).

    .EXAMPLE
        Copy-Files -FilesToCopy "C:\Source\Reports\*.*" -Destination "D:\Archive"
        Expands the wildcard and copies every file in C:\Source\Reports into matching subfolders under D:\Archive.

    .EXAMPLE
        Copy-Files -FilesToCopy @("C:\Source\a.txt","C:\Source\B\b.txt") -Destination "D:\Flat" -noSubFolders
        Copies both files directly into D:\Flat without recreating the B subfolder.

    .EXAMPLE
        $result = Copy-Files -FilesToCopy $fileList -Destination "C:\Target"
        if ($result.allFilesCopied) { "All $($result.filesCopiedCount) files copied." }
        Uses the returned object to confirm success.

    .OUTPUTS
        PSCustomObject:
            allFilesCopied (bool)
            filesToCopyCount (int)
            filesCopied (string[])
            filesCopiedCount (int)
            filesNotCopied (string[])
            filesNotCopiedCount (int)
            subfoldersCreated (string[])
            subfoldersCreatedCount (int)
            message (string[])

    .NOTES
        Requires helper write-log and a pre-set $logFile variable. Designed for Intune / automation scenarios.
        Error details are logged via write-log with -logLevel Error.

    .LINK
        (Internal) write-log
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FilesToCopy,
        [Parameter(Mandatory = $true)]
        [string]$Destination,
        [switch]$noSubFolders,
        [switch]$UseWildCards
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] $($FilesToCopy.count) File paths provided: $($FilesToCopy -join ', ')"     
    write-log -logFile $logFile -module $functionName -message "$($FilesToCopy.count) File paths provided: $($FilesToCopy -join ', ')"                              
    foreach ($file in $FilesToCopy)
    {
        #if the string contains "*.*", get all files in the folder and add them to the list
        Write-Verbose "[$functionName] Checking for wildcards in file path: $file"
        if ($file -like "*`*.*" -and $UseWildCards)
        {
            $folderPath = Split-Path -Parent $file
            Write-Verbose "[$functionName] Expanding wildcard for folder: $folderPath"
            write-log -logFile $logFile -module $functionName -message "Expanding wildcard for folder: $folderPath"
            $expandedFiles = Get-ChildItem -Path $folderPath -File | ForEach-Object { $_.FullName }
            $FilesToCopy += $expandedFiles
            #remove the original wildcard entry
            $FilesToCopy = $FilesToCopy | Where-Object { $_ -ne $file }
            Write-Verbose "[$functionName] Expanded files: $($expandedFiles -join ', ')"        
            write-log -logFile $logFile -module $functionName -message "Expanded files: $($expandedFiles -join ', ')"                           
        }                                                                               
        elseif (-not (Test-Path -Path $file))
        {
            Write-Verbose "[$functionName] File does not exist and will be skipped: $file"                      
            write-log -logFile $logFile -module $functionName -message "File does not exist and will be skipped: $file" -logLevel "Warning"                           
            #remove the non-existing file from the list
            $FilesToCopy = $FilesToCopy | Where-Object { $_ -ne $file }
        }                           
        else
        {
            Write-Verbose "[$functionName] File exists and will be processed: $file"                      
            write-log -logFile $logFile -module $functionName -message "File exists and will be processed: $file"                           
        }                       
    }       
    
    $returnObject = @{
        allFilesCopied         = $false
        filesToCopyCount       = $FilesToCopy.Count       
        filesCopied            = @()
        filesCopiedCount       = 0
        filesNotCopied         = @()
        filesNotCopiedCount    = 0
        subfoldersCreated      = @()         
        subfoldersCreatedCount = 0
        message                = @()
    }
    Write-Host "Copying $($FilesToCopy.count) files to $Destination"
    write-log -logFile $logFile -module $functionName -message "Copying $($FilesToCopy.count) files to $Destination" -verbose                               
    
    foreach ($file in $FilesToCopy                                  )
    {
        if (-not $noSubFolders)
        {
            if (-not $sourceRoot)
            {
                $sourceRoot = Split-Path -Parent $returnObject.filesToCopy[0]
                foreach ($sourceFile in $returnObject.filesToCopy)
                {
                    $sourceParent = Split-Path -Parent $sourceFile
                    while (
                        $sourceRoot -and
                        $sourceParent -and
                        $sourceParent.Length -ge $sourceRoot.Length -and
                        -not $sourceParent.StartsWith($sourceRoot, [System.StringComparison]::OrdinalIgnoreCase)
                    )
                    {
                        $sourceRoot = Split-Path -Parent $sourceRoot
                    }
                }
            }

            #Check if any of the source file is in a subfolder, if so, append the relative subfolder to the destination and create it if needed.
            $subfolder = Split-Path -Parent $file
            Write-Verbose "[$functionName] Checking whether file $file is in a subfolder relative to source root $sourceRoot..."
            write-log -logFile $logFile -module $functionName -message "Checking whether file $file is in a subfolder relative to source root $sourceRoot..."
            if ($subfolder -ne $file)
            {
                $relativeSubfolder = [System.IO.Path]::GetRelativePath($sourceRoot, $subfolder)
                if ($relativeSubfolder -and $relativeSubfolder -ne '.')
                {
                    $destinationFolder = Join-Path -Path $Destination -ChildPath $relativeSubfolder
                    Write-Verbose "[$functionName] File $file is in subfolder $relativeSubfolder. Ensuring destination folder $destinationFolder exists..."
                    write-log -logFile $logFile -module $functionName -message "File $file is in subfolder $relativeSubfolder. Ensuring destination folder $destinationFolder exists..."
                    if (-not (Test-Path -Path $destinationFolder))
                    {
                        Write-Host "Creating folder: $destinationFolder"
                        write-log -logFile $logFile -module $functionName -message "Creating folder: $destinationFolder"
                        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                        $returnObject.subfoldersCreated += $destinationFolder
                        $returnObject.subfoldersCreatedCount++
                    }
                    $currentDestination = $destinationFolder
                    write-log -logFile $logFile -module $functionName -message "Set current destination to $currentDestination"
                }
                else
                {
                    Write-Host "No subfolder found for file: $file"
                    $currentDestination = $Destination
                    write-log -logFile $logFile -module $functionName -message "No subfolder found for file: $file. Set current destination to $currentDestination"
                }
            }
            else
            {
                Write-Host "No subfolder found for file: $file"
                $currentDestination = $Destination
                write-log -logFile $logFile -module $functionName -message "No subfolder found for file: $file. Set current destination to $currentDestination"
            }
        }
        else
        {
            $currentDestination = $Destination
            write-log -logFile $logFile -module $functionName -message "No subfolder processing. Set current destination to $currentDestination"
        }
        Write-Verbose "[$functionName] Processing file: $file"
        try
        {
            Copy-Item -Path $file -Destination $currentDestination -Force
            Write-Verbose "[$functionName] Copied $file to $currentDestination"
            write-log -logFile $logFile -module $functionName -message "Copied $file to $currentDestination"
            $returnObject.filesCopied += $file
            $returnObject.filesCopiedCount++
        }
        catch
        {
            Write-Host "Failed to copy $file to $currentDestination"
            Write-Error $_
            write-log -logFile $logFile -module $functionName -message "Failed to copy $file to $currentDestination. Error: $_" -logLevel "Error"
            $returnObject.filesNotCopied += $file
            $returnObject.filesNotCopiedCount++
        }
    }
    if ($returnObject.filesNotCopiedCount -eq 0 -and $returnObject.filesToCopyCount -eq $returnObject.filesCopiedCount) 
    {
        Write-Verbose "[$functionName] All files copied successfully."                      
        write-log -logFile $logFile -module $functionName -message "All files copied successfully."                         
        $returnObject.allFilesCopied = $true
        $returnObject.message += "All files copied successfully."
    }
    else
    {
        Write-Verbose "[$functionName] Some files failed to copy."                      
        write-log -logFile $logFile -module $functionName -message "Some files  failed to copy." -logLevel "Warning"
        $returnObject.allFilesCopied = $false
        $returnObject.message += "$($returnObject.filesNotCopiedCount) out of $($returnObject.filesToCopyCount) files failed to copy."
    }                                                                                   
    return $returnObject
}
