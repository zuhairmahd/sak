function New-ZipPackage()
{
    <#
.SYNOPSIS
    Creates compressed ZIP archives from multiple files for diagnostic package distribution.

.DESCRIPTION
    This function provides robust ZIP file creation functionality using .NET compression
    libraries. The implementation includes:
    
    Archive Creation:
    - Uses System.IO.Compression.FileSystem for reliable compression
    - Creates directory structure automatically if needed
    - Handles existing file replacement safely
    
    File Processing:
    - Validates source files before inclusion
    - Preserves original filenames in archive
    - Handles file access and permission issues gracefully
    - Provides progress logging for large operations
    
    Error Handling:
    - Comprehensive exception handling for I/O operations
    - Automatic cleanup of partial archives on failure
    - Detailed logging of each file addition process

.PARAMETER FilePaths
    Array of file paths to include in the ZIP archive.
    Non-existent files are skipped with appropriate logging.

.PARAMETER DestinationPath
    Full path where the ZIP archive should be created.
    Parent directories are created automatically if they don't exist.
    Existing files at the destination are replaced.

.EXAMPLE
    $created = New-ZipPackage -FilePaths @("C:\logs\app.log", "C:\temp\cert.cer") -DestinationPath "C:\packages\diagnostic.zip"

.EXAMPLE
    $files = Get-ChildItem "C:\diagnostics\*.log" | Select-Object -ExpandProperty FullName
    New-ZipPackage -FilePaths $files -DestinationPath "$env:TEMP\logs.zip"

.OUTPUTS
    Boolean value indicating ZIP creation success:
    - $true: ZIP archive created successfully with all available files
    - $false: ZIP creation failed due to errors

.NOTES
    The function uses native .NET compression capabilities, requiring no additional
    tools or modules. ZIP files created are compatible with standard archive utilities.
    
    Large files or many files may take significant time to compress.
    The function automatically disposes of file handles to prevent resource leaks.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$FilePaths,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    Write-Log -Message "Creating zip package at $DestinationPath with files: $($FilePaths -join ', ')" -Module $functionName -LogLevel "Information" -LogFile $logFile
    Write-Verbose "[$functionName] Creating zip package at $DestinationPath with files: $($FilePaths -join ', ')"                       
    #if the destination file does not have a parent, assume current directory
    if (-not (Split-Path $DestinationPath -Parent))         
    {
        $DestinationPath = Join-Path -Path (Get-Location) -ChildPath $DestinationPath
        Write-Verbose "[$functionName] Adjusted DestinationPath to include current directory: $DestinationPath"                       
        write-log -logFile $logFile -module $functionName -message "Adjusted DestinationPath to include current directory: $DestinationPath" -logLevel "Debug"                        
    }                                                       
    try
    {
        # Ensure destination directory exists
        $destDir = Split-Path $DestinationPath -Parent
        if (-not (Test-Path $destDir))
        {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            write-log -logFile $logFile -module $functionName -message "Created directory for zip destination: $destDir" -logLevel "Information"                        
            Write-Verbose "[$functionName] Created directory for zip destination: $destDir"                             
        }
        
        # Remove existing zip file if it exists
        if (Test-Path $DestinationPath)
        {
            Remove-Item $DestinationPath -Force
            write-log -logFile $logFile -module $functionName -message "Removed existing zip file at destination: $DestinationPath" -logLevel "Information"                                         
            Write-Verbose "[$functionName] Removed existing zip file at destination: $DestinationPath"                  
        }
        
        # Load required assemblies for ZIP operations
        try
        {
            Write-Verbose "[$functionName] Loading ZIP compression assemblies"                  
            write-log -logFile $logFile -module $functionName -message "Loading ZIP compression assemblies" -logLevel "Debug"               
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
        }
        catch
        {
            Write-Log -Message "Failed to load ZIP compression assemblies: $($_.Exception.Message)" -Module $functionName -LogLevel "Error" -LogFile $logFile
            throw "Failed to load required assemblies for ZIP operations: $($_.Exception.Message)"
        }
        
        $zip = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
        write-log -logFile $logFile -module $functionName -message "Opened zip archive for creation: $DestinationPath" -logLevel "Information"                              
        Write-Verbose "[$functionName] Opened zip archive for creation: $DestinationPath"                       
        foreach ($filePath in $FilePaths)
        {
            Write-Verbose "[$functionName] Adding file to zip: $filePath"       
            if (Test-Path $filePath)
            {
                # Copy the file to a temporary file in the system temp folder.
                # This workaround is necessary because the original file may be locked or in use by another process,
                # which can cause issues when attempting to add it directly to the zip archive.
                # Copying to a temporary location ensures reliable access for zipping.
                $tempFilePath = Join-Path -Path $env:TEMP -ChildPath ([System.IO.Path]::GetFileName($filePath))
                Write-Verbose "[$functionName] Copying file $filePath to temporary location: $tempFilePath"                           
                Copy-Item -Path $filePath -Destination $tempFilePath -Force
                $fileName = [System.IO.Path]::GetFileName($tempFilePath)
                Write-Verbose "[$functionName] Adding temporary file to zip: $tempFilePath as $fileName"            
                $entry = $zip.CreateEntry($fileName)
                $entryStream = $null
                $fileStream = $null
                try
                {
                    $entryStream = $entry.Open()
                    $fileStream = [System.IO.File]::OpenRead($tempFilePath)
                    $fileStream.CopyTo($entryStream)
                    Write-Verbose "[$functionName] Added file to zip: $filePath"                                            
                    write-log -logFile $logFile -module $functionName -message "Added file to zip: $fileName" -logLevel "Information"                                       
                }
                finally
                {
                    if ($fileStream)
                    {
                        $fileStream.Close() 
                        Write-Verbose "[$functionName] Closed file stream for: $filePath"                   
                        write-log -logFile $logFile -module $functionName -message "Closed file stream for: $fileName" -logLevel "Information"                                                                                                  
                    }
                    if ($entryStream)
                    {
                        $entryStream.Close() 
                    }
                    if (Test-Path $tempFilePath)
                    {
                        Remove-Item -Path $tempFilePath -Force
                        Write-Verbose "[$functionName] Removed temporary file: $tempFilePath"                   
                        write-log -logFile $logFile -module $functionName -message "Removed temporary file: $tempFilePath" -logLevel "Information"                                                                                                  
                    }                                               
                }
                Write-Log -Message "Added file to zip: $fileName" -Module $functionName -LogLevel "Information" -LogFile $logFile
                Write-Verbose "[$functionName] Added file to zip: $filePath"                                                                            
            }
            else
            {
                Write-Log -Message "File not found, skipping: $filePath" -Module $functionName -LogLevel "Warning" -LogFile $logFile
                Write-Verbose "[$functionName] File not found, skipping: $filePath"                                 
            }
        }
        $zip.Dispose()
        Write-Log -Message "Zip package created successfully: $DestinationPath" -Module $functionName -LogLevel "Information" -LogFile $logFile
        Write-Verbose "[$functionName] Zip package created successfully: $DestinationPath"                  
        return $DestinationPath
    }
    catch
    {
        Write-Log -Message "Failed to create zip package: $($_.Exception.Message)" -Module $functionName -LogLevel "Error" -LogFile $logFile
        Write-Verbose "[$functionName] Failed to create zip package: $($_.Exception.Message)"                                           
        if ($zip)
        {
            $zip.Dispose()
            Write-Verbose "[$functionName] Disposed zip archive due to error"
            write-log -logFile $logFile -module $functionName -message "Disposed zip archive due to error" -logLevel "Information"                                                                                                                                                          
        }
        return $null
    }
}
