function Remove-FilesAndFolders()
{
    <#
.SYNOPSIS
    Removes or moves files and folders to the TEMP directory.

.DESCRIPTION
    The Remove-FilesAndFolders function removes files and folders or moves them to the TEMP directory. It supports wildcard patterns
    for batch file operations and provides detailed logging and return information about the operation results.
    The function automatically detects whether a path is a file or folder and handles it accordingly.

.PARAMETER PathsToRemove
    An array of file or folder paths to remove. Supports wildcard patterns (e.g., "C:\Folder\*.*") to target multiple files.
    For folders, use the path without wildcards (e.g., "C:\Folder").

.PARAMETER MoveToTemp
    When specified, files and folders are moved to the TEMP directory instead of being deleted.

.OUTPUTS
    Returns a hashtable with the following properties:
    - allItemsRemoved: Boolean indicating if all items were successfully removed/moved
    - itemsToRemoveCount: Total number of items targeted for removal
    - itemsRemoved: Array of successfully removed/moved file/folder paths
    - itemsRemovedCount: Count of successfully removed/moved items
    - itemsNotRemoved: Array of file/folder paths that failed to remove/move
    - itemsNotRemovedCount: Count of items that failed to remove/move
    - message: Array of result messages

.EXAMPLE
    Remove-FilesAndFolders -PathsToRemove @("C:\temp\file1.txt", "C:\temp\file2.log")
    Removes the specified files.

.EXAMPLE
    Remove-FilesAndFolders -PathsToRemove @("C:\temp\*.log") -MoveToTemp
    Moves all .log files from C:\temp to the TEMP directory.

.EXAMPLE
    Remove-FilesAndFolders -PathsToRemove @("C:\temp\OldFolder")
    Removes the specified folder and all its contents.

.EXAMPLE
    Remove-FilesAndFolders -PathsToRemove @("C:\temp\file.txt", "C:\logs\OldLogs") -MoveToTemp
    Moves both the file and folder to the TEMP directory.

.NOTES
    Requires the write-log function for logging operations.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$PathsToRemove,
        [switch]$MoveToTemp,
        [switch]$UseWildCards
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] $($PathsToRemove.Count) Paths provided: $($PathsToRemove -join ', ')"                
    write-log -logFile $logFile -module $functionName -message "$($PathsToRemove.Count) Paths provided: $($PathsToRemove -join ', ')"                           
    foreach ($path in $PathsToRemove)
    {
        #if the string contains "*.*", get all files in the folder and add them to the list
        Write-Verbose "[$functionName] Checking for wildcards in path: $path"
        write-log -logFile $logFile -module $functionName -message "Checking for wildcards in path: $path"                                     
        if ($path -like "*`*.*" -and $UseWildCards          )
        {
            $folderPath = Split-Path -Parent $path
            Write-Verbose "[$functionName] Expanding wildcard for folder: $folderPath"
            write-log -logFile $logFile -module $functionName -message "Expanding wildcard for folder: $folderPath"
            $expandedFiles = Get-ChildItem -Path $folderPath -File | ForEach-Object { $_.FullName }
            $PathsToRemove += $expandedFiles
            #remove the original wildcard entry
            $PathsToRemove = $PathsToRemove | Where-Object { $_ -ne $path }
            Write-Verbose "[$functionName] Expanded files: $($expandedFiles -join ', ')"        
            write-log -logFile $logFile -module $functionName -message "Expanded files: $($expandedFiles -join ', ')"                           
        }                                                                               
        elseif (-not (Test-Path -Path $path))
        {
            Write-Verbose "[$functionName] Path does not exist and will be skipped: $path"                      
            write-log -logFile $logFile -module $functionName -message "Path does not exist and will be skipped: $path" -logLevel "Warning"                           
            #remove the non-existing path from the list
            $PathsToRemove = $PathsToRemove | Where-Object { $_ -ne $path }
        }                                                   
        else
        {
            Write-Verbose "[$functionName] Path exists and will be processed: $path"                      
            write-log -logFile $logFile -module $functionName -message "Path exists and will be processed: $path"                           
        }                   
    }       
    Write-Verbose "[$functionName] Total items to remove after expansion: $($PathsToRemove.Count)"                              
    write-log -logFile $logFile -module $functionName -message "Total items to remove after expansion: $($PathsToRemove.Count)" -verbose                                    

    $returnObject = @{
        allItemsRemoved      = $false
        itemsToRemoveCount   = $PathsToRemove.Count       
        itemsRemoved         = @()
        itemsRemovedCount    = 0
        itemsNotRemoved      = @()
        itemsNotRemovedCount = 0
        message              = @()
    }
    if ($PathsToRemove.Count -eq 0)
    {
        Write-Verbose "[$functionName] No valid items to remove after expansion. Exiting function."                      
        write-log -logFile $logFile -module $functionName -message "No valid items to remove after expansion. Exiting function."                           
        $returnObject.allItemsRemoved = $true
        $returnObject.message += "No valid items to remove after expansion."
        return $returnObject
    }                               
    if ($MoveToTemp)
    {
        Write-Verbose "[$functionName] MoveToTemp switch is set. Items will be moved to TEMP instead of deleted."                      
        write-log -logFile $logFile -module $functionName -message "MoveToTemp switch is set. Items will be moved to TEMP instead of deleted."                           
    }                                           
    foreach ($item in $PathsToRemove)
    {
        # Determine if the path is a file or folder
        $isFolder = Test-Path -Path $item -PathType Container
        $itemType = if ($isFolder) { "folder" } else { "file" }
        
        Write-Verbose "[$functionName] Attempting to remove/move $($itemType): $item"                       
        write-log -logFile $logFile -module $functionName -message "Attempting to remove/move $($itemType): $item"                              
        try
        {
            if ($MoveToTemp)
            {
                Move-Item -Path $item -Destination $env:TEMP -Force -ErrorAction Stop | Out-Null        
                Write-Verbose "[$functionName] Moved $($itemType): $item to $env:TEMP"                           
                write-log -logFile $logFile -module $functionName -message "Moved $($itemType)  : $item to $env:TEMP"
            }
            else
            {
                if ($isFolder)
                {
                    Remove-Item -Path $item -Recurse -Force -ErrorAction Stop | Out-Null
                }
                else
                {
                    Remove-Item -Path $item -Force -ErrorAction Stop | Out-Null
                }
                Write-Verbose "[$functionName] Removed $($itemType): $item"                                                                 
                write-log -logFile $logFile -module $functionName -message "Removed $($itemType)                    : $item"                        
            }                           
            $returnObject.itemsRemoved += $item                 
            $returnObject.itemsRemovedCount++           
            write-log -logFile $logFile -module $functionName -message "Successfully removed/moved $($itemType): $item"             
            Write-Verbose "[$functionName] Successfully removed/moved $($itemType)      : $item"          
        }
        catch
        {
            Write-Verbose "[$functionName] Failed to remove/move $($itemType): $item. Error: $($_.Exception.Message)"                      
            write-log -logFile $logFile -module $functionName -message "Failed to remove/move $($itemType): $item. Error: $($_.Exception.Message)" -logLevel "Error"                           
            $returnObject.itemsNotRemoved += $item
            $returnObject.itemsNotRemovedCount++
            continue
        }                                                                                   
    }
    if ($returnObject.itemsNotRemovedCount -eq 0 -and $returnObject.itemsToRemoveCount -eq $returnObject.itemsRemoved.Count) 
    {
        Write-Verbose "[$functionName] All items removed successfully."                      
        write-log -logFile $logFile -module $functionName -message "All items removed successfully."                         
        $returnObject.allItemsRemoved = $true
        $returnObject.message += "All items removed successfully."
    }
    else
    {
        Write-Verbose "[$functionName] Some items failed to remove."                      
        write-log -logFile $logFile -module $functionName -message "Some items failed to remove." -logLevel "Warning"
        $returnObject.allItemsRemoved = $false
        $returnObject.message += "$($returnObject.itemsNotRemovedCount) out of $($returnObject.itemsToRemoveCount) items failed to remove."
    }                                                                                   
    return $returnObject
}
