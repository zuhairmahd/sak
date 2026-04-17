function Remove-RegPaths()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )

    $functionName = $MyInvocation.MyCommand.Name
    $returnObject = @{
        success           = $false
        totalPaths        = $Paths.count   
        removedPaths      = @()
        totalRemovedPaths = 0
        message           = ""
    }   
    Write-Verbose "[$functionName] Starting removal of $($Paths.count) registry paths."
    write-log -logFile $logFile -Module $functionName -Message "Starting removal of $($Paths.count) registry paths."
    if ($Paths.count -eq 0)
    {
        Write-Verbose "[$functionName] No registry paths provided for removal. Exiting."
        write-log -logFile $logFile -Module $functionName -Message "No registry paths provided for removal. Exiting."
        $returnObject.message = "No registry paths provided for removal."       
        return $returnObject
    }       
    foreach ($path in $Paths)
    {
        Write-Verbose "[$functionName] Attempting to remove registry path: $path"
        write-log -logFile $logFile -Module $functionName -Message "Attempting to remove registry path: $path"          
        try
        {
            if (Test-Path -Path $path)
            {
                Write-Verbose "[$functionName] Removing registry path: $path"
                write-log -logFile $logFile -Module $functionName -Message "Removing registry path: $path"       
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop                   
                Write-Verbose "[$functionName] Successfully removed registry path: $path"
                write-log -logFile $logFile -Module $functionName -Message "Successfully removed registry path: $path"  
            }
            else
            {
                Write-Verbose "[$functionName] Registry path does not exist: $path. Skipping."
                write-log -logFile $logFile -Module $functionName -Message "Registry path does not exist: $path. Skipping."         
                $returnObject.message += "Path $path does not exist. Skipping.`n"
            }                   
            $returnObject.removedPaths += $path
            $returnObject.totalRemovedPaths++
        }
        catch
        {
            Write-Verbose "[$functionName] Error removing registry path: $path. Error: $_"
            write-log -logFile $logFile -Module $functionName -Message "Error removing registry path: $path. Error: $_" -logLevel "ERROR"
            $returnObject.message += "Failed to remove path $path. Error: $_`n"                     
        }
    }                                               
    $returnObject.success = ($returnObject.totalRemovedPaths -eq $returnObject.totalPaths)
    Write-Verbose "[$functionName] Completed removal of registry paths. Total removed: $($returnObject.totalRemovedPaths) out of $($returnObject.totalPaths). Success: $($returnObject.success) "
    write-log -logFile $logFile -Module $functionName -Message "Completed removal of registry paths. Total removed: $($returnObject.totalRemovedPaths) out of $($returnObject.totalPaths). Success: $($returnObject.success) "                                          
    return $returnObject                                
}