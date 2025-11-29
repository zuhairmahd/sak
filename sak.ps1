[CmdletBinding()]
param(
    [string]$inputString,
    [switch]$CheckRegKeyExists,
    [switch]$GetUninstallCommands,
    [switch]$KillGuiltyProcesses,
    [string]$ExportPath
)

#region helper functions
function ConvertFrom-UninstallCommand()
{
    <#
    .SYNOPSIS
        Parses an uninstall command string into FilePath and Arguments components.
    .PARAMETER cmd
        The uninstall command string to parse.
    .OUTPUTS
        Returns a PSCustomObject with FilePath and Arguments properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$cmd
    )

    $functionName = $MyInvocation.MyCommand.Name                
    $filePath = $null
    $arguments = $null
    
    if ($cmd -match '^"([^"]+)"(.*)$')
    {
        # Quoted path (e.g., "C:\Program Files\App\uninstall.exe" /args)
        $filePath = $matches[1]
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
    }
    elseif ($cmd -match '^(MsiExec\.exe)\s+(.*)$')
    {
        # Special case for MsiExec.exe (case-insensitive)
        $filePath = "MsiExec.exe"
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"                                   
    }
    elseif ($cmd -match '^([A-Z]:\\.+\.exe)\s+(.*)$')
    {
        # Unquoted full path with .exe extension
        Write-Verbose "[$functionName] Parsing unquoted full path with .exe extension: $cmd"
        write-log -logFile $logFile -Module $functionName -Message "Parsing unquoted full path with .exe extension: $cmd"   
        $exeIndex = $cmd.LastIndexOf('.exe')
        if ($exeIndex -ge 0)
        {
            $filePath = $cmd.Substring(0, $exeIndex + 4).Trim()
            $arguments = $cmd.Substring($exeIndex + 4).Trim()
        }
        else
        {
            $filePath = $matches[1]
            $arguments = $matches[2].Trim()
        }
        Write-Verbose "[$functionName] Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"                                    
    }
    elseif ($cmd -match '^(\S+\.exe)\s*(.*)$')
    {
        # Simple executable name without path
        $filePath = $matches[1]
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"             
        write-log -logFile $logFile -Module $functionName -Message "Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"                                
    }
    else
    {
        # Fallback: treat the whole command as filepath
        $filePath = $cmd.Trim()
        $arguments = ""
        Write-Verbose "[$functionName] Fallback parsing: FilePath='$filePath', Arguments='$arguments'"                          
        write-log -logFile $logFile -Module $functionName -Message "Fallback parsing: FilePath='$filePath', Arguments='$arguments'"                     
    }
    
    return [PSCustomObject]@{
        FilePath  = $filePath
        Arguments = $arguments
    }
}

function Get-UserInput()
{
    <#
    .SYNOPSIS
        Prompts the user for input with a specified message.
    .PARAMETER message
        The message to display to the user.
    .OUTPUTS
        Returns the user's input as a string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$message,
        [ValidateSet("string", "int", "bool", "array")]
        [string]$inputType = "string"   
    )
    
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Prompting user with message: $message"
    if ($inputType -eq "int")
    {
        while ($true)
        {
            $userInput = Read-Host -Prompt $message
            if ([int]::TryParse($userInput, [ref]$null))
            {
                break
            }
            else
            {
                Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Yellow
                #beep
                [console]::beep(1000, 300)                                   
            }
        }
        Write-Verbose "[$functionName] User input received: $userInput"
        return [int]$userInput
    }                               
    elseif ($inputType -eq "bool")
    {
        while ($true)
        {
            $userInput = Read-Host -Prompt "$message (y/n)"
            if ($userInput -match '^(y|yes)$')
            {
                Write-Verbose "[$functionName] User input received: True"
                return $true
            }
            elseif ($userInput -match '^(n|no)$')
            {
                Write-Verbose "[$functionName] User input received: False"
                return $false
            }
            else
            {
                Write-Host "Invalid input. Please enter 'y' for yes or 'n' for no." -ForegroundColor Yellow
                #beep
                [console]::beep(1000, 300)                                   
            }
        }
    }               
    elseif ($inputType -eq "array")
    {
        Write-Host "$message (Enter multiple values one per line, finish with an empty line):"
        $inputArray = [System.Collections.ArrayList]@()
        while ($true)
        {
            $line = Read-Host -Prompt "> "
            if ([string]::IsNullOrWhiteSpace($line)) { break }
            [void]$inputArray.Add($line.Trim())
        }   
        Write-Verbose "[$functionName] User input received: $($inputArray -join ', ')"
        return $inputArray
    }                                           
    # Default to string input
    $userInput = Read-Host -Prompt $message
    Write-Verbose "[$functionName] User input received: $userInput"
    return $userInput
}                       
#endregion helper functions

#region import functions.
. $PSScriptRoot\functions\Find-FolderPath.ps1                               
. $PSScriptRoot\functions\Test-PowerShellSyntax.ps1
$functionsFolder = Find-FolderPath -Path "$psscriptRoot" -FolderName "functions"
if (Test-Path $functionsFolder)
{
    Write-Verbose "[$scriptName] Importing functions from $functionsFolder"
    $functions = Get-ChildItem -Path "$functionsFolder\*.ps1" -File
    foreach ($function in $functions)
    {
        Write-Verbose " [$scriptName] Importing function $function"
        $syntaxCheck = Test-PowerShellSyntax -File $function -errorsOnly
        if ($syntaxCheck.HasErrors)
        {
            Write-Host "Syntax errors found in $($function.FullName). Skipping import." -ForegroundColor Red
            write-log -logFile $logFile -Module $scriptName -Message "Syntax errors found in $($function.FullName). Skipping import." -LogLevel "Error"
            continue
        }                                           
        . $function.FullName
    }
}
else
{
    Write-Host 'Cannot find the functions folder. Exiting script.' -ForegroundColor Red
    exit 1
}
#endregion import functions.

#region define variables
$scriptName = $MyInvocation.MyCommand.Name
$logFile = Join-Path -Path $env:TEMP\sak -ChildPath "logs\$($scriptName)_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$exitCode = 0
$guiltyProcessesToStop = (
    "dm",
    "acrobat",
    "acrocef",
    "acrolicapp",
    "adobecollabsync",
    "adobe_licensing_wf_acro",
    "adobe_licensing_wf_helper_acro",
    "chrome",
    "excel",
    "explorer", 
    "firefox",
    "msaccess",
    "msedge",
    "msedgewebview2",
    "officeclicktorun",
    "onedrive",
    "onenote",
    "onenotem",
    "outlook",
    "powerpnt",
    "teams",
    "winword"
)
$menuItems = @(
    @{    
        name        = "CheckRegKeyExists"
        description = "Check if registry keys from a file exist with correct values."   
    },
    @{  
        name        = "GetUninstallCommands"
        description = "Discover uninstall commands for installed software based on keywords."                           
    },
    @{  
        name        = "KillGuiltyProcesses" 
        description = "Close most common interfering processes before performing operations."                           
    }
)
#endregion define variables

write-log -logFile $logFile -StartLogging
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "Welcome to SAK, the Swiss Army Knife for System Administrators!" -ForegroundColor Cyan      
Write-Host "===============================================================" -ForegroundColor Cyan
    
#if no commandline parameters are passed, display the menu.
if (-not $PSBoundParameters.Keys.Count)
{
    write-log -logFile $LogFile -Module $scriptName -Message "No parameters provided. Displaying menu for user selection." -LogLevel "Information"                                    
    $userChoice = Show-NumericMenu -choices $menuItems -banner "Select an action to perform:" -RequireEnter
    
    switch ($userChoice)
    {
        "CheckRegKeyExists"
        {
            $inputString = Get-UserInput -message "Enter the path to the registry keys file:"
            $CheckRegKeyExists = $true
        }
        "GetUninstallCommands"
        {
            [array]$inputString = Get-UserInput -message "Enter keywords to search for uninstall commands:" -inputType "array"
            $GetUninstallCommands = $true
            $exportPathInput = Get-UserInput -message "Enter export path for uninstall commands CSV (or leave blank to skip export):"
            if (-not [string]::IsNullOrWhiteSpace($exportPathInput))
            {
                $ExportPath = $exportPathInput
            }
        }
        "KillGuiltyProcesses"
        {
            $KillGuiltyProcesses = $true
        }
        default
        {
            Write-Host "No valid selection made. Exiting script."
            write-log -logFile $LogFile -Module $scriptName -Message "No valid selection made. Exiting script." -LogLevel "Warning"                                    
            exit 1
        }
    }
}                                                                                               

if ($KillGuiltyProcesses)
{
    $processesClosed = Close-Process -ProcessNames $guiltyProcessesToStop
    Write-Verbose "[$scriptName] Processes closed result: $($processesClosed | Out-String)"         
    write-log -logFile $LogFile -Module $scriptName -Message "Processes closed result: $($processesClosed | Out-String)" -LogLevel "Information"                                    
    if ($processesClosed.allProcessesClosed)
    {
        Write-Host "All interfering processes closed successfully."
        Write-Log -LogFile $LogFile -Module $scriptName -Message "All interfering processes closed successfully." -LogLevel "Information"
    }
    else
    {
        Write-Host "Some processes could not be closed. See log for details."
        Write-Host "Processes not closed:"
        foreach ($proc in $processesClosed.processesNotClosed)
        {
            Write-Host " - $proc"
            write-log -logFile $LogFile -Module $scriptName -Message " - $proc" -LogLevel "Information"                         
        }                                           
        $exitCode = 1
    }                                           
}       

if ($CheckRegKeyExists)
{
    $regKeys = Import-RegKeysFromFile -FilePath $inputString
    if ($null -ne $regKeys)
    {
        $allKeysExist = Set-RegKeys -RegistryEntries $regKeys -CheckOnly
        Write-Host "Processed a total of $($allKeysExist.totalEntries) registry entries."
        if ($allKeysExist -and $allKeysExist.hasCorrectValues)
        {
            Write-Host "All registry keys exist with correct values: $($allKeysExist.message)"
            Write-Log -LogFile $logFile -Module $scriptName -Message "All registry keys exist with correct values." -LogLevel "Information"
        }
        else
        {
            Write-Host "Some registry keys are missing or have incorrect values."
            Write-Log -LogFile $logFile -Module $scriptName -Message "Some registry keys are missing or have incorrect values." -LogLevel "Warning"
            if ($allKeysExist.missingKeysCount -gt 0)
            {
                Write-Host "Missing keys count: $($allKeysExist.missingKeysCount)"
                Write-Log -LogFile $logFile -Module $scriptName -Message "Missing keys count: $($allKeysExist.missingKeysCount)" -LogLevel "Warning"
                foreach ($missingKey in $allKeysExist.missingKeys)
                {
                    Write-Host "Key path: $($missingKey.path)"
                    Write-Host "Key name: $($missingKey.name)"
                    Write-Host "Key value: $($missingKey.value)"
                    Write-Host "Key hive: $($missingKey.hive)   "
                    Write-Log -LogFile $logFile -Module $scriptName -Message " - $missingKey" -LogLevel "Warning"
                }
            }
            if ($allKeysExist.entriesFailedCount -gt 0)
            {
                Write-Host "Entries with incorrect values count: $($allKeysExist.entriesFailedCount)"
                Write-Log -LogFile $logFile -Module $scriptName -Message "Entries with incorrect values count: $($allKeysExist.entriesFailedCount)" -LogLevel "Warning"
                foreach ($failedEntry in $allKeysExist.entriesFailed)
                {
                    Write-Host "Key path: $($failedEntry.path)"
                    Write-Host "Key name: $($failedEntry.name)"
                    Write-Host "Expected value: $($failedEntry.expectedValue)"
                    Write-Host "Actual value: $($failedEntry.actualValue)"
                    Write-Log -LogFile $logFile -Module $scriptName -Message " - $failedEntry" -LogLevel "Warning"
                }               
            }

        }
    }
    else
    {
        Write-Host "No registry keys to check."
        Write-Log -LogFile $logFile -Module $scriptName -Message "No registry keys to check." -LogLevel "Warning"
        $exitCode = 1
    }
}

if ($GetUninstallCommands)
{
    $uninstallData = Get-UninstallCommand -keywords $inputString
    if ($uninstallData.hasErrors)
    {
        Write-Host "Error discovering products: $($uninstallData.message)" -ForegroundColor Yellow
        write-log -logFile $LogFile -Module $scriptName -Message "Error discovering products: $($uninstallData.message)" -LogLevel "Warning"
        $exitCode = 1
    }
    elseif ($uninstallData.products.Count -eq 0)
    {
        Write-Host "No products found matching keywords: $inputString. Nothing to uninstall."
        write-log -logFile $LogFile -Module $scriptName -Message "No products found matching keywords. Nothing to uninstall." -LogLevel "Information"
        $exitCode = 0
    }
    else
    {
        Write-Host "`n===================================================================" -ForegroundColor Cyan
        Write-Host "Found $($uninstallData.products.Count) product(s) matching keyword(s): $inputString" -ForegroundColor Cyan
        Write-Host "===================================================================" -ForegroundColor Cyan
        write-log -logFile $LogFile -Module $scriptName -Message "Found $($uninstallData.products.Count) product(s) to uninstall." -LogLevel "Information"
        
        # Display most likely candidate first if exists
        if ($uninstallData.mostLikelyMatch)
        {
            if ($uninstallData.products.Count -gt 1)
            {
                Write-Host "`n*** MOST LIKELY MAIN APPLICATION ***" -ForegroundColor Green
            }
            $mostLikely = $uninstallData.mostLikelyMatch
            Write-Host "Product Name:    $($mostLikely.Name)" -ForegroundColor White
            Write-Host "Version:         $($mostLikely.Version)"
            Write-Host "Publisher:       $($mostLikely.Publisher)"
            Write-Host "Size:            $($mostLikely.SizeMB) MB ($($mostLikely.SizeKB) KB)"
            if (-not ([string]::IsNullOrEmpty($mostLikely.InstallDate)))
            {
                Write-Host "Install Date:    $($mostLikely.InstallDate)"
            }
            if (-not ([string]::IsNullOrEmpty($mostLikely.InstallLocation)))
            {           
                Write-Host "Install Location: $($mostLikely.InstallLocation)"
            }
            Write-Host "Registry path: $($mostLikely.RegistryPath)"
            Write-Host "Registry Key:    $($mostLikely.RegKey)"
            Write-Host "`n--- Uninstall Commands ---" -ForegroundColor Yellow
            # Parse and display standard uninstall command
            if (-not [string]::IsNullOrWhiteSpace($mostLikely.UninstallCmd))
            {
                Write-Host "  Raw Uninstall Command:" -ForegroundColor Cyan
                Write-Host "    $($mostLikely.UninstallCmd)"
                $parsed = ConvertFrom-UninstallCommand -cmd $mostLikely.UninstallCmd
                Write-Host "  Parsed Uninstall Command:" -ForegroundColor Cyan
                Write-Host "    FilePath:  $($parsed.FilePath)"
                Write-Host "    Arguments: $($parsed.Arguments)"
            }
            # Parse and display quiet uninstall command if exists
            if (-not [string]::IsNullOrWhiteSpace($mostLikely.QuietUninstall))
            {
                Write-Host "`n  Raw Quiet Uninstall Command:" -ForegroundColor Cyan
                Write-Host "    $($mostLikely.QuietUninstall)"
                $parsedQuiet = ConvertFrom-UninstallCommand -cmd $mostLikely.QuietUninstall
                Write-Host "  Parsed Quiet Uninstall Command:" -ForegroundColor Cyan
                Write-Host "    FilePath:  $($parsedQuiet.FilePath)"
                Write-Host "    Arguments: $($parsedQuiet.Arguments)"
            }
            Write-Host "`n===================================================================" -ForegroundColor Green
        }
        # Collection for export
        if ($uninstallData.products.count -gt 1)
        {
            $exportData = @()
            Write-Host "`n`n*** ALL MATCHING PRODUCTS ***" -ForegroundColor Cyan
            foreach ($product in $uninstallData.products)
            {
                write-log -logFile $LogFile -Module $scriptName -Message "Processing uninstall for: $($product.Name) (Version: $($product.Version), Size: $($product.SizeMB)MB, Publisher: $($product.Publisher))" -LogLevel "Information"
            
                Write-Host "`n-------------------------------------------------------------------" -ForegroundColor DarkGray
                if ($product.IsMostLikely)
                {
                    Write-Host "Product Name:    $($product.Name) [MOST LIKELY]" -ForegroundColor Green
                }
                else
                {
                    Write-Host "Product Name:    $($product.Name)"
                }
                Write-Host "Version:         $($product.Version)"
                Write-Host "Publisher:       $($product.Publisher)"
                Write-Host "Size:            $($product.SizeMB) MB"
                Write-Host "Install Date:    $($product.InstallDate)"
                Write-Host "Registry path: $($product.RegistryPath)"
                Write-Host "Registry Key:    $($product.RegKey)"
                Write-Host "Install Location: $($product.InstallLocation)"      
                # Create export object
                $exportObj = [PSCustomObject]@{
                    ProductName             = $product.Name
                    Version                 = $product.Version
                    Publisher               = $product.Publisher
                    SizeMB                  = $product.SizeMB
                    InstallDate             = $product.InstallDate
                    RegistryPath            = $product.RegistryPath             
                    RegistryKey             = $product.RegKey
                    InstallLocation         = $product.InstallLocation
                    IsMostLikely            = $product.IsMostLikely
                    RawUninstallCmd         = $product.UninstallCmd
                    UninstallFilePath       = $null
                    UninstallArguments      = $null
                    RawQuietUninstallCmd    = $product.QuietUninstall
                    QuietUninstallFilePath  = $null
                    QuietUninstallArguments = $null
                }
                # Parse and display standard uninstall command
                if (-not [string]::IsNullOrWhiteSpace($product.UninstallCmd))
                {
                    Write-Host "`n  Raw Uninstall Command:" -ForegroundColor Cyan
                    Write-Host "    $($product.UninstallCmd)"
                    write-log -logFile $LogFile -Module $scriptName -Message "Raw uninstall command: $($product.UninstallCmd)" -LogLevel "Verbose"
                
                    $parsed = ConvertFrom-UninstallCommand -cmd $product.UninstallCmd
                    Write-Host "  Parsed Uninstall Command:" -ForegroundColor Cyan
                    Write-Host "    FilePath:  $($parsed.FilePath)"
                    Write-Host "    Arguments: $($parsed.Arguments)"
                    write-log -logFile $LogFile -Module $scriptName -Message "Parsed - FilePath: $($parsed.FilePath), Arguments: $($parsed.Arguments)" -LogLevel "Information"
                
                    $exportObj.UninstallFilePath = $parsed.FilePath
                    $exportObj.UninstallArguments = $parsed.Arguments
                }
                else
                {
                    Write-Host "`n  No standard uninstall command available" -ForegroundColor Yellow
                }
                # Parse and display quiet uninstall command if exists
                if (-not [string]::IsNullOrWhiteSpace($product.QuietUninstall))
                {
                    Write-Host "`n  Raw Quiet Uninstall Command:" -ForegroundColor Cyan
                    Write-Host "    $($product.QuietUninstall)"
                    write-log -logFile $LogFile -Module $scriptName -Message "Raw quiet uninstall command: $($product.QuietUninstall)" -LogLevel "Verbose"
                
                    $parsedQuiet = ConvertFrom-UninstallCommand -cmd $product.QuietUninstall
                    Write-Host "  Parsed Quiet Uninstall Command:" -ForegroundColor Cyan
                    Write-Host "    FilePath:  $($parsedQuiet.FilePath)"
                    Write-Host "    Arguments: $($parsedQuiet.Arguments)"
                    write-log -logFile $LogFile -Module $scriptName -Message "Parsed Quiet - FilePath: $($parsedQuiet.FilePath), Arguments: $($parsedQuiet.Arguments)" -LogLevel "Information"
                
                    $exportObj.QuietUninstallFilePath = $parsedQuiet.FilePath
                    $exportObj.QuietUninstallArguments = $parsedQuiet.Arguments
                }
                else
                {
                    Write-Host "`n  No quiet uninstall command available" -ForegroundColor DarkGray
                }
                $exportData += $exportObj
            }
            Write-Host "`n===================================================================" -ForegroundColor Cyan
            Write-Host "End of product list" -ForegroundColor Cyan
            Write-Host "===================================================================" -ForegroundColor Cyan
        }
        # Export if requested
        if (-not [string]::IsNullOrWhiteSpace($ExportPath))
        {
            try
            {
                $exportData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nData exported successfully to: $ExportPath" -ForegroundColor Green
                write-log -logFile $LogFile -Module $scriptName -Message "Data exported to: $ExportPath" -LogLevel "Information"
            }
            catch
            {
                Write-Host "`nFailed to export data: $_" -ForegroundColor Red
                write-log -logFile $LogFile -Module $scriptName -Message "Failed to export data: $_" -LogLevel "Error"
                $exitCode = 1
            }
        }
    }
}

write-log -logFile $logFile -FinishLogging
exit $exitCode

