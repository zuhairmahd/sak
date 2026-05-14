[CmdletBinding()]
param(
    [string]$inputString,
    [switch]$CheckRegKeyExists,
    [switch]$GetUninstallCommands,
    [switch]$KillGuiltyProcesses,
    [switch]$ManageServices,
    [string[]]$ServiceNames,
    [string]$ServiceOperation,
    [switch]$CreateZIPArchive,
    [string]$ArchiveDestination,
    [switch]$CheckProductInstallStatus,
    [string]$ProductStatusExportPath,
    [switch]$ExtractEmailAddresses,
    [string]$EmailExportPath,
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
            if ([string]::IsNullOrWhiteSpace($line))
            {
                break
            }
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
        $syntaxCheck = Test-PowerShellSyntax -File $function
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
    },
    @{
        name        = "ManageServices"
        description = "Start, stop, restart, or check status of Windows services."
    },
    @{
        name        = "CreateZIPArchive"
        description = "Create a ZIP archive from a list of files."
    },
    @{
        name        = "CheckProductInstallStatus"
        description = "Check installation status and details for software by keyword."
    },
    @{
        name        = "ExtractEmailAddresses"
        description = "Extract all unique email addresses from a text file."
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
        "ManageServices"
        {
            $ManageServices = $true
            [array]$ServiceNames = Get-UserInput -message "Enter service name(s) to manage:" -inputType "array"
            $operationChoices = @(
                @{ name = "Start"; description = "Start the service(s)." },
                @{ name = "Stop"; description = "Stop the service(s)." },
                @{ name = "Restart"; description = "Restart the service(s)." },
                @{ name = "Status"; description = "Check the current status of the service(s)." }
            )
            $ServiceOperation = Show-NumericMenu -choices $operationChoices -banner "Select the operation to perform:" -RequireEnter
        }
        "CreateZIPArchive"
        {
            $CreateZIPArchive = $true
            [array]$inputString = Get-UserInput -message "Enter file path(s) to include in the ZIP:" -inputType "array"
            $ArchiveDestination = Get-UserInput -message "Enter destination path for the ZIP file (e.g. C:\output\archive.zip):"
        }
        "CheckProductInstallStatus"
        {
            $CheckProductInstallStatus = $true
            [array]$inputString = Get-UserInput -message "Enter keyword(s) to search for installed products:" -inputType "array"
            $productStatusExportInput = Get-UserInput -message "Enter export path for results CSV (or leave blank to skip):"
            if (-not [string]::IsNullOrWhiteSpace($productStatusExportInput))
            {
                $ProductStatusExportPath = $productStatusExportInput
            }
        }
        "ExtractEmailAddresses"
        {
            $ExtractEmailAddresses = $true
            $inputString = Get-UserInput -message "Enter the path to the text file to scan for email addresses:"
            $emailExportInput = Get-UserInput -message "Enter export path to save email addresses (or leave blank to skip):"
            if (-not [string]::IsNullOrWhiteSpace($emailExportInput))
            {
                $EmailExportPath = $emailExportInput
            }
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
    Write-Host "Select whether you want to kill all processes or a single process:"
    $choice = Read-Host -Prompt "Enter 'all' to kill all guilty processes or 'single' to specify one process"
    while ($choice -notin @('all', 'single', 'a', 's'))
    {
        Write-Host "Invalid choice. Please enter 'all' or 'single'." -ForegroundColor Yellow
        #beep
        [console]::beep(1000, 300)
        $choice = Read-Host -Prompt "Enter 'all' to kill all guilty processes or 'single' to specify one process"
    }
    $choice = if ($choice -in @('a', 'all'))
    {
        'all'
    }
    else
    {
        'single'
    }
    switch ($choice)
    {
        all
        {
            $processesClosed = Close-Process -ProcessNames $guiltyProcessesToStop
        }
        single
        {
            $runningProcesses = Get-Process | Select-Object -ExpandProperty ProcessName
            $guiltyProcessesToStop = $guiltyProcessesToStop | Where-Object { $runningProcesses -contains $_ }
            $processChoice = Show-NumericMenu -choices $guiltyProcessesToStop -banner "Select a process to close:" -RequireEnter
            if ($null -eq $processChoice -or $processChoice -eq 0)
            {
                Write-Host "No process selected. Exiting process closure."
                write-log -logFile $LogFile -Module $scriptName -Message "No process selected for closure. Exiting." -LogLevel "Warning"
            }
            else
            {
                Write-Host "You selected to close process: $processChoice"
                $processesClosed = Close-Process -ProcessNames $processChoice
            }
        }
    }
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
        $exportData = @()
        if ($uninstallData.products.count -gt 1)
        {
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
        else
        {
            # Single product - build export object from the mostLikelyMatch already displayed above
            $exportObj = [PSCustomObject]@{
                ProductName             = $mostLikely.Name
                Version                 = $mostLikely.Version
                Publisher               = $mostLikely.Publisher
                SizeMB                  = $mostLikely.SizeMB
                InstallDate             = $mostLikely.InstallDate
                RegistryPath            = $mostLikely.RegistryPath
                RegistryKey             = $mostLikely.RegKey
                InstallLocation         = $mostLikely.InstallLocation
                IsMostLikely            = $true
                RawUninstallCmd         = $mostLikely.UninstallCmd
                UninstallFilePath       = $(if ($parsed)
                    {
                        $parsed.FilePath
                    }
                    else
                    {
                        $null
                    })
                UninstallArguments      = $(if ($parsed)
                    {
                        $parsed.Arguments
                    }
                    else
                    {
                        $null
                    })
                RawQuietUninstallCmd    = $mostLikely.QuietUninstall
                QuietUninstallFilePath  = $(if ($parsedQuiet)
                    {
                        $parsedQuiet.FilePath
                    }
                    else
                    {
                        $null
                    })
                QuietUninstallArguments = $(if ($parsedQuiet)
                    {
                        $parsedQuiet.Arguments
                    }
                    else
                    {
                        $null
                    })
            }
            $exportData += $exportObj
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

if ($ManageServices)
{
    if (-not $ServiceOperation)
    {
        Write-Host "No service operation specified. Exiting." -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "ManageServices: no operation specified." -LogLevel "Error"
        $exitCode = 1
    }
    elseif (-not $ServiceNames -or $ServiceNames.Count -eq 0)
    {
        Write-Host "No service names provided. Exiting." -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "ManageServices: no service names provided." -LogLevel "Error"
        $exitCode = 1
    }
    else
    {
        Write-Host "`n===============================================================" -ForegroundColor Cyan
        Write-Host "Managing service(s): $($ServiceNames -join ', ') - Operation: $ServiceOperation" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        write-log -logFile $LogFile -Module $scriptName -Message "ManageServices: Operation=$ServiceOperation, Services=$($ServiceNames -join ', ')" -LogLevel "Information"

        $serviceResults = Invoke-ServiceManager -serviceNames $ServiceNames -Operation $ServiceOperation
        foreach ($result in $serviceResults)
        {
            Write-Host "`n-------------------------------------------------------------------" -ForegroundColor DarkGray
            Write-Host "Service:     $($result.DisplayName) ($($result.Name))"
            if ($result.NotFound)
            {
                Write-Host "Status:      Not found" -ForegroundColor Red
                write-log -logFile $LogFile -Module $scriptName -Message "Service '$($result.Name)' not found." -LogLevel "Warning"
            }
            else
            {
                Write-Host "Operation:   $($result.Operation)"
                Write-Host "Status:      $($result.Status)"
                Write-Host "New Status:  $($result.NewStatus)"
                write-log -logFile $LogFile -Module $scriptName -Message "Service '$($result.Name)': Operation=$($result.Operation), Status=$($result.Status), NewStatus=$($result.NewStatus)" -LogLevel "Information"
            }
        }
        Write-Host "`n===============================================================" -ForegroundColor Cyan
        Write-Host "Service management complete." -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
    }
}

if ($CreateZIPArchive)
{
    if ([string]::IsNullOrWhiteSpace($ArchiveDestination))
    {
        Write-Host "No destination path specified for the ZIP archive. Exiting." -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "CreateZIPArchive: no destination path specified." -LogLevel "Error"
        $exitCode = 1
    }
    elseif (-not $inputString -or $inputString.Count -eq 0)
    {
        Write-Host "No file paths provided for the ZIP archive. Exiting." -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "CreateZIPArchive: no file paths provided." -LogLevel "Error"
        $exitCode = 1
    }
    else
    {
        Write-Host "`n===============================================================" -ForegroundColor Cyan
        Write-Host "Creating ZIP archive at: $ArchiveDestination" -ForegroundColor Cyan
        Write-Host "Files to include: $($inputString -join ', ')" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        write-log -logFile $LogFile -Module $scriptName -Message "CreateZIPArchive: Destination=$ArchiveDestination, Files=$($inputString -join ', ')" -LogLevel "Information"

        $zipResult = New-ZipPackage -FilePaths $inputString -DestinationPath $ArchiveDestination
        if ($zipResult)
        {
            Write-Host "`nZIP archive created successfully: $ArchiveDestination" -ForegroundColor Green
            write-log -logFile $LogFile -Module $scriptName -Message "ZIP archive created successfully: $ArchiveDestination" -LogLevel "Information"
        }
        else
        {
            Write-Host "`nFailed to create ZIP archive. Check log for details." -ForegroundColor Red
            write-log -logFile $LogFile -Module $scriptName -Message "Failed to create ZIP archive: $ArchiveDestination" -LogLevel "Error"
            $exitCode = 1
        }
    }
}

if ($CheckProductInstallStatus)
{
    if (-not $inputString -or $inputString.Count -eq 0)
    {
        Write-Host "No keywords provided to search for installed products. Exiting." -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: no keywords provided." -LogLevel "Error"
        $exitCode = 1
    }
    else
    {
        Write-Host "`n===============================================================" -ForegroundColor Cyan
        Write-Host "Checking installation status for keyword(s): $($inputString -join ', ')" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: Keywords=$($inputString -join ', ')" -LogLevel "Information"

        $statusData = Get-ProductInstallStatus -Products $inputString
        if ($statusData.HasErrors)
        {
            Write-Host "Error checking product status: $($statusData.Message)" -ForegroundColor Red
            write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus error: $($statusData.Message)" -LogLevel "Error"
            $exitCode = 1
        }
        elseif ($statusData.Products.Count -eq 0)
        {
            Write-Host "No installed products found matching keyword(s): $($inputString -join ', ')" -ForegroundColor Yellow
            write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: no products found." -LogLevel "Information"
        }
        else
        {
            Write-Host "Found $($statusData.Products.Count) installed product(s):" -ForegroundColor Green
            write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: found $($statusData.Products.Count) product(s)." -LogLevel "Information"

            $productExportData = @()
            foreach ($product in $statusData.Products)
            {
                Write-Host "`n-------------------------------------------------------------------" -ForegroundColor DarkGray
                Write-Host "Product Name: $($product.Name)" -ForegroundColor White
                Write-Host "Version:      $($product.Version)"
                Write-Host "Publisher:    $($product.Publisher)"
                Write-Host "Install Date: $($product.InstallDate)"
                Write-Host "Size:         $($product.SizeMB) MB"
                write-log -logFile $LogFile -Module $scriptName -Message "Product: $($product.Name), Version: $($product.Version), Publisher: $($product.Publisher)" -LogLevel "Information"

                $productExportData += [PSCustomObject]@{
                    ProductName = $product.Name
                    Version     = $product.Version
                    Publisher   = $product.Publisher
                    InstallDate = $product.InstallDate
                    SizeMB      = $product.SizeMB
                }
            }
            Write-Host "`n===============================================================" -ForegroundColor Cyan

            if (-not [string]::IsNullOrWhiteSpace($ProductStatusExportPath))
            {
                try
                {
                    $productExportData | Export-Csv -Path $ProductStatusExportPath -NoTypeInformation -Encoding UTF8
                    Write-Host "Results exported to: $ProductStatusExportPath" -ForegroundColor Green
                    write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: exported to $ProductStatusExportPath" -LogLevel "Information"
                }
                catch
                {
                    Write-Host "Failed to export results: $_" -ForegroundColor Red
                    write-log -logFile $LogFile -Module $scriptName -Message "CheckProductInstallStatus: export failed: $_" -LogLevel "Error"
                    $exitCode = 1
                }
            }
        }
    }
}

if ($ExtractEmailAddresses)
{
    if ([string]::IsNullOrWhiteSpace($inputString) -or -not (Test-Path $inputString))
    {
        Write-Host "File not found or no path provided: '$inputString'" -ForegroundColor Red
        write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: file not found: '$inputString'" -LogLevel "Error"
        $exitCode = 1
    }
    else
    {
        Write-Host "`n===============================================================" -ForegroundColor Cyan
        Write-Host "Extracting email addresses from: $inputString" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: scanning file '$inputString'" -LogLevel "Information"

        $fileLines = Get-Content -Path $inputString
        $emails = Get-EmailAddresses -lines $fileLines

        if ($emails -eq "No email addresses found." -or $emails.Count -eq 0)
        {
            Write-Host "No email addresses found in: $inputString" -ForegroundColor Yellow
            write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: no addresses found." -LogLevel "Information"
        }
        else
        {
            Write-Host "Found $($emails.Count) unique email address(es):" -ForegroundColor Green
            foreach ($email in $emails)
            {
                Write-Host "  $email"
            }
            write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: found $($emails.Count) address(es)." -LogLevel "Information"

            if (-not [string]::IsNullOrWhiteSpace($EmailExportPath))
            {
                try
                {
                    $emails | Out-File -FilePath $EmailExportPath -Encoding UTF8
                    Write-Host "`nEmail addresses saved to: $EmailExportPath" -ForegroundColor Green
                    write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: saved to '$EmailExportPath'" -LogLevel "Information"
                }
                catch
                {
                    Write-Host "Failed to save email addresses: $_" -ForegroundColor Red
                    write-log -logFile $LogFile -Module $scriptName -Message "ExtractEmailAddresses: export failed: $_" -LogLevel "Error"
                    $exitCode = 1
                }
            }
        }
    }
}

write-log -logFile $logFile -FinishLogging
exit $exitCode

