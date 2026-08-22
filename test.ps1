[CmdletBinding(DefaultParameterSetName = "Menu")]
param(
    [string]$keyword
)

#region helper functions
function ConvertFrom-UninstallCommand() {
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

    if ($cmd -match '^"([^"]+)"(.*)$') {
        # Quoted path (e.g., "C:\Program Files\App\uninstall.exe" /args)
        $filePath = $matches[1]
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed quoted path: FilePath='$filePath', Arguments='$arguments'"
    }
    elseif ($cmd -match '^(MsiExec\.exe)\s+(.*)$') {
        # Special case for MsiExec.exe (case-insensitive)
        $filePath = "MsiExec.exe"
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed MsiExec.exe command: FilePath='$filePath', Arguments='$arguments'"
    }
    elseif ($cmd -match '^([A-Z]:\\.+\.exe)\s+(.*)$') {
        # Unquoted full path with .exe extension
        Write-Verbose "[$functionName] Parsing unquoted full path with .exe extension: $cmd"
        write-log -logFile $logFile -Module $functionName -Message "Parsing unquoted full path with .exe extension: $cmd"
        $exeIndex = $cmd.LastIndexOf('.exe')
        if ($exeIndex -ge 0) {
            $filePath = $cmd.Substring(0, $exeIndex + 4).Trim()
            $arguments = $cmd.Substring($exeIndex + 4).Trim()
        }
        else {
            $filePath = $matches[1]
            $arguments = $matches[2].Trim()
        }
        Write-Verbose "[$functionName] Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed unquoted full path with .exe extension: FilePath='$filePath', Arguments='$arguments'"
    }
    elseif ($cmd -match '^(\S+\.exe)\s*(.*)$') {
        # Simple executable name without path
        $filePath = $matches[1]
        $arguments = $matches[2].Trim()
        Write-Verbose "[$functionName] Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"
        write-log -logFile $logFile -Module $functionName -Message "Parsed simple executable name: FilePath='$filePath', Arguments='$arguments'"
    }
    else {
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

function Get-UserInput() {
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
    if ($inputType -eq "int") {
        while ($true) {
            $userInput = Read-Host -Prompt $message
            if ([int]::TryParse($userInput, [ref]$null)) {
                break
            }
            else {
                Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Yellow
                #beep
                [console]::beep(1000, 300)
            }
        }
        Write-Verbose "[$functionName] User input received: $userInput"
        return [int]$userInput
    }
    elseif ($inputType -eq "bool") {
        while ($true) {
            $userInput = Read-Host -Prompt "$message (y/n)"
            if ($userInput -match '^(y|yes)$') {
                Write-Verbose "[$functionName] User input received: True"
                return $true
            }
            elseif ($userInput -match '^(n|no)$') {
                Write-Verbose "[$functionName] User input received: False"
                return $false
            }
            else {
                Write-Host "Invalid input. Please enter 'y' for yes or 'n' for no." -ForegroundColor Yellow
                #beep
                [console]::beep(1000, 300)
            }
        }
    }
    elseif ($inputType -eq "array") {
        Write-Host "$message (Enter multiple values one per line, finish with an empty line):"
        $inputArray = [System.Collections.ArrayList]@()
        while ($true) {
            $line = Read-Host -Prompt "> "
            if ([string]::IsNullOrWhiteSpace($line)) {
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
if (Test-Path $functionsFolder) {
    Write-Verbose "[$scriptName] Importing functions from $functionsFolder"
    $functions = Get-ChildItem -Path "$functionsFolder\*.ps1" -File
    foreach ($function in $functions) {
        Write-Verbose " [$scriptName] Importing function $function"
        $syntaxCheck = Test-PowerShellSyntax -File $function
        if ($syntaxCheck.HasErrors) {
            Write-Host "Syntax errors found in $($function.FullName). Skipping import." -ForegroundColor Red
            write-log -logFile $logFile -Module $scriptName -Message "Syntax errors found in $($function.FullName). Skipping import." -LogLevel "Error"
            continue
        }
        . $function.FullName
    }
}
else {
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
        name        = "ExtractEmailAddresses"
        description = "Extract all unique email addresses from a text file."
    },
    @{
        name        = "DownloadFileFromURL"
        description = "Download a file from a URL to a local destination."
    },
    @{
        name        = "GetLocalComputerInfo"
        description = "Retrieve session type, client OS, name, and IP address of this computer."
    },
    @{
        name        = "WhoisLookup"
        description = "Perform a WHOIS lookup on a domain name or IP address."
    },
    @{
        name        = "CleanupNetworkProfiles"
        description = "Remove network profiles matching a keyword from the system."
    },
    @{
        name        = "GetMSIProperties"
        description = "Retrieve properties from MSI files."
    }
)
$AdminMessage = "You must be an administrator to perform this operation. Please run the script as an administrator."
#endregion define variables

try {
    $uninstallData = Get-UninstallCommand -keywords "Python"
    if ($uninstallData.hasErrors) {
        Write-Host "Error discovering products: $($uninstallData.message)" -ForegroundColor Yellow
        write-log -logFile $LogFile -Module $scriptName -Message "Error discovering products: $($uninstallData.message)" -LogLevel "Warning"
        $exitCode = 1
        return
    }
    elseif ($uninstallData.products.Count -eq 0) {
        Write-Host "No products found matching keywords: Python. Nothing to uninstall."
        write-log -logFile $LogFile -Module $scriptName -Message "No products found matching keywords: Python. Nothing to uninstall." -LogLevel "Information"
        $exitCode = 0
        return
    }
    Write-Host "`n===================================================================" -ForegroundColor Cyan
    Write-Host "Found $($uninstallData.products.Count) product(s) matching keyword(s): $inputString" -ForegroundColor Cyan
    Write-Host "===================================================================" -ForegroundColor Cyan
    write-log -logFile $LogFile -Module $scriptName -Message "Found $($uninstallData.products.Count) product(s) to uninstall." -LogLevel "Information"
    $global:allProducts = $uninstallData.products


    $product = $uninstallData.products[0]
    $availableParameters = @{}
    foreach ($key in $product.PSObject.Properties.Name) {
        $value = $product.$key
        if ($value -ne $null) {
            if (-not $availableParameters.ContainsKey($key)) {
                $availableParameters[$key] = @()
            }
            $availableParameters[$key] += $value
        }
    }

    #Now display all available parameters
    foreach ($key in $availableParameters.Keys) {
        $values = $availableParameters[$key] -join ", "
        Write-Host "${key}: $values"
    }
}
catch {
    write-log -logFile $LogFile -Module $scriptName -Message "Exception occurred: $($_.Exception.Message)" -LogLevel "Error"
}
finally {
    # Perform any necessary cleanup here
    Write-Host "Script execution completed with exit code $exitCode"
    exit $exitCode
}