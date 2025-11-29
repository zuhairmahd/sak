# SAK - Swiss Army Knife for System Administrators

SAK is a PowerShell toolkit designed for Windows system administrators. It provides a collection of utility functions to streamline common administrative tasks such as process management, registry operations, software discovery, service management, and more.

## Features

SAK includes the following main interactive menu options:

### CheckRegKeyExists
Check if registry keys from a file exist with their correct values. This tool imports registry entries from a `.reg` file and validates that they exist on the system with the expected data.

### GetUninstallCommands
Discover uninstall commands for installed software based on keywords. Searches the Windows registry for installed applications matching your search terms and returns both standard and quiet uninstall commands, along with product details like version, publisher, and install location.

### KillGuiltyProcesses
Close common interfering processes before performing operations. This is useful when you need to stop applications that may be locking files or resources during administrative tasks (e.g., browsers, Office applications, Adobe products).

## Available Functions

SAK provides a rich library of PowerShell functions in the `functions` folder:

| Function | Description |
|----------|-------------|
| `Close-Process` | Gracefully closes processes by name, with fallback to forceful termination |
| `Find-FolderPath` | Locates a folder by name within a specified path |
| `Get-EmailAddresses` | Extracts and deduplicates email addresses from text content |
| `Get-FileFromURL` | Downloads files from a URL to a specified destination |
| `Get-FileHandle` | Uses Sysinternals handle.exe to find processes locking a file |
| `Get-UninstallCommand` | Retrieves uninstall commands for installed applications |
| `Get-WhoisInfo` | Performs WHOIS lookups for domains or IP addresses |
| `Import-RegKeysFromFile` | Parses .reg files into PowerShell-compatible registry entries |
| `Invoke-ExternalProcess` | Launches external processes with monitoring, timeout, and retry support |
| `Invoke-ServiceManager` | Manages Windows services (start, stop, restart, status) |
| `New-ZIPPackage` | Creates ZIP archives from multiple files |
| `Remove-FilesAndFolders` | Removes or moves files and folders with wildcard support |
| `Remove-RegKeys` | Removes registry keys and values |
| `Send-EmailWithAttachments` | Sends emails via Microsoft Graph API or Outlook COM automation |
| `Set-RegKeys` | Creates, updates, or validates registry keys and values |
| `Show-NumericMenu` | Displays an interactive numeric menu for user selection |
| `Test-IsProductInstalled` | Checks if a product is installed via file, registry, or MSI |
| `Test-IsServiceRunning` | Checks if a Windows service is currently running |
| `Test-PowerShellSyntax` | Validates PowerShell script syntax before execution |
| `Test-RegKeyExists` | Tests whether a registry key exists |
| `Test-isAdmin` | Checks if the current session has administrator privileges |
| `Write-Log` | Provides structured logging to file with timestamps |

## Usage

### Interactive Mode

Run the script without parameters to display an interactive menu:

```powershell
.\sak.ps1
```

### Command-Line Mode

You can also run specific operations directly via command-line parameters:

```powershell
# Check registry keys from a file
.\sak.ps1 -CheckRegKeyExists -inputString "C:\path\to\keys.reg"

# Get uninstall commands for software matching keywords
.\sak.ps1 -GetUninstallCommands -inputString "Chrome","Firefox" -ExportPath "C:\output\uninstall.csv"

# Close common interfering processes
.\sak.ps1 -KillGuiltyProcesses
```

## Requirements

- Windows PowerShell 5.1 or later
- Administrator privileges may be required for some operations
- For email functionality via Graph API: Microsoft.Graph PowerShell module
- For file handle detection: Sysinternals handle.exe

## License

This project is distributed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please see the [CONTRIBUTING](CONTRIBUTING.md) guide for details on how to contribute to this project.
