# Contributing to SAK

Thank you for your interest in contributing to SAK (Swiss Army Knife for System Administrators)! This document provides guidelines and instructions for contributing.

## How to Contribute

### Reporting Issues

If you find a bug or have a feature request:

1. Check if the issue already exists in the issue tracker
2. If not, create a new issue with a clear title and description
3. Include steps to reproduce the issue (for bugs)
4. Provide relevant system information (PowerShell version, Windows version)

### Submitting Changes

1. **Fork the repository** and create your branch from `main`
2. **Make your changes** following the coding standards below
3. **Test your changes** thoroughly on a Windows system
4. **Submit a pull request** with a clear description of the changes

## Coding Standards

### PowerShell Style Guidelines

- Use **PascalCase** for function names (e.g., `Get-UninstallCommand`)
- Use **camelCase** for variable names (e.g., `$functionName`)
- Include **comment-based help** for all public functions with:
  - `.SYNOPSIS` - Brief description
  - `.DESCRIPTION` - Detailed description
  - `.PARAMETER` - Document each parameter
  - `.EXAMPLE` - Provide usage examples
  - `.OUTPUTS` - Describe return values
- Use **`[CmdletBinding()]`** for advanced functions
- Include **verbose logging** using `Write-Verbose` and the `Write-Log` function
- Handle errors gracefully with `try/catch` blocks

### Function Template

```powershell
function Verb-Noun()
{
    <#
    .SYNOPSIS
        Brief description of what the function does.

    .DESCRIPTION
        Detailed description of the function's behavior.

    .PARAMETER ParameterName
        Description of the parameter.

    .EXAMPLE
        Verb-Noun -ParameterName "value"
        Description of what this example does.

    .OUTPUTS
        Description of return value.

    .NOTES
        Additional notes about the function.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting operation"
    Write-Log -LogFile $logFile -Module $functionName -Message "Starting operation"

    # Function implementation

    return $result
}
```

### File Organization

- Place new functions in the `functions` folder as individual `.ps1` files
- Name the file after the function (e.g., `Get-NewFunction.ps1`)
- Keep related functionality together in a single function

## Testing

Before submitting changes:

1. Test your code on Windows PowerShell 5.1
2. Verify that the script loads without syntax errors
3. Test both interactive and command-line modes if applicable
4. Ensure backward compatibility with existing functionality

## Pull Request Process

1. Update documentation if you're changing functionality
2. Ensure your code follows the coding standards
3. Write a clear PR description explaining:
   - What changes were made
   - Why the changes were necessary
   - How to test the changes
4. Be responsive to feedback and make requested changes

## Code of Conduct

- Be respectful and constructive in all interactions
- Focus on the technical aspects of contributions
- Help create a welcoming environment for all contributors

## Questions?

If you have questions about contributing, feel free to open an issue for discussion.
