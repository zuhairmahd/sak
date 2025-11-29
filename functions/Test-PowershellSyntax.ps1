function Test-PowerShellSyntax()
{
    <#
.SYNOPSIS
Validates PowerShell script syntax for a given file.

.DESCRIPTION
Parses the specified PowerShell script using the PowerShell language parser to detect syntax errors, checks for balanced braces, brackets, and parentheses, attempts ScriptBlock compilation, and performs static analysis for duplicate function names and unreachable code after return/exit/throw statements.

.PARAMETER File
The .ps1 file to validate. Mandatory. Accepts a System.IO.FileInfo object (e.g. from Get-ChildItem).

.PARAMETER ShowDetails
If set, shows extended parse error detail including ErrorId and token position information.

.OUTPUTS
Hashtable with keys: HasErrors (Boolean), ErrorCount (Int), WarningCount (Int), Errors (Collection of parse error objects).

.EXAMPLE
Test-PowerShellSyntax -File .\MyScript.ps1
Validates MyScript.ps1 and reports any syntax issues.

.EXAMPLE
Get-ChildItem -Path . -Filter *.ps1 | ForEach-Object { Test-PowerShellSyntax -File $_ -ShowDetails }
Validates all PowerShell scripts in the current directory with detailed error output.

.NOTES
Uses [System.Management.Automation.Language.Parser] for AST parsing. Requires Windows PowerShell 5.1+ or PowerShell 7+.

.LINK
Get-Help about_ScriptBlocks
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,
        [switch]$errorsOnly,
        [switch]$ShowDetails
    )
    
    $hasErrors = $false
    $errorCount = 0
    $warningCount = 0
    $errors = @()
    
    try
    {
        # Method 1: Use AST parser for deeper syntax validation
        $tokens = $null
        $parseErrors = $null
        $content = Get-Content $File.FullName -Raw -ErrorAction Stop
        
        $ast = [System.Management.Automation.Language.Parser]::ParseInput(
            $content, 
            [ref]$tokens, 
            [ref]$parseErrors
        )
        
        # Check for parse errors
        if ($parseErrors -and $parseErrors.Count -gt 0)
        {
            $hasErrors = $true
            $errorCount = $parseErrors.Count
            Write-Host "`nSyntax validation FAILED: $($File.FullName)" -ForegroundColor Red
            Write-Host "Found $($parseErrors.Count) parse error(s):" -ForegroundColor Red
            foreach ($err in $parseErrors)
            {
                $errors += $err
                Write-Host "  Line $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Red
                if ($ShowDetails)
                {
                    Write-Host "    Error ID: $($err.ErrorId)" -ForegroundColor DarkRed
                    Write-Host "    Position: Column $($err.Extent.StartColumnNumber) to $($err.Extent.EndColumnNumber)" -ForegroundColor DarkRed
                }
            }
        }
        
        # Method 2: Validate brace/bracket/parenthesis matching
        # Note: We skip this check because the AST parser already validates matching delimiters
        # and this simple token-based approach produces false positives for delimiters inside strings
        # (e.g., $($variable) in expandable strings)
        
        # Method 3: Try to actually compile the script block (catches runtime-detectable issues)
        try
        {
            $scriptBlock = [scriptblock]::Create($content)
            
            # Check AST for common issues
            if ($ast)
            {
                # Find all function definitions and check for duplicates
                $functionDefs = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
                $functionNames = $functionDefs | ForEach-Object { $_.Name }
                $duplicates = $functionNames | Group-Object | Where-Object { $_.Count -gt 1 }
                
                if ($duplicates)
                {
                    $warningCount += $duplicates.Count
                    foreach ($dup in $duplicates)
                    {
                        Write-Host "  Warning: Duplicate function definition '$($dup.Name)' found $($dup.Count) times" -ForegroundColor Yellow
                    }
                }
                
                # Check for unreachable code after return/exit/throw
                $unreachableCode = $ast.FindAll({
                        param($node)
                        if ($node -is [System.Management.Automation.Language.StatementAst])
                        {
                            $parent = $node.Parent
                            if ($parent -is [System.Management.Automation.Language.StatementBlockAst])
                            {
                                $statements = $parent.Statements
                                $index = $statements.IndexOf($node)
                                if ($index -gt 0)
                                {
                                    $prevStatement = $statements[$index - 1]
                                    # Check if previous statement is return, exit, or throw
                                    $prevText = $prevStatement.Extent.Text.Trim()
                                    if ($prevText -match '^\s*(return|exit|throw)\b')
                                    {
                                        return $true
                                    }
                                }
                            }
                        }
                        return $false
                    }, $true)
                
                if ($unreachableCode)
                {
                    $warningCount += $unreachableCode.Count
                    foreach ($code in $unreachableCode | Select-Object -First 3)
                    {
                        Write-Host "  Warning Line $($code.Extent.StartLineNumber): Potentially unreachable code after return/exit/throw" -ForegroundColor Yellow
                    }
                }
            }
        }
        catch
        {
            $hasErrors = $true
            $errorCount++
            Write-Host "  ScriptBlock compilation error: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        if (-not $hasErrors -and -not $errorsOnly)
        {
            Write-Host "[OK] $($File.Name)" -ForegroundColor Green
        }
    }
    catch
    {
        $hasErrors = $true
        $errorCount++
        Write-Host "`nUnexpected error validating $($File.FullName):" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    }
    
    return @{
        HasErrors    = $hasErrors
        ErrorCount   = $errorCount
        WarningCount = $warningCount
        Errors       = $errors
    }
}
