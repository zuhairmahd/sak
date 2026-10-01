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
