function Get-EmailAddresses()
{
    [CmdletBinding()                            ]
    param(
        [array]$lines
    )
    # Process all content to extract ALL email addresses
    $pattern = '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'
    $seen = @{}
    $result = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines)
    {
        if ([string]::IsNullOrWhiteSpace($line))
        {
            continue
        }
    
        # Extract ALL email matches from the line
        $stringMatches = [regex]::Matches($line, $pattern)

        foreach ($match in $stringMatches)
        {
            $email = $match.Value
            $key = $email.ToLower()
        
            if (-not $seen.ContainsKey($key))
            {
                $seen[$key] = $true
                [void]$result.Add($email)
            }
        }
    }

    if ($result.Count -eq 0)
    {
        Write-Output "No email addresses found."
    }
    #sort the result in alphabetical order
    $result = $result | Sort-Object -Unique          
    return $result
}
