[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$FilePath
)

function Get-MSIProperties()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$FilePath
    )

    $Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($File in $FilePath)
    {
        try
        {
            # Ensure the path is absolute
            $FullPath = (Resolve-Path $File).Path

            # Create the Windows Installer Object
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer

            # Open the MSI database (0 = Read-only)
            $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($FullPath, 0))

            # Query all properties from the Property table
            $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, "SELECT Property, Value FROM Property")
            $null = $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)

            $Properties = [ordered]@{
                FileName = Split-Path $FullPath -Leaf
            }

            while ($null -ne ($Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)))
            {
                $Name = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
                $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 2)
                $Properties[$Name] = $Value
            }

            $Results.Add([PSCustomObject]$Properties)

            # Cleanup COM objects
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($View) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Database) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        }
        catch
        {
            Write-Error "Failed to read MSI file '$File': $($_.Exception.Message)"
        }
    }

    return , $Results.ToArray()
}

# Usage:
$global:MSIProperties = Get-MSIProperties -FilePath $FilePath
#display in human readable form.
$global:MSIProperties | Format-List
