Add-Type -AssemblyName System.Xml.Linq

function Get-GpoXmlSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('FullName')]
        [string[]]$Path
    )

    begin {
        function Normalize-GpoValue ([string]$val) {
            if ([string]::IsNullOrWhiteSpace($val)) { return "Not Configured" }
            return $val.Trim()
        }

        # Maps numeric legacy audit values (0-3) to their display names
        function Convert-AuditValue ([string]$val) {
            switch ($val) {
                '0' { return 'No Auditing' }
                '1' { return 'Success' }
                '2' { return 'Failure' }
                '3' { return 'Success and Failure' }
                default { return Normalize-GpoValue $val }
            }
        }

        # Serializes Policy parameter sub-elements (CheckBox, DropDownList, EditText, ListBox, etc.)
        function Get-PolicyParameters ([System.Xml.Linq.XElement]$policyNode) {
            $structural = [System.Collections.Generic.HashSet[string]]@(
                'Name', 'State', 'Supported', 'Category', 'Explain', 'Annotation'
            )
            $parts = foreach ($child in $policyNode.Elements()) {
                if ($structural.Contains($child.Name.LocalName)) { continue }
                # Name lives in a child <Name> element; fall back to the element's local name
                $nameEl = $child.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1
                $paramName = if ($nameEl -and -not [string]::IsNullOrWhiteSpace($nameEl.Value)) { $nameEl.Value.Trim() } else { $child.Name.LocalName }
                $paramVal = switch ($child.Name.LocalName) {
                    'CheckBox' {
                        ($child.Elements() | Where-Object { $_.Name.LocalName -eq 'State' } | Select-Object -First 1).Value
                    }
                    'DropDownList' {
                        # Selected option is inside <Value><Name>...
                        $valEl = $child.Elements() | Where-Object { $_.Name.LocalName -eq 'Value' } | Select-Object -First 1
                        if ($valEl) { ($valEl.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value } else { $null }
                    }
                    'ListBox' {
                        ($child.Descendants() | Where-Object { $_.Name.LocalName -eq 'Element' } |
                        ForEach-Object { $_.Value }) -join '; '
                    }
                    default {
                        # EditText, Numeric, Text: value is in <Value> child or direct text
                        $valEl = $child.Elements() | Where-Object { $_.Name.LocalName -eq 'Value' } | Select-Object -First 1
                        if ($valEl) { $valEl.Value } else { $child.Value }
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($paramName) -and -not [string]::IsNullOrWhiteSpace($paramVal)) {
                    "$paramName=$paramVal"
                }
            }
            return ($parts | Where-Object { $_ }) -join '; '
        }
    }

    process {
        foreach ($filePath in $Path) {
            $resolvedFiles = Resolve-Path -Path $filePath -ErrorAction SilentlyContinue
            if (-not $resolvedFiles) {
                Write-Warning "File not found or could not be resolved: $filePath"
                continue
            }

            foreach ($file in $resolvedFiles) {
                $settings = [System.Xml.XmlReaderSettings]::new()
                $settings.IgnoreWhitespace = $true

                $reader = [System.Xml.XmlReader]::Create($file.Path, $settings)
                $gpoName = $null
                $configScope = "Computer Configuration"

                try {
                    $reader.MoveToContent() | Out-Null

                    while ($reader.Read()) {
                        if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }

                        # Guard scope updates to depth 1 so nested <User>/<Computer> preference items cannot flip it
                        if ($reader.Depth -eq 1) {
                            if ($reader.LocalName -eq 'Computer') { $configScope = "Computer Configuration" }
                            elseif ($reader.LocalName -eq 'User') { $configScope = "User Configuration" }
                        }

                        # Capture top-level GPO Name (first <Name> seen)
                        if ($reader.LocalName -eq 'Name' -and $null -eq $gpoName) {
                            $gpoName = Normalize-GpoValue ($reader.ReadElementContentAsString())
                        }

                        # 1. Administrative Templates (<Policy>)
                        if ($reader.LocalName -eq 'Policy') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $pName = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                            $pState = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'State' } | Select-Object -First 1).Value
                            $pCat = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Category' } | Select-Object -First 1).Value
                            $pExp = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Explain' } | Select-Object -First 1).Value
                            $pParams = Get-PolicyParameters $node

                            [PSCustomObject]@{
                                FileName     = $file.ProviderPath
                                GPOName      = Normalize-GpoValue $gpoName
                                Scope        = $configScope
                                Category     = "Administrative Templates: " + (Normalize-GpoValue $pCat)
                                PolicyName   = Normalize-GpoValue $pName
                                SettingValue = Normalize-GpoValue $pState
                                Parameters   = Normalize-GpoValue $pParams
                                Details      = Normalize-GpoValue $pExp
                            }
                        }

                        # 2a. Security Settings: SystemAccess and KerberosPolicy (Name + SettingNumber/Boolean/String)
                        elseif ($reader.LocalName -in @('SystemAccess', 'KerberosPolicy')) {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            foreach ($child in $node.Elements()) {
                                $pName = ($child.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                $pVal = ($child.Elements() | Where-Object { $_.Name.LocalName -in @('SettingBoolean', 'SettingNumber', 'SettingString') } | Select-Object -First 1).Value

                                if ($pName) {
                                    [PSCustomObject]@{
                                        FileName     = $file.ProviderPath
                                        GPOName      = Normalize-GpoValue $gpoName
                                        Scope        = $configScope
                                        Category     = "Security: " + $node.Name.LocalName
                                        PolicyName   = Normalize-GpoValue $pName
                                        SettingValue = Normalize-GpoValue $pVal
                                        Parameters   = "Not Configured"
                                        Details      = "Not Configured"
                                    }
                                }
                            }
                        }

                        # 2b. Security Options — each <SecurityOptions> element is one setting (not a container)
                        elseif ($reader.LocalName -eq 'SecurityOptions') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)
                            $keyName = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'KeyName' } | Select-Object -First 1).Value
                            $display = $node.Elements() | Where-Object { $_.Name.LocalName -eq 'Display' } | Select-Object -First 1
                            $displayName = if ($display) { ($display.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value } else { $null }
                            $displayStr = if ($display) { ($display.Elements() | Where-Object { $_.Name.LocalName -eq 'DisplayString' } | Select-Object -First 1).Value } else { $null }
                            # Prefer Display/DisplayString; fall back to SettingNumber/Boolean/String
                            $pVal = if (-not [string]::IsNullOrWhiteSpace($displayStr)) {
                                $displayStr
                            }
                            else {
                                ($node.Elements() | Where-Object { $_.Name.LocalName -in @('SettingNumber', 'SettingBoolean', 'SettingString') } | Select-Object -First 1).Value
                            }
                            # Prefer human-readable Display/Name; fall back to registry KeyName
                            $pName = if (-not [string]::IsNullOrWhiteSpace($displayName)) { $displayName } else { $keyName }

                            if ($pName) {
                                [PSCustomObject]@{
                                    FileName     = $file.ProviderPath
                                    GPOName      = Normalize-GpoValue $gpoName
                                    Scope        = $configScope
                                    Category     = "Security: SecurityOptions"
                                    PolicyName   = Normalize-GpoValue $pName
                                    SettingValue = Normalize-GpoValue $pVal
                                    Parameters   = "Not Configured"
                                    Details      = Normalize-GpoValue $keyName
                                }
                            }
                        }

                        # 2c. Legacy Audit Policy (<EventAuditSetting>) — SettingValue 0-3 mapped to text
                        elseif ($reader.LocalName -eq 'EventAuditSetting') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            foreach ($child in $node.Elements()) {
                                $pName = ($child.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                $pVal = ($child.Elements() | Where-Object { $_.Name.LocalName -eq 'SettingValue' } | Select-Object -First 1).Value

                                if ($pName) {
                                    [PSCustomObject]@{
                                        FileName     = $file.ProviderPath
                                        GPOName      = Normalize-GpoValue $gpoName
                                        Scope        = $configScope
                                        Category     = "Security: EventAuditSetting"
                                        PolicyName   = Normalize-GpoValue $pName
                                        SettingValue = Convert-AuditValue $pVal
                                        Parameters   = "Not Configured"
                                        Details      = "Not Configured"
                                    }
                                }
                            }
                        }

                        # 3. User Rights Assignment (<UserRightsAssignment>)
                        elseif ($reader.LocalName -eq 'UserRightsAssignment') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            foreach ($right in $node.Elements()) {
                                $pName = ($right.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                # Use Member/Name to avoid including the SID that lives alongside it
                                $members = ($right.Elements() | Where-Object { $_.Name.LocalName -eq 'Member' } | ForEach-Object {
                                        ($_.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                    } | Where-Object { $_ }) -join ', '

                                [PSCustomObject]@{
                                    FileName     = $file.ProviderPath
                                    GPOName      = Normalize-GpoValue $gpoName
                                    Scope        = $configScope
                                    Category     = "User Rights Assignment"
                                    PolicyName   = Normalize-GpoValue $pName
                                    SettingValue = Normalize-GpoValue $members
                                    Parameters   = "Not Configured"
                                    Details      = "Not Configured"
                                }
                            }
                        }

                        # 4. Advanced Audit Policies (<AuditSetting>)
                        elseif ($reader.LocalName -eq 'AuditSetting') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $subCat = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'SubcategoryName' } | Select-Object -First 1).Value
                            $incVal = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'SettingValue' } | Select-Object -First 1).Value

                            [PSCustomObject]@{
                                FileName     = $file.ProviderPath
                                GPOName      = Normalize-GpoValue $gpoName
                                Scope        = $configScope
                                Category     = "Advanced Audit Policy"
                                PolicyName   = Normalize-GpoValue $subCat
                                SettingValue = Normalize-GpoValue $incVal
                                Parameters   = "Not Configured"
                                Details      = "Not Configured"
                            }
                        }

                        # 5. GPO Preferences: Registry (<Registry>)
                        elseif ($reader.LocalName -eq 'Registry' -and $reader.NamespaceURI -like "*GroupPolicy*") {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)
                            $props = $node.Elements() | Where-Object { $_.Name.LocalName -eq 'Properties' } | Select-Object -First 1

                            if ($props) {
                                # Use .Value to get the attribute value string, not the full 'name="value"' XAttribute form
                                $action = if ($props.Attribute('action')) { $props.Attribute('action').Value } else { '' }
                                $key = if ($props.Attribute('key')) { $props.Attribute('key').Value } else { '' }
                                $val = if ($props.Attribute('value')) { $props.Attribute('value').Value } else { '' }
                                $data = if ($props.Attribute('data')) { $props.Attribute('data').Value } else { '' }

                                [PSCustomObject]@{
                                    FileName     = $file.ProviderPath
                                    GPOName      = Normalize-GpoValue $gpoName
                                    Scope        = $configScope
                                    Category     = "Preference Registry"
                                    PolicyName   = Normalize-GpoValue $val
                                    SettingValue = Normalize-GpoValue $data
                                    Parameters   = "Not Configured"
                                    Details      = "Action: $(Normalize-GpoValue $action) | Key: $(Normalize-GpoValue $key)"
                                }
                            }
                        }

                        # 6. Windows Firewall Rules (<FirewallRule>)
                        elseif ($reader.LocalName -eq 'FirewallRule') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $ruleName = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'RuleName' } | Select-Object -First 1).Value
                            $action = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Action' } | Select-Object -First 1).Value
                            $dir = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Direction' } | Select-Object -First 1).Value

                            [PSCustomObject]@{
                                FileName     = $file.ProviderPath
                                GPOName      = Normalize-GpoValue $gpoName
                                Scope        = $configScope
                                Category     = "Windows Firewall Rule"
                                PolicyName   = Normalize-GpoValue $ruleName
                                SettingValue = Normalize-GpoValue $action
                                Parameters   = "Not Configured"
                                Details      = "Direction: $(Normalize-GpoValue $dir)"
                            }
                        }

                        # 7. EFS / Public Key Settings (<EFSSettings>) — scalar fields only; certificate blobs skipped
                        elseif ($reader.LocalName -eq 'EFSSettings') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            foreach ($child in $node.Elements()) {
                                if ($child.HasElements) { continue }
                                [PSCustomObject]@{
                                    FileName     = $file.ProviderPath
                                    GPOName      = Normalize-GpoValue $gpoName
                                    Scope        = $configScope
                                    Category     = "Security: EFSSettings"
                                    PolicyName   = $child.Name.LocalName
                                    SettingValue = Normalize-GpoValue $child.Value
                                    Parameters   = "Not Configured"
                                    Details      = "Not Configured"
                                }
                            }
                        }

                        # 8. Restricted Groups (<RestrictedGroup>)
                        elseif ($reader.LocalName -eq 'RestrictedGroup') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $groupName = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'GroupName' } | ForEach-Object {
                                    ($_.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                } | Select-Object -First 1)
                            $members = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Member' } | ForEach-Object {
                                    ($_.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                } | Where-Object { $_ }) -join ', '

                            [PSCustomObject]@{
                                FileName     = $file.ProviderPath
                                GPOName      = Normalize-GpoValue $gpoName
                                Scope        = $configScope
                                Category     = "Restricted Groups"
                                PolicyName   = Normalize-GpoValue $groupName
                                SettingValue = Normalize-GpoValue $members
                                Parameters   = "Not Configured"
                                Details      = "Not Configured"
                            }
                        }

                        # 9. Logon / Startup Scripts (<Script>)
                        elseif ($reader.LocalName -eq 'Script') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $scriptType = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Type' } | Select-Object -First 1).Value
                            $cmdLine = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'CmdLine' } | Select-Object -First 1).Value
                            $scriptParams = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Parameters' } | Select-Object -First 1).Value

                            [PSCustomObject]@{
                                FileName     = $file.ProviderPath
                                GPOName      = Normalize-GpoValue $gpoName
                                Scope        = $configScope
                                Category     = "Script"
                                PolicyName   = Normalize-GpoValue $scriptType
                                SettingValue = Normalize-GpoValue $cmdLine
                                Parameters   = Normalize-GpoValue $scriptParams
                                Details      = "Not Configured"
                            }
                        }

                        # 10. System Services (<Service>)
                        elseif ($reader.LocalName -eq 'Service') {
                            $node = [System.Xml.Linq.XElement]::ReadFrom($reader)

                            $props = $node.Elements() | Where-Object { $_.Name.LocalName -eq 'Properties' } | Select-Object -First 1
                            if ($props) {
                                # Preferences-style Service (attributes on <Properties>)
                                $svcName = [string]$props.Attribute('serviceName')
                                $startupType = [string]$props.Attribute('startupType')
                            }
                            else {
                                # Security Settings-style Service (child elements)
                                $svcName = ($node.Elements() | Where-Object { $_.Name.LocalName -eq 'Name' } | Select-Object -First 1).Value
                                $startupType = ($node.Elements() | Where-Object { $_.Name.LocalName -in @('StartupType', 'Startup') } | Select-Object -First 1).Value
                            }

                            if ($svcName) {
                                [PSCustomObject]@{
                                    FileName     = $file.ProviderPath
                                    GPOName      = Normalize-GpoValue $gpoName
                                    Scope        = $configScope
                                    Category     = "System Service"
                                    PolicyName   = Normalize-GpoValue $svcName
                                    SettingValue = Normalize-GpoValue $startupType
                                    Parameters   = "Not Configured"
                                    Details      = "Not Configured"
                                }
                            }
                        }
                    }
                }
                finally {
                    $reader.Close()
                    $reader.Dispose()
                }
            }
        }
    }
}

$filesToProcess = @(
    (Join-Path $PSScriptRoot "m365customizations.xml"),
    (Join-Path $PSScriptRoot "dodUser.xml"),
    (Join-Path $PSScriptRoot "dodComputer.xml")
)
$global:GPOSettings = Get-GpoXmlSetting -Path $filesToProcess

