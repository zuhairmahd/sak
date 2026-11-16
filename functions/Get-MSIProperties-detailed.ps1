function Get-MSIProperties {
    <#
    .SYNOPSIS
        Lists every property an MSI defines or uses, with scope (public/private),
        hidden and secure flags, authored default, UI-offered values, and runtime assignments.

    .DESCRIPTION
        The Property table only holds properties that have an authored default value.
        Many important properties (often public ones meant for the command line) have no
        default and are only defined by the UI, AppSearch, Upgrade, custom actions, or
        conditions. This function collects all of those sources.

        One object is returned per property per file.

        Scope    : Public if the name has no lowercase letters, otherwise Private
                   (Windows Installer naming rule; only public properties can be set
                   on the msiexec command line).
        Hidden   : True if the value is suppressed from the install log - listed in
                   MsiHiddenProperties, bound to a password edit control, or assigned
                   by a custom action with the HideTarget bit.
        Secure   : True if listed in SecureCustomProperties (allowed to pass from the
                   client to the server side in managed/elevated installs).

    .PARAMETER FilePath
        One or more MSI file paths. Accepts pipeline input, including from Get-ChildItem.
        Paths are taken literally (no wildcards); use Get-ChildItem *.msi | Get-MSIProperties
        for wildcard matching.

    .PARAMETER SkipReferenceScan
        Do not scan conditions and formatted strings for properties that are only referenced.
        Use this to cut down output to properties that are defined or assigned.

    .EXAMPLE
        Get-MSIProperties .\setup.msi | Where-Object Scope -eq 'Public' | Format-List

    .EXAMPLE
        # Show the values a UI offers for each property
        Get-MSIProperties .\setup.msi |
            Where-Object { $_.PossibleValues.Count -gt 0 } |
            ForEach-Object { "Property $($_.Name):"; $_.PossibleValues | Format-List Value, Text, Source }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]]$FilePath,

        [switch]$SkipReferenceScan
    )

    begin {
        # Windows Installer property names are case-sensitive, so every lookup uses ordinal comparison.
        $Ordinal = [System.StringComparer]::Ordinal
        $PropertyNamePattern = '^[A-Za-z_][A-Za-z0-9_.]*$'
        $LogicalOperators = 'NOT', 'AND', 'OR', 'XOR', 'EQV', 'IMP'

        # Columns of type Condition (property names appear bare, e.g. "Installed AND NOT REMOVE")
        $ConditionSources = @(
            'LaunchCondition.Condition', 'Condition.Condition', 'Component.Condition',
            'ControlCondition.Condition', 'ControlEvent.Condition',
            'InstallExecuteSequence.Condition', 'InstallUISequence.Condition',
            'AdminExecuteSequence.Condition', 'AdminUISequence.Condition',
            'AdvtExecuteSequence.Condition'
        )

        # Columns of type Formatted (property names appear in brackets, e.g. "[INSTALLDIR]")
        $FormattedSources = @(
            'Property.Value', 'CustomAction.Target', 'Registry.Value', 'Shortcut.Arguments',
            'Environment.Value', 'IniFile.Value', 'ServiceInstall.Arguments',
            'LaunchCondition.Description', 'Control.Text'
        )

        # Late-bound COM helpers. Windows PowerShell 5.1 cannot call the
        # WindowsInstaller.Installer members directly, so InvokeMember is used.
        function Invoke-MsiMethod($Object, [string]$Name, [object[]]$Arguments) {
            $Object.GetType().InvokeMember($Name, 'InvokeMethod', $null, $Object, $Arguments)
        }

        function Get-MsiTableRows($Database, [string]$Table, [string[]]$Columns) {
            $columnList = ($Columns | ForEach-Object { '`' + $_ + '`' }) -join ', '
            $sql = 'SELECT {0} FROM `{1}`' -f $columnList, $Table
            $view = Invoke-MsiMethod $Database 'OpenView' @($sql)
            try {
                [void](Invoke-MsiMethod $view 'Execute' $null)
                while ($null -ne ($record = Invoke-MsiMethod $view 'Fetch' $null)) {
                    try {
                        $row = [ordered]@{}
                        for ($i = 0; $i -lt $Columns.Count; $i++) {
                            $row[$Columns[$i]] = $record.GetType().InvokeMember(
                                'StringData', 'GetProperty', $null, $record, @($i + 1))
                        }
                        [pscustomobject]$row
                    }
                    finally {
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($record)
                    }
                }
            }
            finally {
                [void](Invoke-MsiMethod $view 'Close' $null)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view)
            }
        }

        function Add-PropertySource($Entries, [string]$Name, [string]$Source) {
            if ([string]::IsNullOrWhiteSpace($Name) -or $Name -notmatch $PropertyNamePattern) { return $null }
            if (-not $Entries.ContainsKey($Name)) {
                $Entries[$Name] = [pscustomobject]@{
                    DefaultValue       = $null
                    HasDefault         = $false
                    DefinedIn          = [System.Collections.Generic.SortedSet[string]]::new()
                    HiddenReasons      = [System.Collections.Generic.SortedSet[string]]::new()
                    PossibleValues     = [System.Collections.Generic.List[object]]::new()
                    RuntimeAssignments = [System.Collections.Generic.List[object]]::new()
                    Controls           = [System.Collections.Generic.List[string]]::new()
                }
            }
            $entry = $Entries[$Name]
            [void]$entry.DefinedIn.Add($Source)
            $entry
        }

        function Get-ConditionReferences([string]$Condition) {
            if ([string]::IsNullOrWhiteSpace($Condition)) { return }
            # Drop string literals, then take bare identifiers that are not prefixed by
            # feature/component/environment state operators ($ ! & ? %).
            $stripped = [regex]::Replace($Condition, '"[^"]*"', '""')
            foreach ($m in [regex]::Matches($stripped, '(?<![\w.$!&?%])[A-Za-z_][\w.]*')) {
                if ($m.Value -notin $LogicalOperators) { $m.Value }
            }
        }

        function Get-FormattedReferences([string]$Text) {
            if ([string]::IsNullOrEmpty($Text)) { return }
            # [NAME] only; [#file], [!file], [$comp], [%env], [\x], [~] are excluded by the pattern.
            foreach ($m in [regex]::Matches($Text, '\[([A-Za-z_][\w.]*)\]')) { $m.Groups[1].Value }
        }

        function Get-PropertyList($Entries, [string]$Name) {
            $set = [System.Collections.Generic.HashSet[string]]::new($Ordinal)
            if ($Entries.ContainsKey($Name) -and $Entries[$Name].DefaultValue) {
                foreach ($item in $Entries[$Name].DefaultValue -split ';') {
                    if ($item.Trim()) { [void]$set.Add($item.Trim()) }
                }
            }
            , $set
        }

        $Installer = New-Object -ComObject WindowsInstaller.Installer
    }

    process {
        foreach ($File in $FilePath) {
            $Database = $null
            try {
                # ProviderPath turns PSDrive paths into real file system / UNC paths COM can open.
                $FullPath = (Resolve-Path -LiteralPath $File -ErrorAction Stop).ProviderPath
                if (-not (Test-Path -LiteralPath $FullPath -PathType Leaf)) {
                    throw "Path is not a file."
                }

                # 0 = msiOpenDatabaseModeReadOnly
                $Database = Invoke-MsiMethod $Installer 'OpenDatabase' @($FullPath, 0)

                $tables = [System.Collections.Generic.HashSet[string]]::new($Ordinal)
                foreach ($row in Get-MsiTableRows $Database '_Tables' @('Name')) { [void]$tables.Add($row.Name) }

                $entries = [System.Collections.Generic.Dictionary[string, object]]::new($Ordinal)

                # Property table: properties with an authored default value
                if ($tables.Contains('Property')) {
                    foreach ($row in Get-MsiTableRows $Database 'Property' @('Property', 'Value')) {
                        $entry = Add-PropertySource $entries $row.Property 'Property'
                        if ($entry) {
                            $entry.DefaultValue = $row.Value
                            $entry.HasDefault = $true
                        }
                    }
                }

                # AppSearch: property receives the result of a file/registry/component search
                if ($tables.Contains('AppSearch')) {
                    foreach ($row in Get-MsiTableRows $Database 'AppSearch' @('Property', 'Signature_')) {
                        $entry = Add-PropertySource $entries $row.Property 'AppSearch'
                        if ($entry) {
                            $entry.RuntimeAssignments.Add([pscustomobject]@{
                                Source = 'AppSearch'
                                Detail = "Search result for signature $($row.Signature_)"
                            })
                        }
                    }
                }

                # Upgrade: ActionProperty receives the ProductCodes of related products found by FindRelatedProducts
                if ($tables.Contains('Upgrade')) {
                    foreach ($row in Get-MsiTableRows $Database 'Upgrade' @('UpgradeCode', 'ActionProperty')) {
                        $entry = Add-PropertySource $entries $row.ActionProperty 'Upgrade'
                        if ($entry) {
                            $entry.RuntimeAssignments.Add([pscustomobject]@{
                                Source = 'Upgrade'
                                Detail = "ProductCodes found for UpgradeCode $($row.UpgradeCode)"
                            })
                        }
                    }
                }

                # Custom actions type 51 (set property) and type 35 (set directory)
                if ($tables.Contains('CustomAction')) {
                    foreach ($row in Get-MsiTableRows $Database 'CustomAction' @('Action', 'Type', 'Source', 'Target')) {
                        $type = [int]$row.Type
                        if (($type -band 0x3F) -in 51, 35) {
                            $entry = Add-PropertySource $entries $row.Source 'CustomAction'
                            if ($entry) {
                                $entry.RuntimeAssignments.Add([pscustomobject]@{
                                    Source = "CustomAction $($row.Action)"
                                    Detail = $row.Target
                                })
                                # msidbCustomActionTypeHideTarget
                                if ($type -band 0x2000) { [void]$entry.HiddenReasons.Add("HideTarget on $($row.Action)") }
                            }
                        }
                    }
                }

                # Directory table: public directory IDs (e.g. INSTALLFOLDER) can be set on the command line
                if ($tables.Contains('Directory')) {
                    foreach ($row in Get-MsiTableRows $Database 'Directory' @('Directory')) {
                        if ($row.Directory -cnotmatch '[a-z]') {
                            [void](Add-PropertySource $entries $row.Directory 'Directory')
                        }
                    }
                }

                # Control table: properties bound to dialog controls
                if ($tables.Contains('Control')) {
                    foreach ($row in Get-MsiTableRows $Database 'Control' @('Dialog_', 'Control', 'Type', 'Attributes', 'Property')) {
                        $entry = Add-PropertySource $entries $row.Property 'Control'
                        if (-not $entry) { continue }
                        $entry.Controls.Add("$($row.Dialog_)/$($row.Control) ($($row.Type))")
                        # msidbControlAttributesPasswordInput
                        if ($row.Attributes -and ([int]$row.Attributes -band 0x200000)) {
                            [void]$entry.HiddenReasons.Add('Password control')
                        }
                    }
                }

                # Option tables: the values a UI control offers
                foreach ($optionTable in 'RadioButton', 'ComboBox', 'ListBox', 'ListView') {
                    if (-not $tables.Contains($optionTable)) { continue }
                    foreach ($row in Get-MsiTableRows $Database $optionTable @('Property', 'Order', 'Value', 'Text')) {
                        $entry = Add-PropertySource $entries $row.Property $optionTable
                        if ($entry) {
                            $entry.PossibleValues.Add([pscustomobject]@{
                                Value  = $row.Value
                                Text   = $row.Text
                                Source = $optionTable
                                Order  = [int]$row.Order
                            })
                        }
                    }
                }

                # CheckBox table: value set when checked; the property is unset when unchecked
                if ($tables.Contains('CheckBox')) {
                    foreach ($row in Get-MsiTableRows $Database 'CheckBox' @('Property', 'Value')) {
                        $entry = Add-PropertySource $entries $row.Property 'CheckBox'
                        if ($entry) {
                            $entry.PossibleValues.Add([pscustomobject]@{
                                Value  = $row.Value
                                Text   = 'Checked (property is unset when unchecked)'
                                Source = 'CheckBox'
                                Order  = 0
                            })
                        }
                    }
                }

                # ControlEvent: an Event of "[NAME]" sets that property when the control is activated
                if ($tables.Contains('ControlEvent')) {
                    foreach ($row in Get-MsiTableRows $Database 'ControlEvent' @('Dialog_', 'Control_', 'Event', 'Argument')) {
                        if ($row.Event -match '^\[([A-Za-z_][\w.]*)\]$') {
                            $entry = Add-PropertySource $entries $Matches[1] 'ControlEvent'
                            if ($entry) {
                                $entry.RuntimeAssignments.Add([pscustomobject]@{
                                    Source = "ControlEvent $($row.Dialog_)/$($row.Control_)"
                                    Detail = if ($row.Argument -eq '{}') { '(unset)' } else { $row.Argument }
                                })
                            }
                        }
                    }
                }

                # Properties that are only referenced (conditions, formatted strings)
                if (-not $SkipReferenceScan) {
                    $scans = @(
                        foreach ($s in $ConditionSources) { @{ Spec = $s; Kind = 'Condition' } }
                        foreach ($s in $FormattedSources) { @{ Spec = $s; Kind = 'Formatted' } }
                    )
                    foreach ($scan in $scans) {
                        $table, $column = $scan.Spec -split '\.', 2
                        if (-not $tables.Contains($table)) { continue }
                        try {
                            foreach ($row in Get-MsiTableRows $Database $table @($column)) {
                                $names = if ($scan.Kind -eq 'Condition') {
                                    Get-ConditionReferences $row.$column
                                } else {
                                    Get-FormattedReferences $row.$column
                                }
                                foreach ($name in $names) {
                                    [void](Add-PropertySource $entries $name "$($scan.Kind) reference")
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Skipped $($scan.Spec) in '$FullPath': $($_.Exception.GetBaseException().Message)"
                        }
                    }
                }

                $hiddenList = Get-PropertyList $entries 'MsiHiddenProperties'
                $secureList = Get-PropertyList $entries 'SecureCustomProperties'
                $fileName = Split-Path -Path $FullPath -Leaf

                foreach ($name in ($entries.Keys | Sort-Object)) {
                    $entry = $entries[$name]
                    if ($hiddenList.Contains($name)) { [void]$entry.HiddenReasons.Add('MsiHiddenProperties') }
                    $scope = if ($name -cmatch '[a-z]') { 'Private' } else { 'Public' }

                    [pscustomobject]@{
                        FileName           = $fileName
                        Name               = $name
                        Scope              = $scope
                        Hidden             = $entry.HiddenReasons.Count -gt 0
                        HiddenReasons      = [string[]]@($entry.HiddenReasons)
                        Secure             = $secureList.Contains($name)
                        HasDefault         = $entry.HasDefault
                        DefaultValue       = $entry.DefaultValue
                        PossibleValues     = @($entry.PossibleValues | Sort-Object Source, Order)
                        RuntimeAssignments = @($entry.RuntimeAssignments)
                        Controls           = [string[]]@($entry.Controls)
                        DefinedIn          = [string[]]@($entry.DefinedIn)
                        FilePath           = $FullPath
                    }
                }
            }
            catch {
                Write-Error -Message "Failed to read MSI file '$File': $($_.Exception.GetBaseException().Message)" -TargetObject $File
            }
            finally {
                if ($Database) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Database) }
            }
        }
    }

    end {
        if ($Installer) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Installer) }
    }
}
