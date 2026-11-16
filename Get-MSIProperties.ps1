function Get-MSIProperties {
    <#
    .SYNOPSIS
        Lists every property an MSI defines, sets, or reads, with scope (public/private),
        hidden and secure flags, authored default, possible values, and where each was found.

    .DESCRIPTION
        The Property table only holds properties that have an authored default value.
        Many command-line options have no default and exist only as references elsewhere.
        This function collects them from:

          - The Property, AppSearch, Upgrade, Directory, Control, CheckBox, RadioButton,
            ComboBox, ListBox, ListView, ControlEvent and Feature tables.
          - Every string column of every table, including vendor and WiX extension tables:
            [NAME] references in formatted columns, and bare names in Condition columns.
          - Columns whose _Validation category is Property (the value is a property name).
          - Script custom actions (VBScript/JScript, inline or stored in the Binary table):
            Session.Property("NAME") calls.
          - With -ScanBinaries: strings inside DLL/EXE custom actions stored in the Binary
            table, including managed (DTF) custom actions packed in an embedded cabinet.
            This is the only place properties read by compiled code (MsiGetProperty) show up.
          - With -IncludeStandardProperties: documented Windows Installer properties that
            every package accepts on the command line, with their documented values.

        One object is returned per property per file.

        Scope      : Public if the name has no lowercase letters, otherwise Private.
                     Only public properties can be set on the msiexec command line.
        Hidden     : Value is kept out of the install log (MsiHiddenProperties, a password
                     edit control, or a custom action with the HideTarget bit).
        Secure     : Listed in SecureCustomProperties (passed to the server side in
                     managed or elevated installs).
        Confidence : High   - found in database tables.
                     Medium - found only in script custom action code.
                     Low    - found only as a string inside a compiled custom action binary.
                              Expect noise; check these against vendor documentation.

    .PARAMETER FilePath
        One or more MSI file paths. Accepts pipeline input, including from Get-ChildItem.

    .PARAMETER TransformPath
        One or more transforms (.mst) to apply before reading, in order. The MSI is copied
        to a temporary file first; the original is never modified.

    .PARAMETER ScanBinaries
        Extract strings from DLL and EXE custom actions in the Binary table and report
        upper-case identifiers as possible properties (Confidence Low). Requires Windows
        (uses expand.exe for managed custom action cabinets).

    .PARAMETER IncludeStandardProperties
        Add documented Windows Installer properties (ALLUSERS, REBOOT, REINSTALLMODE, and
        so on) with their documented values, even if the package does not reference them.

    .PARAMETER SkipReferenceScan
        Do not scan table columns for referenced properties. Output is limited to
        properties that are defined, assigned, or offered by the UI.

    .EXAMPLE
        Get-MSIProperties .\setup.msi | Where-Object Scope -eq 'Public' | Format-List Name, Confidence, DefaultValue, DefinedIn

    .EXAMPLE
        Get-MSIProperties .\setup.msi -ScanBinaries |
            Where-Object { $_.Scope -eq 'Public' -and $_.Confidence -eq 'Low' } |
            Format-List Name, DefinedIn

    .EXAMPLE
        # Show the values offered for each property
        Get-MSIProperties .\setup.msi -IncludeStandardProperties |
            Where-Object { $_.PossibleValues.Count -gt 0 } |
            ForEach-Object { "Property $($_.Name):"; $_.PossibleValues | Format-List Value, Text, Source }

    .EXAMPLE
        Get-MSIProperties .\setup.msi -TransformPath .\custom.mst | Format-List Name, DefaultValue
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]]$FilePath,

        [string[]]$TransformPath,

        [switch]$ScanBinaries,

        [switch]$IncludeStandardProperties,

        [switch]$SkipReferenceScan
    )

    begin {
        # Windows Installer property names are case-sensitive, so every lookup uses ordinal comparison.
        $Ordinal = [System.StringComparer]::Ordinal
        $PropertyNamePattern = '^[A-Za-z_][A-Za-z0-9_.]*$'
        $LogicalOperators = 'NOT', 'AND', 'OR', 'XOR', 'EQV', 'IMP'

        # Columns that hold a property name; combined with any found in the _Validation table.
        $FallbackPropertyColumns = @(
            'Control.Property', 'CheckBox.Property', 'RadioButton.Property', 'ComboBox.Property',
            'ListBox.Property', 'ListView.Property', 'AppSearch.Property', 'Upgrade.ActionProperty',
            'RemoveFile.DirProperty', 'IniFile.DirProperty', 'RemoveIniFile.DirProperty',
            'MoveFile.SourceFolder', 'MoveFile.DestFolder', 'DuplicateFile.DestFolder', 'Shortcut.WkDir'
        )

        # Properties whose value is a list of features from the Feature table.
        $FeatureListProperties = 'ADDLOCAL', 'ADDSOURCE', 'ADDDEFAULT', 'ADVERTISE', 'REMOVE', 'REINSTALL'

        # Documented Windows Installer properties: Name|Value|Meaning. A value of * means free-form.
        $StandardPropertyData = @'
ALLUSERS||Not set: per-user installation context
ALLUSERS|1|Per-machine installation context
ALLUSERS|2|Per-user or per-machine, decided by privileges and MSIINSTALLPERUSER
MSIINSTALLPERUSER||With ALLUSERS=2: per-machine installation
MSIINSTALLPERUSER|1|With ALLUSERS=2: per-user installation
REBOOT|Force|Always prompt for a restart at the end; restarts automatically when there is no UI
REBOOT|Suppress|Suppress restart prompts at the end of the installation; ForceReboot still prompts
REBOOT|ReallySuppress|Suppress all restarts and restart prompts
MSIRESTARTMANAGERCONTROL|0|Default: always attempt to use Restart Manager
MSIRESTARTMANAGERCONTROL|Disable|Do not use Restart Manager; use the FilesInUse dialog
MSIRESTARTMANAGERCONTROL|DisableShutdown|Use Restart Manager only to detect files in use
MSIDISABLERMRESTART|0|Restart services and registered applications that were shut down
MSIDISABLERMRESTART|1|Do not restart processes shut down by Restart Manager
MSIRMSHUTDOWN|0|Shut down processes using affected files
MSIRMSHUTDOWN|1|Also force shutdown of processes that do not respond to Restart Manager
MSIRMSHUTDOWN|2|Shut down only if every affected process is registered for restart
MSIFASTINSTALL|0|Default
MSIFASTINSTALL|1|No system restore point; values can be added together, for example 7
MSIFASTINSTALL|2|Perform only file costing
MSIFASTINSTALL|4|Reduce the frequency of progress messages
REINSTALLMODE|omus|Default; letters can be combined
REINSTALLMODE|p|Reinstall only if the file is missing
REINSTALLMODE|o|Reinstall if the file is missing or older
REINSTALLMODE|e|Reinstall if the file is missing, equal or older
REINSTALLMODE|d|Reinstall if the file is missing or a different version
REINSTALLMODE|c|Reinstall if the checksum is missing or wrong
REINSTALLMODE|a|Force all files to be reinstalled
REINSTALLMODE|u|Rewrite per-user registry entries
REINSTALLMODE|m|Rewrite per-machine registry entries
REINSTALLMODE|s|Reinstall all shortcuts and re-cache icons
REINSTALLMODE|v|Run from source and re-cache the local package
ARPSYSTEMCOMPONENT|1|Hide the product from Programs and Features
ARPNOREMOVE|1|Disable the Remove option in Programs and Features
ARPNOMODIFY|1|Disable the Change option in Programs and Features
ARPNOREPAIR|1|Disable the Repair option in Programs and Features
INSTALLLEVEL|*|Integer; features whose Level is at or below this value are installed
TARGETDIR|*|Root destination directory
TRANSFORMS|*|Semicolon-separated list of transforms to apply
PATCH|*|Semicolon-separated list of patches to apply
ADDLOCAL|ALL|Install all features locally
ADDSOURCE|ALL|Run all features from source
ADDDEFAULT|ALL|Install all features in their default configuration
ADVERTISE|ALL|Advertise all features
REMOVE|ALL|Remove all features (uninstall)
REINSTALL|ALL|Reinstall all features
'@ -split '\r?\n' | Where-Object { $_ }

        # Upper-case strings that commonly appear in binaries and are not properties.
        $BinaryNoise = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
            'KERNEL32', 'USER32', 'ADVAPI32', 'SHELL32', 'OLE32', 'OLEAUT32', 'SHLWAPI', 'VERSION',
            'MSVCRT', 'MSVCP140', 'VCRUNTIME140', 'UCRTBASE', 'NTDLL', 'MSI', 'RSDS', 'GCTL',
            'VS_VERSION_INFO', 'HKEY_LOCAL_MACHINE', 'HKEY_CURRENT_USER', 'HKEY_CLASSES_ROOT',
            'HKEY_USERS', 'HKLM', 'HKCU', 'HKCR', 'HKU', 'SOFTWARE', 'SYSTEM', 'NULL', 'TRUE', 'FALSE',
            'ERROR', 'WARNING', 'INFO', 'DEBUG', 'TRACE', 'UTF', 'ANSI', 'ASCII', 'GUID', 'CLSID',
            'MZP', 'PADDING', 'PADDINGX', 'XPADDING', 'PADDINGPADDINGXXPADDING', 'TEXT', 'DATA',
            'BSS', 'IDATA', 'RDATA', 'PDATA', 'EDATA', 'RELOC', 'RSRC', 'TLS', 'CRT', 'MSCF',
            'MSCOREE', 'COMCTL32', 'GDI32', 'WS2_32', 'CRYPT32', 'WINTRUST', 'SETUPAPI', 'RPCRT4',
            'NETAPI32', 'USERENV', 'WININET', 'WINHTTP', 'SECUR32', 'PSAPI', 'DBGHELP', 'COMDLG32',
            'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'SELECT', 'FROM', 'WHERE', 'ORDER', 'INSERT', 'UPDATE',
            'DELETE', 'VALUES', 'INTO', 'FAILED', 'QNAN', 'SNAN', 'IND', 'INF', 'NAN', 'INITY', 'RTTI',
            'ROOT', 'CIMV2', 'DOS', 'COM', 'LPT', 'CON', 'AUX', 'PRN'
        ), $Ordinal)

        # Late-bound COM helpers. Windows PowerShell 5.1 cannot call the
        # WindowsInstaller.Installer members directly, so InvokeMember is used.
        function Invoke-MsiMethod($Object, [string]$Name, [object[]]$Arguments) {
            $Object.GetType().InvokeMember($Name, 'InvokeMethod', $null, $Object, $Arguments)
        }

        function Get-MsiComProperty($Object, [string]$Name, [object[]]$Arguments) {
            $Object.GetType().InvokeMember($Name, 'GetProperty', $null, $Object, $Arguments)
        }

        function Close-MsiObject($Object) {
            if ($null -ne $Object) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) }
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
                            $row[$Columns[$i]] = Get-MsiComProperty $record 'StringData' @($i + 1)
                        }
                        [pscustomobject]$row
                    }
                    finally { Close-MsiObject $record }
                }
            }
            finally {
                [void](Invoke-MsiMethod $view 'Close' $null)
                Close-MsiObject $view
            }
        }

        # Returns the name and column definition (for example s72, L255, i2, V0) of every column.
        function Get-MsiTableColumns($Database, [string]$Table) {
            $view = Invoke-MsiMethod $Database 'OpenView' @(('SELECT * FROM `{0}`' -f $Table))
            try {
                [void](Invoke-MsiMethod $view 'Execute' $null)
                $names = Get-MsiComProperty $view 'ColumnInfo' @(0)
                $types = Get-MsiComProperty $view 'ColumnInfo' @(1)
                try {
                    $count = Get-MsiComProperty $names 'FieldCount' $null
                    for ($i = 1; $i -le $count; $i++) {
                        [pscustomobject]@{
                            Name = Get-MsiComProperty $names 'StringData' @($i)
                            Type = Get-MsiComProperty $types 'StringData' @($i)
                        }
                    }
                }
                finally {
                    Close-MsiObject $names
                    Close-MsiObject $types
                }
            }
            finally {
                [void](Invoke-MsiMethod $view 'Close' $null)
                Close-MsiObject $view
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

        function Add-PossibleValue($Entry, $Value, [string]$Text, [string]$Source, [int]$Order) {
            $Entry.PossibleValues.Add([pscustomobject]@{
                Value  = $Value
                Text   = $Text
                Source = $Source
                Order  = $Order
            })
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
            # [NAME] only; [#file], [!file], [$comp], [%env], [\x], [~] and [1] are excluded by the pattern.
            foreach ($m in [regex]::Matches($Text, '\[([A-Za-z_][\w.]*)\]')) { $m.Groups[1].Value }
        }

        function Get-ScriptReferences([string]$Script) {
            if ([string]::IsNullOrEmpty($Script)) { return }
            # VBScript and JScript: Session.Property("NAME"), including assignments.
            foreach ($m in [regex]::Matches($Script, '(?i)\bProperty\s*\(\s*"([A-Za-z_][\w.]*)"\s*\)')) {
                $m.Groups[1].Value
            }
        }

        function Get-BinaryStringCandidates([byte[]]$Bytes) {
            $found = [System.Collections.Generic.HashSet[string]]::new($Ordinal)
            if (-not $Bytes -or $Bytes.Length -lt 4) { return , $found }
            # ASCII strings (native code), and UTF-16LE strings at both byte alignments
            # (native wide strings and the .NET user-string heap).
            $texts = @(
                [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
                [System.Text.Encoding]::Unicode.GetString($Bytes)
                [System.Text.Encoding]::Unicode.GetString($Bytes, 1, $Bytes.Length - 1)
            )
            foreach ($text in $texts) {
                # Only whole null-terminated strings; this skips machine code that happens to be printable.
                foreach ($m in [regex]::Matches($text, '(?<=\x00)[A-Z][A-Z0-9_]{3,63}(?=\x00)')) {
                    $value = $m.Value
                    if ($BinaryNoise.Contains($value)) { continue }
                    if ($value -match '^[0-9A-F]{6,}$') { continue }   # hex runs
                    if ($value -notmatch '[A-Z]{2}') { continue }
                    if ($value -match '^(.)\1+$') { continue }         # AAAA, XXXX
                    [void]$found.Add($value)
                }
            }
            , $found
        }

        # Managed (DTF) custom actions are a native wrapper with the assembly packed in a cabinet.
        function Expand-EmbeddedCabinets([byte[]]$Bytes, [string]$Destination) {
            $latin = [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
            $offset = $latin.IndexOf('MSCF', [System.StringComparison]::Ordinal)
            $index = 0
            while ($offset -ge 0 -and $offset + 12 -lt $Bytes.Length) {
                # CFHEADER: signature (4), reserved (4), cbCabinet (4)
                $size = [System.BitConverter]::ToUInt32($Bytes, $offset + 8)
                if ($size -gt 36 -and $offset + $size -le $Bytes.Length) {
                    $cab = Join-Path $Destination ("embedded{0}.cab" -f $index)
                    $out = Join-Path $Destination ("embedded{0}" -f $index)
                    $cabBytes = [byte[]]::new($size)
                    [System.Array]::Copy($Bytes, $offset, $cabBytes, 0, $size)
                    [System.IO.File]::WriteAllBytes($cab, $cabBytes)
                    [void](New-Item -ItemType Directory -Path $out -Force)
                    $expand = Join-Path $env:SystemRoot 'System32\expand.exe'
                    if (Test-Path -LiteralPath $expand) {
                        & $expand -F:* $cab $out *> $null
                        Get-ChildItem -LiteralPath $out -File -Recurse
                    }
                    $index++
                    $offset = $latin.IndexOf('MSCF', $offset + $size, [System.StringComparison]::Ordinal)
                }
                else {
                    $offset = $latin.IndexOf('MSCF', $offset + 4, [System.StringComparison]::Ordinal)
                }
            }
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
            $tempCopy = $null
            $workDir = $null
            try {
                # ProviderPath turns PSDrive paths into real file system / UNC paths COM can open.
                $FullPath = (Resolve-Path -LiteralPath $File -ErrorAction Stop).ProviderPath
                if (-not (Test-Path -LiteralPath $FullPath -PathType Leaf)) {
                    throw "Path is not a file."
                }

                $stream = [System.IO.File]::OpenRead($FullPath)
                try { $signature = @($stream.ReadByte(), $stream.ReadByte()) } finally { $stream.Dispose() }
                if ($signature[0] -eq 0x4D -and $signature[1] -eq 0x5A) {
                    throw ("This is an executable (setup bootstrapper), not an MSI database. Its switches belong to " +
                        "the bootstrapper. Run it with /? or extract the embedded MSI and read that instead.")
                }

                $workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("MsiProps_" + [guid]::NewGuid().ToString('N'))
                [void](New-Item -ItemType Directory -Path $workDir -Force)

                $appliedTransforms = @()
                if ($TransformPath) {
                    # Transforms need a writable database. Work on a copy and never commit.
                    $tempCopy = Join-Path $workDir 'package.msi'
                    Copy-Item -LiteralPath $FullPath -Destination $tempCopy
                    (Get-Item -LiteralPath $tempCopy).IsReadOnly = $false
                    # 1 = msiOpenDatabaseModeTransact
                    $Database = Invoke-MsiMethod $Installer 'OpenDatabase' @($tempCopy, 1)
                    foreach ($transform in $TransformPath) {
                        $transformFull = (Resolve-Path -LiteralPath $transform -ErrorAction Stop).ProviderPath
                        Invoke-MsiMethod $Database 'ApplyTransform' @($transformFull, 0)
                        $appliedTransforms += $transformFull
                    }
                }
                else {
                    # 0 = msiOpenDatabaseModeReadOnly
                    $Database = Invoke-MsiMethod $Installer 'OpenDatabase' @($FullPath, 0)
                }

                $tables = [System.Collections.Generic.HashSet[string]]::new($Ordinal)
                foreach ($row in Get-MsiTableRows $Database '_Tables' @('Name')) { [void]$tables.Add($row.Name) }

                $entries = [System.Collections.Generic.Dictionary[string, object]]::new($Ordinal)

                # Property table: properties with an authored default value
                if ($tables.Contains('Property')) {
                    foreach ($row in Get-MsiTableRows $Database 'Property' @('Property', 'Value')) {
                        $entry = Add-PropertySource $entries $row.Property 'Property table'
                        if ($entry) {
                            $entry.DefaultValue = $row.Value
                            $entry.HasDefault = $true
                        }
                    }
                }

                # AppSearch: property receives the result of a file/registry/component search
                if ($tables.Contains('AppSearch')) {
                    foreach ($row in Get-MsiTableRows $Database 'AppSearch' @('Property', 'Signature_')) {
                        $entry = Add-PropertySource $entries $row.Property 'AppSearch table'
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
                        $entry = Add-PropertySource $entries $row.ActionProperty 'Upgrade table'
                        if ($entry) {
                            $entry.RuntimeAssignments.Add([pscustomobject]@{
                                Source = 'Upgrade'
                                Detail = "ProductCodes found for UpgradeCode $($row.UpgradeCode)"
                            })
                        }
                    }
                }

                # Custom actions: type 51 sets a property, type 35 sets a directory
                $customActions = @()
                if ($tables.Contains('CustomAction')) {
                    $customActions = @(Get-MsiTableRows $Database 'CustomAction' @('Action', 'Type', 'Source', 'Target'))
                    foreach ($row in $customActions) {
                        $type = [int]$row.Type
                        if (($type -band 0x3F) -in 51, 35) {
                            $entry = Add-PropertySource $entries $row.Source 'CustomAction table (sets value)'
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
                            [void](Add-PropertySource $entries $row.Directory 'Directory table')
                        }
                    }
                }

                # Feature table: feature names are the values ADDLOCAL, REMOVE and similar accept
                if ($tables.Contains('Feature')) {
                    $features = @(Get-MsiTableRows $Database 'Feature' @('Feature', 'Title', 'Level'))
                    foreach ($name in $FeatureListProperties) {
                        $entry = Add-PropertySource $entries $name 'Feature table (feature-selection property)'
                        $order = 0
                        foreach ($feature in $features) {
                            $text = if ($feature.Title) { $feature.Title } else { '(no title)' }
                            if ($feature.Level -eq '0') { $text += ' - disabled by default (Level 0)' }
                            Add-PossibleValue $entry $feature.Feature $text 'Feature' (++$order)
                        }
                    }
                }

                # Control table: properties bound to dialog controls
                if ($tables.Contains('Control')) {
                    foreach ($row in Get-MsiTableRows $Database 'Control' @('Dialog_', 'Control', 'Type', 'Attributes', 'Property')) {
                        $entry = Add-PropertySource $entries $row.Property 'Control table'
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
                        $entry = Add-PropertySource $entries $row.Property "$optionTable table"
                        if ($entry) { Add-PossibleValue $entry $row.Value $row.Text $optionTable ([int]$row.Order) }
                    }
                }

                # CheckBox table: value set when checked; the property is unset when unchecked
                if ($tables.Contains('CheckBox')) {
                    foreach ($row in Get-MsiTableRows $Database 'CheckBox' @('Property', 'Value')) {
                        $entry = Add-PropertySource $entries $row.Property 'CheckBox table'
                        if ($entry) {
                            Add-PossibleValue $entry $row.Value 'Checked (property is unset when unchecked)' 'CheckBox' 0
                        }
                    }
                }

                # ControlEvent: an Event of "[NAME]" sets that property when the control is activated
                if ($tables.Contains('ControlEvent')) {
                    foreach ($row in Get-MsiTableRows $Database 'ControlEvent' @('Dialog_', 'Control_', 'Event', 'Argument')) {
                        if ($row.Event -match '^\[([A-Za-z_][\w.]*)\]$') {
                            $entry = Add-PropertySource $entries $Matches[1] 'ControlEvent table (sets value)'
                            if ($entry) {
                                $entry.RuntimeAssignments.Add([pscustomobject]@{
                                    Source = "ControlEvent $($row.Dialog_)/$($row.Control_)"
                                    Detail = if ($row.Argument -eq '{}') { '(unset)' } else { $row.Argument }
                                })
                            }
                        }
                    }
                }

                # Generic scan of every string column in every table, including extension tables
                if (-not $SkipReferenceScan) {
                    $propertyColumns = [System.Collections.Generic.HashSet[string]]::new($Ordinal)
                    if ($tables.Contains('_Validation')) {
                        foreach ($row in Get-MsiTableRows $Database '_Validation' @('Table', 'Column', 'Category')) {
                            if ($row.Category -eq 'Property') { [void]$propertyColumns.Add("$($row.Table).$($row.Column)") }
                        }
                    }
                    # _Validation is often incomplete, so the known list is always included.
                    foreach ($spec in $FallbackPropertyColumns) { [void]$propertyColumns.Add($spec) }

                    foreach ($table in ($tables | Sort-Object)) {
                        if ($table.StartsWith('_')) { continue }
                        try {
                            $columns = @(Get-MsiTableColumns $Database $table | Where-Object { $_.Type -match '^[sSlL]' })
                            if ($columns.Count -eq 0) { continue }
                            $columnNames = [string[]]@($columns.Name)
                            foreach ($row in Get-MsiTableRows $Database $table $columnNames) {
                                foreach ($column in $columnNames) {
                                    $value = $row.$column
                                    if ([string]::IsNullOrEmpty($value)) { continue }
                                    $spec = "$table.$column"
                                    if ($propertyColumns.Contains($spec)) {
                                        [void](Add-PropertySource $entries $value "Property column: $spec")
                                    }
                                    elseif ($column -eq 'Condition') {
                                        foreach ($name in Get-ConditionReferences $value) {
                                            [void](Add-PropertySource $entries $name "Condition: $spec")
                                        }
                                    }
                                    else {
                                        foreach ($name in Get-FormattedReferences $value) {
                                            [void](Add-PropertySource $entries $name "Reference: $spec")
                                        }
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Verbose "Skipped table $table in '$FullPath': $($_.Exception.GetBaseException().Message)"
                        }
                    }
                }

                # Export the Binary table once if any custom action code needs reading.
                # Each stream is written to <workDir>\Binary\<file>.ibd, listed in Binary.idt.
                $binaryFiles = [System.Collections.Generic.Dictionary[string, string]]::new($Ordinal)
                $needsBinary = $customActions | Where-Object {
                    $base = [int]$_.Type -band 0x3F
                    $base -in 5, 6 -or ($ScanBinaries -and $base -in 1, 2)
                }
                if ($needsBinary -and $tables.Contains('Binary')) {
                    try {
                        Invoke-MsiMethod $Database 'Export' @('Binary', $workDir, 'Binary.idt')
                        $idt = Get-Content -LiteralPath (Join-Path $workDir 'Binary.idt')
                        foreach ($line in ($idt | Select-Object -Skip 3)) {
                            $parts = $line -split "`t", 2
                            if ($parts.Count -eq 2) {
                                $binaryFiles[$parts[0]] = Join-Path (Join-Path $workDir 'Binary') $parts[1]
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not export the Binary table from '$FullPath': $($_.Exception.GetBaseException().Message)"
                    }
                }

                # Script custom actions: VBScript/JScript inline, in a property, or in the Binary table
                foreach ($ca in $customActions) {
                    $base = [int]$ca.Type -band 0x3F
                    $script = $null
                    switch ($base) {
                        { $_ -in 37, 38 } { $script = $ca.Target }
                        { $_ -in 53, 54 } { if ($entries.ContainsKey($ca.Source)) { $script = $entries[$ca.Source].DefaultValue } }
                        { $_ -in 5, 6 } {
                            if ($binaryFiles.ContainsKey($ca.Source) -and (Test-Path -LiteralPath $binaryFiles[$ca.Source])) {
                                $script = [System.IO.File]::ReadAllText($binaryFiles[$ca.Source])
                            }
                        }
                    }
                    foreach ($name in Get-ScriptReferences $script) {
                        [void](Add-PropertySource $entries $name "Script: CustomAction $($ca.Action)")
                    }
                }

                # Compiled custom actions: strings in DLL/EXE streams (and DTF cabinets inside them)
                if ($ScanBinaries) {
                    $binaryKeys = $customActions |
                        Where-Object { ([int]$_.Type -band 0x3F) -in 1, 2 } |
                        Select-Object -ExpandProperty Source -Unique
                    foreach ($key in $binaryKeys) {
                        if (-not $binaryFiles.ContainsKey($key) -or -not (Test-Path -LiteralPath $binaryFiles[$key])) { continue }
                        if ($key -match '^Wix') {
                            Write-Verbose "Skipped WiX toolset custom action binary $key (table-driven, not package-specific)."
                            continue
                        }
                        $bytes = [System.IO.File]::ReadAllBytes($binaryFiles[$key])
                        $candidates = Get-BinaryStringCandidates $bytes
                        $cabDir = Join-Path $workDir ("cab_" + ($key -replace '[^\w.-]', '_'))
                        [void](New-Item -ItemType Directory -Path $cabDir -Force)
                        foreach ($inner in Expand-EmbeddedCabinets $bytes $cabDir) {
                            $innerCandidates = Get-BinaryStringCandidates ([System.IO.File]::ReadAllBytes($inner.FullName))
                            $candidates.UnionWith($innerCandidates)
                        }
                        foreach ($name in $candidates) {
                            [void](Add-PropertySource $entries $name "Binary strings: $key")
                        }
                    }
                }

                # Documented Windows Installer properties every package accepts
                if ($IncludeStandardProperties) {
                    $order = 0
                    foreach ($line in $StandardPropertyData) {
                        $name, $value, $text = $line -split '\|', 3
                        $entry = Add-PropertySource $entries $name 'Windows Installer standard property'
                        if ($value -eq '*') { $value = $null }
                        Add-PossibleValue $entry $value $text 'Windows Installer documentation' (++$order)
                    }
                }

                $hiddenList = Get-PropertyList $entries 'MsiHiddenProperties'
                $secureList = Get-PropertyList $entries 'SecureCustomProperties'
                $fileName = Split-Path -Path $FullPath -Leaf

                foreach ($name in ($entries.Keys | Sort-Object)) {
                    $entry = $entries[$name]
                    if ($hiddenList.Contains($name)) { [void]$entry.HiddenReasons.Add('MsiHiddenProperties') }
                    $scope = if ($name -cmatch '[a-z]') { 'Private' } else { 'Public' }
                    $sources = @($entry.DefinedIn)
                    $confidence = if ($sources | Where-Object { $_ -notlike 'Script:*' -and $_ -notlike 'Binary strings:*' }) { 'High' }
                    elseif ($sources | Where-Object { $_ -like 'Script:*' }) { 'Medium' }
                    else { 'Low' }

                    [pscustomobject]@{
                        FileName           = $fileName
                        Name               = $name
                        Scope              = $scope
                        Confidence         = $confidence
                        Hidden             = $entry.HiddenReasons.Count -gt 0
                        HiddenReasons      = [string[]]@($entry.HiddenReasons)
                        Secure             = $secureList.Contains($name)
                        HasDefault         = $entry.HasDefault
                        DefaultValue       = $entry.DefaultValue
                        PossibleValues     = @($entry.PossibleValues | Sort-Object Source, Order)
                        RuntimeAssignments = @($entry.RuntimeAssignments)
                        Controls           = [string[]]@($entry.Controls)
                        DefinedIn          = [string[]]$sources
                        Transforms         = [string[]]$appliedTransforms
                        FilePath           = $FullPath
                    }
                }
            }
            catch {
                Write-Error -Message "Failed to read MSI file '$File': $($_.Exception.GetBaseException().Message)" -TargetObject $File
            }
            finally {
                Close-MsiObject $Database
                $Database = $null
                if ($workDir -and (Test-Path -LiteralPath $workDir)) {
                    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    end {
        Close-MsiObject $Installer
    }
}
