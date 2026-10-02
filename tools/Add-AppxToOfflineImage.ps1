<#
.SYNOPSIS
    Provisions every APPX/MSIX package found under a folder into an offline Windows image,
    resolving framework dependencies and offline license files automatically. Optionally
    injects INF drivers first.

.DESCRIPTION
    The script:
      1. Mounts a WIM index (or uses an image you already mounted).
      2. Optionally adds drivers from -DriverPath with Add-WindowsDriver -Recurse.
      3. Recursively scans -SourcePath for .appx, .msix, .appxbundle and .msixbundle files.
      4. Reads each package's manifest (AppxManifest.xml, or AppxBundleManifest.xml plus the
         manifests of the inner application packages for bundles) to classify it as a
         main app, a framework (dependency) or a resource package, and to collect its
         PackageDependency entries (Name + MinVersion).
      5. For each main app, resolves dependencies against the framework packages it found:
         matching Name, Version >= MinVersion, and an architecture the image can run.
         The highest qualifying version per architecture is passed to -DependencyPackagePath,
         because DISM requires every architecture-specific dependency the target image needs.
      6. Finds the app's offline license by reading every *.xml whose root element is
         <License>: first by the <PFN> (package family name) inside the license, then by a
         single unbound license in the same folder, then by file name.
      7. Calls Add-AppxProvisionedPackage with -PackagePath, -DependencyPackagePath and
         -LicensePath (or -SkipLicense only when -AllowUnlicensed is set).
      8. Commits the image (or discards it on a fatal error, or when -NoCommit is used).

    Results are written to the pipeline and to Results.csv in the log directory, along with
    one DISM log per operation.

.PARAMETER ImagePath
    Path to the .wim file to mount. The file must not be read-only. ESD files cannot be mounted;
    export the index to a WIM first.

.PARAMETER Index
    Image index inside the WIM.

.PARAMETER MountPath
    Mount directory. In WIM mode it must be empty or not exist. In mounted mode (no -ImagePath)
    it is the root of an image you already mounted; the script will not dismount it.

.PARAMETER SourcePath
    Folder to search recursively for packages and license files.

.PARAMETER DriverPath
    Optional folder of INF drivers to add (recursively) before the apps.

.PARAMETER Architecture
    Override the image architecture (x86, x64, arm, arm64). Normally detected automatically.

.PARAMETER Regions
    Passed to -Regions on Add-AppxProvisionedPackage, for example "all". Without it, DISM only
    provisions the app if it is pinned to the Start layout (client OS behavior).

.PARAMETER AllowUnlicensed
    Provision apps with no matching license using -SkipLicense. Microsoft documents that this
    should only be used for apps that do not require a license on Enterprise or Server editions.

.PARAMETER RequireAllDependencies
    Skip an app when any declared dependency is not found under -SourcePath. By default the
    script warns and still attempts provisioning, because the framework may already be in the image.

.PARAMETER Force
    Provision even if the image already has the same or a newer version of the app.

.PARAMETER NoCommit
    Discard all changes when dismounting (dry run against a real mount).

.PARAMETER LogDirectory
    Where DISM logs and Results.csv are written. Defaults to a timestamped folder in %TEMP%.

.EXAMPLE
    .\Add-AppxToOfflineImage.ps1 -ImagePath D:\Images\install.wim -Index 3 -MountPath D:\Mount `
        -SourcePath D:\Apps -DriverPath D:\Drivers\Model123 -Regions all

.EXAMPLE
    # Preview what would be provisioned, without mounting or changing anything.
    .\Add-AppxToOfflineImage.ps1 -ImagePath D:\Images\install.wim -Index 3 -MountPath D:\Mount `
        -SourcePath D:\Apps -WhatIf

.EXAMPLE
    # Work against an image you already mounted with Mount-WindowsImage.
    .\Add-AppxToOfflineImage.ps1 -MountPath D:\Mount -SourcePath D:\Apps -AllowUnlicensed

.NOTES
    Run in Windows PowerShell 5.1, elevated.
    Reference: https://learn.microsoft.com/powershell/module/dism/add-appxprovisionedpackage
               https://learn.microsoft.com/windows-hardware/manufacture/desktop/sideload-apps-with-dism-s14
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules Dism

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Mounted')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Wim')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ImagePath,

    [Parameter(Mandatory, ParameterSetName = 'Wim')]
    [ValidateRange(1, 999)]
    [int]$Index,

    [Parameter(Mandatory, ParameterSetName = 'Wim')]
    [Parameter(Mandatory, ParameterSetName = 'Mounted')]
    [string]$MountPath,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$SourcePath,

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DriverPath,

    [ValidateSet('x86', 'x64', 'arm', 'arm64')]
    [string]$Architecture,

    [string]$Regions,

    [switch]$AllowUnlicensed,
    [switch]$RequireAllDependencies,
    [switch]$Force,
    [switch]$NoCommit,

    [string]$LogDirectory = (Join-Path $env:TEMP ('AppxProvision_{0:yyyyMMdd_HHmmss}' -f (Get-Date)))
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

#region Helpers

function ConvertTo-ArchName {
    param($Value)
    switch -Regex ("$Value".Trim().ToLowerInvariant()) {
        '^(0|x86)$'         { return 'x86' }
        '^(5|arm)$'         { return 'arm' }
        '^(9|x64|amd64)$'   { return 'x64' }
        '^(12|arm64)$'      { return 'arm64' }
        '^(11|neutral|)$'   { return 'neutral' }
    }
    return $null
}

function Get-CompatibleArchitectures {
    param([string]$ImageArch)
    # Ordered by preference: native first.
    switch ($ImageArch) {
        'x64'   { return @('x64', 'x86', 'neutral') }
        'x86'   { return @('x86', 'neutral') }
        'arm64' { return @('arm64', 'arm', 'x64', 'x86', 'neutral') }
        'arm'   { return @('arm', 'neutral') }
    }
    throw "Unsupported image architecture '$ImageArch'."
}

function ConvertTo-Version {
    param([string]$Text)
    $v = $null
    if ([version]::TryParse($Text, [ref]$v)) { return $v }
    return [version]'0.0.0.0'
}

function Get-ZipEntry {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $wanted = $EntryName.Replace('\', '/')
    foreach ($e in $Zip.Entries) {
        if ($e.FullName -ieq $wanted -or [uri]::UnescapeDataString($e.FullName) -ieq $wanted) { return $e }
    }
    return $null
}

function Read-ZipEntryXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = Get-ZipEntry -Zip $Zip -EntryName $EntryName
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ManifestDependencies {
    param([xml]$Manifest)
    $nodes = $Manifest.SelectNodes("//*[local-name()='Dependencies']/*[local-name()='PackageDependency']")
    foreach ($d in $nodes) {
        [pscustomobject]@{
            Name       = $d.GetAttribute('Name')
            MinVersion = ConvertTo-Version $d.GetAttribute('MinVersion')
        }
    }
}

function Get-AppxFileInfo {
    param([System.IO.FileInfo]$File, [string[]]$CompatArch)

    $isBundle = $File.Extension -match 'bundle$'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)
    try {
        if ($isBundle) {
            $xml = Read-ZipEntryXml -Zip $zip -EntryName 'AppxMetadata/AppxBundleManifest.xml'
            if (-not $xml) { throw 'AppxMetadata/AppxBundleManifest.xml not found.' }

            $id = $xml.SelectSingleNode("/*[local-name()='Bundle']/*[local-name()='Identity']")
            $appPkgs = @($xml.SelectNodes("//*[local-name()='Packages']/*[local-name()='Package']") |
                Where-Object { $_.GetAttribute('Type') -eq 'application' })

            $arches = @($appPkgs | ForEach-Object { ConvertTo-ArchName $_.GetAttribute('Architecture') } | Select-Object -Unique)
            $deps = New-Object System.Collections.Generic.List[object]

            # Dependencies live in the inner application packages, not the bundle manifest.
            foreach ($p in $appPkgs) {
                $arch = ConvertTo-ArchName $p.GetAttribute('Architecture')
                if ($CompatArch -notcontains $arch) { continue }

                $innerEntry = Get-ZipEntry -Zip $zip -EntryName $p.GetAttribute('FileName')
                if (-not $innerEntry) {
                    Write-Warning "$($File.Name): inner package '$($p.GetAttribute('FileName'))' not found."
                    continue
                }
                $tmp = Join-Path $env:TEMP ([guid]::NewGuid().ToString() + '.appx')
                try {
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($innerEntry, $tmp, $true)
                    $innerZip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
                    try {
                        $m = Read-ZipEntryXml -Zip $innerZip -EntryName 'AppxManifest.xml'
                        if ($m) { foreach ($d in @(Get-ManifestDependencies $m)) { $deps.Add($d) } }
                    }
                    finally { $innerZip.Dispose() }
                }
                finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue -WhatIf:$false }
            }

            return [pscustomobject]@{
                Path          = $File.FullName
                FileName      = $File.Name
                Directory     = $File.DirectoryName
                IsBundle      = $true
                Name          = $id.GetAttribute('Name')
                Publisher     = $id.GetAttribute('Publisher')
                Version       = ConvertTo-Version $id.GetAttribute('Version')
                Architectures = [string[]]$arches
                IsFramework   = $false
                IsResource    = $false
                Dependencies  = $deps.ToArray()
            }
        }
        else {
            $xml = Read-ZipEntryXml -Zip $zip -EntryName 'AppxManifest.xml'
            if (-not $xml) { throw 'AppxManifest.xml not found.' }

            $id  = $xml.SelectSingleNode("/*[local-name()='Package']/*[local-name()='Identity']")
            $fw  = $xml.SelectSingleNode("/*[local-name()='Package']/*[local-name()='Properties']/*[local-name()='Framework']")
            $res = $xml.SelectSingleNode("/*[local-name()='Package']/*[local-name()='Properties']/*[local-name()='ResourcePackage']")

            return [pscustomobject]@{
                Path          = $File.FullName
                FileName      = $File.Name
                Directory     = $File.DirectoryName
                IsBundle      = $false
                Name          = $id.GetAttribute('Name')
                Publisher     = $id.GetAttribute('Publisher')
                Version       = ConvertTo-Version $id.GetAttribute('Version')
                Architectures = [string[]]@(ConvertTo-ArchName $id.GetAttribute('ProcessorArchitecture'))
                IsFramework   = [bool]($fw -and $fw.InnerText.Trim() -eq 'true')
                IsResource    = [bool]($res -and $res.InnerText.Trim() -eq 'true')
                Dependencies  = @(Get-ManifestDependencies $xml)
            }
        }
    }
    finally { $zip.Dispose() }
}

function Get-LicenseCatalog {
    param([string]$Root)
    foreach ($x in Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.xml') {
        if ($x.Name -ieq 'AppxManifest.xml') { continue }
        try { [xml]$lx = Get-Content -LiteralPath $x.FullName -Raw } catch { continue }
        if (-not $lx.DocumentElement -or $lx.DocumentElement.LocalName -ne 'License') { continue }
        $pfn = $lx.SelectSingleNode("//*[local-name()='PFN']")
        [pscustomobject]@{
            Path      = $x.FullName
            FileName  = $x.Name
            Directory = $x.DirectoryName
            PFN       = $(if ($pfn) { $pfn.InnerText.Trim() } else { $null })
        }
    }
}

function Find-License {
    param($Package, $Licenses)
    if (-not $Licenses) { return $null }

    # 1. License bound to this package family (PFN = <Name>_<PublisherId>).
    $byPfn = @($Licenses | Where-Object { $_.PFN -and $_.PFN -like "$($Package.Name)_*" })
    if ($byPfn.Count) {
        $same = @($byPfn | Where-Object { $_.Directory -eq $Package.Directory })
        if ($same.Count) { return $same[0] }
        return $byPfn[0]
    }

    # Licenses that carry a PFN for a different app are never used as a fallback.
    $unbound = @($Licenses | Where-Object { -not $_.PFN })

    # 2. Exactly one unbound license in the package's own folder.
    $sameDir = @($unbound | Where-Object { $_.Directory -eq $Package.Directory })
    if ($sameDir.Count -eq 1) { return $sameDir[0] }

    # 3. File name contains the package name or the package file's base name.
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Package.FileName)
    $byName = @($unbound | Where-Object { $_.FileName -like "*$($Package.Name)*" -or $_.FileName -like "$base*" })
    if ($byName.Count) { return $byName[0] }

    return $null
}

function Resolve-PackageDependencies {
    param($App, $Frameworks, [string[]]$CompatArch)

    $paths   = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]

    $required = @($App.Dependencies | Group-Object -Property Name | ForEach-Object {
        [pscustomobject]@{
            Name       = $_.Name
            MinVersion = ($_.Group | Sort-Object -Property MinVersion -Descending | Select-Object -First 1).MinVersion
        }
    })

    foreach ($r in $required) {
        $cands = @($Frameworks | Where-Object {
            $_.Name -eq $r.Name -and $_.Version -ge $r.MinVersion -and $CompatArch -contains $_.Architectures[0]
        })
        if (-not $cands.Count) { $missing.Add("$($r.Name) >= $($r.MinVersion)"); continue }

        # Highest qualifying version for each architecture the image can use.
        foreach ($g in ($cands | Group-Object -Property { $_.Architectures[0] })) {
            $paths.Add(($g.Group | Sort-Object -Property Version -Descending | Select-Object -First 1).Path)
        }
    }

    [pscustomobject]@{
        Paths   = [string[]]@($paths | Select-Object -Unique)
        Missing = [string[]]$missing.ToArray()
    }
}

#endregion Helpers

#region Main

New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false | Out-Null
$MountPath = [System.IO.Path]::GetFullPath($MountPath)
$results = New-Object System.Collections.Generic.List[object]
$mountedHere = $false
$fatal = $false

try {
    # --- Image and architecture -------------------------------------------------------
    $imageArch = $Architecture
    if ($PSCmdlet.ParameterSetName -eq 'Wim') {
        $img = Get-WindowsImage -ImagePath $ImagePath -Index $Index
        if (-not $imageArch) { $imageArch = ConvertTo-ArchName $img.Architecture }
        Write-Verbose "Image: $($img.ImageName), architecture $imageArch"

        if (Test-Path -LiteralPath $MountPath) {
            if (Get-ChildItem -LiteralPath $MountPath -Force | Select-Object -First 1) {
                throw "Mount directory '$MountPath' is not empty."
            }
        }
        elseif ($PSCmdlet.ShouldProcess($MountPath, 'Create mount directory')) {
            New-Item -ItemType Directory -Path $MountPath | Out-Null
        }

        if ($PSCmdlet.ShouldProcess("$ImagePath index $Index", "Mount to $MountPath")) {
            Mount-WindowsImage -ImagePath $ImagePath -Index $Index -Path $MountPath `
                -LogPath (Join-Path $LogDirectory 'dism_mount.log') | Out-Null
            $mountedHere = $true
        }
    }
    else {
        if (-not (Test-Path -LiteralPath (Join-Path $MountPath 'Windows'))) {
            throw "'$MountPath' does not look like the root of a mounted Windows image."
        }
        if (-not $imageArch) {
            $m = Get-WindowsImage -Mounted | Where-Object {
                [System.IO.Path]::GetFullPath($_.Path).TrimEnd('\') -ieq $MountPath.TrimEnd('\')
            } | Select-Object -First 1
            if ($m) {
                $imageArch = ConvertTo-ArchName (Get-WindowsImage -ImagePath $m.ImagePath -Index $m.ImageIndex).Architecture
            }
        }
    }
    if (-not $imageArch -or $imageArch -eq 'neutral') {
        throw 'Could not determine the image architecture. Specify -Architecture.'
    }
    $compat = Get-CompatibleArchitectures $imageArch
    $imageIsMounted = Test-Path -LiteralPath (Join-Path $MountPath 'Windows')

    # --- Drivers ----------------------------------------------------------------------
    if ($DriverPath) {
        if ($PSCmdlet.ShouldProcess($MountPath, "Add drivers from $DriverPath (recursive)")) {
            $drv = @(Add-WindowsDriver -Path $MountPath -Driver $DriverPath -Recurse `
                -LogPath (Join-Path $LogDirectory 'dism_drivers.log'))
            Write-Output "Added $($drv.Count) driver(s) from $DriverPath."
        }
    }

    # --- Package and license discovery ------------------------------------------------
    $extensions = @('.appx', '.msix', '.appxbundle', '.msixbundle')
    $files = @(Get-ChildItem -LiteralPath $SourcePath -Recurse -File | Where-Object {
        $extensions -contains $_.Extension.ToLowerInvariant() -and
        -not $_.FullName.StartsWith($MountPath + '\', [System.StringComparison]::OrdinalIgnoreCase)
    })
    Write-Verbose "Found $($files.Count) package file(s) under $SourcePath."

    $catalog = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        try { $catalog.Add((Get-AppxFileInfo -File $f -CompatArch $compat)) }
        catch {
            Write-Warning "Could not read '$($f.FullName)': $($_.Exception.Message)"
            $results.Add([pscustomobject]@{
                Name = $f.Name; Version = ''; Status = 'Unreadable'; License = ''
                Dependencies = ''; Detail = $_.Exception.Message
            })
        }
    }

    $licenses   = @(Get-LicenseCatalog -Root $SourcePath)
    $frameworks = @($catalog | Where-Object { $_.IsFramework })

    # Main apps: not framework, not resource, runnable on the image. One per Name:
    # highest version, then most native architecture.
    $apps = @($catalog |
        Where-Object { -not $_.IsFramework -and -not $_.IsResource } |
        Where-Object { $a = $_.Architectures; @($compat | Where-Object { $a -contains $_ }).Count -gt 0 } |
        Group-Object -Property Name |
        ForEach-Object {
            $_.Group | Sort-Object -Property `
                @{ Expression = { $_.Version }; Descending = $true },
                @{ Expression = { $a = $_.Architectures; ($compat | ForEach-Object { $a -contains $_ }).IndexOf($true) } } |
            Select-Object -First 1
        })

    Write-Output ("Image architecture {0}. {1} app(s), {2} framework package(s), {3} license file(s) found." -f `
        $imageArch, $apps.Count, $frameworks.Count, $licenses.Count)

    # --- Already provisioned ----------------------------------------------------------
    $existing = @{}
    if ($imageIsMounted) {
        foreach ($p in Get-AppxProvisionedPackage -Path $MountPath) {
            $existing[$p.DisplayName] = ConvertTo-Version $p.Version
        }
    }

    # --- Provision --------------------------------------------------------------------
    foreach ($app in $apps) {
        $deps = Resolve-PackageDependencies -App $app -Frameworks $frameworks -CompatArch $compat
        $lic  = Find-License -Package $app -Licenses $licenses
        $row  = [pscustomobject]@{
            Name         = $app.Name
            Version      = "$($app.Version)"
            Status       = ''
            License      = $(if ($lic) { $lic.FileName } else { '(none)' })
            Dependencies = (($deps.Paths | ForEach-Object { Split-Path $_ -Leaf }) -join '; ')
            Detail       = ''
        }

        if ($deps.Missing.Count) {
            $msg = 'Missing dependency: ' + ($deps.Missing -join ', ')
            $row.Detail = $msg
            if ($RequireAllDependencies) { $row.Status = 'Skipped'; $results.Add($row); Write-Warning "$($app.Name): $msg"; continue }
            Write-Warning "$($app.Name): $msg. Attempting anyway in case the image already contains it."
        }

        if (-not $Force -and $existing.ContainsKey($app.Name) -and $existing[$app.Name] -ge $app.Version) {
            $row.Status = 'AlreadyProvisioned'
            $row.Detail = "Image has version $($existing[$app.Name])"
            $results.Add($row); continue
        }

        if (-not $lic -and -not $AllowUnlicensed) {
            $row.Status = 'Skipped'
            $row.Detail = (($row.Detail, 'No license found; use -AllowUnlicensed to provision with -SkipLicense') | Where-Object { $_ }) -join ' | '
            Write-Warning "$($app.Name): no license found. Skipped."
            $results.Add($row); continue
        }

        $params = @{
            Path        = $MountPath
            PackagePath = $app.Path
            LogPath     = (Join-Path $LogDirectory ("dism_{0}.log" -f $app.Name))
        }
        if ($deps.Paths.Count) { $params.DependencyPackagePath = $deps.Paths }
        if ($lic) { $params.LicensePath = $lic.Path } else { $params.SkipLicense = $true }
        if ($Regions) { $params.Regions = $Regions }

        if ($PSCmdlet.ShouldProcess($app.Name, "Provision $($app.FileName) with $($deps.Paths.Count) dependency package(s), license: $($row.License)")) {
            try {
                Add-AppxProvisionedPackage @params | Out-Null
                $row.Status = $(if ($lic) { 'Provisioned' } else { 'ProvisionedUnlicensed' })
            }
            catch {
                $row.Status = 'Failed'
                $row.Detail = (($row.Detail, $_.Exception.Message) | Where-Object { $_ }) -join ' | '
                Write-Warning "$($app.Name): $($_.Exception.Message)"
            }
        }
        else { $row.Status = 'WhatIf' }
        $results.Add($row)
    }
}
catch {
    $fatal = $true
    Write-Error "Fatal: $($_.Exception.Message)" -ErrorAction Continue
}
finally {
    if ($mountedHere) {
        $dismountLog = Join-Path $LogDirectory 'dism_dismount.log'
        if ($fatal -or $NoCommit) {
            Write-Output 'Discarding changes and dismounting.'
            Dismount-WindowsImage -Path $MountPath -Discard -LogPath $dismountLog | Out-Null
        }
        else {
            Write-Output 'Committing changes and dismounting.'
            Dismount-WindowsImage -Path $MountPath -Save -LogPath $dismountLog | Out-Null
        }
    }

    $csv = Join-Path $LogDirectory 'Results.csv'
    $results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8 -WhatIf:$false
    $results

    $summary = $results | Group-Object -Property Status | ForEach-Object { "$($_.Name): $($_.Count)" }
    Write-Output ("Summary. " + ($summary -join '. ') + ". Logs and Results.csv are in $LogDirectory")
}

#endregion Main
