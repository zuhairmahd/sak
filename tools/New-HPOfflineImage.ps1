<#
.SYNOPSIS
    Builds a Windows install.wim for a specific HP platform: downloads that platform's drivers
    and UWP (hardware support) apps with HP CMSL, injects both offline, and writes a new WIM.

.DESCRIPTION
    Designed to run on a build machine that is NOT the target HP model, so the platform, OS and
    OS version are always passed explicitly to CMSL instead of being read from the local device.

    Flow:
      1. Preflight: CMSL commands present, image index readable, image build compared with
         -Os/-OsVer (warning only), architecture detected.
      2. Content download (before mounting, so the image is not held open during downloads):
           Drivers  - New-HPDriverPack -Format NoCompressedFile -RemoveOlder.
                      Fallback: HP's published driver pack SoftPaq (Get-HPSoftpaqList -Category Driverpack).
           UWP apps - New-HPUWPDriverPack -Format NoCompressedFile.
                      Fallback: every SoftPaq with the UWP characteristic, extracted.
         With -ContentSource Auto (default) the fallback runs when the build cmdlet is missing,
         throws, or produces nothing. HP documents New-HPUWPDriverPack as unsupported on non-HP
         devices and limited to G8+ platforms and 22H2+, which is why the fallback exists.
      3. Exports the chosen index to a staging WIM (also converts ESD to WIM and leaves the
         source untouched), mounts it.
      4. Add-WindowsDriver -Recurse on the driver folder.
      5. Add-AppxToOfflineImage.ps1 on the UWP folder (dependency and license resolution).
      6. Dismounts with -Save (or -Discard on a fatal error), then exports the staging image to
         -OutputImagePath with maximum compression, which also drops superseded data.

.PARAMETER Platform
    HP platform ID (4 hex characters, for example 8A05). On an HP device: (Get-CimInstance Win32_BaseBoard).Product.

.PARAMETER Os
    win10 or win11.

.PARAMETER OsVer
    Feature update, for example 23H2 or 24H2.

.PARAMETER ImagePath
    Source install.wim or install.esd. It is never modified.

.PARAMETER Index
    Index in the source image (Get-WindowsImage -ImagePath <file> lists them).

.PARAMETER OutputImagePath
    The finished WIM. Must not exist unless -Force is used, because Export-WindowsImage
    appends to an existing file instead of replacing it.

.PARAMETER WorkPath
    Root for downloads, staging WIM, mount folder and logs. Needs room for the image plus
    several GB of drivers. A per-platform subfolder is created.

.PARAMETER ContentSource
    Auto (default): build cmdlets first, SoftPaq fallback. Build: build cmdlets only.
    Softpaq: SoftPaq route only.

.PARAMETER UnselectList
    SoftPaq numbers or partial names to exclude; passed to New-HPDriverPack/New-HPUWPDriverPack.

.PARAMETER Regions
    Passed to Add-AppxProvisionedPackage. Default "all" so apps provision without a Start pin.

.PARAMETER AllowUnlicensed
    Provision UWP apps that have no license file with -SkipLicense.

.PARAMETER AppxScriptPath
    Path to Add-AppxToOfflineImage.ps1. Defaults to the same folder as this script.

.EXAMPLE
    .\New-HPOfflineImage.ps1 -Platform 8A05 -Os win11 -OsVer 24H2 `
        -ImagePath D:\Media\sources\install.wim -Index 3 `
        -OutputImagePath D:\Images\8A05_win11_24H2.wim -AllowUnlicensed

.EXAMPLE
    # Preview the SoftPaqs and image details only; nothing is downloaded or mounted.
    .\New-HPOfflineImage.ps1 -Platform 8A05 -Os win11 -OsVer 24H2 -ImagePath D:\install.wim -Index 3 `
        -OutputImagePath D:\out.wim -WhatIf

.NOTES
    Windows PowerShell 5.1, elevated, not PowerShell ISE (CMSL does not support ISE).
    References:
      https://developers.hp.com/hp-client-management/doc/new-hpdriverpack
      https://developers.hp.com/hp-client-management/doc/new-hpuwpdriverpack
      https://developers.hp.com/hp-client-management/doc/get-softpaqlist
      https://developers.hp.com/hp-client-management/doc/get-softpaq
      https://learn.microsoft.com/powershell/module/dism/add-appxprovisionedpackage
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules Dism

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9A-Fa-f]{4}$')]
    [string]$Platform,

    [Parameter(Mandatory)]
    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [Parameter(Mandatory)]
    [ValidatePattern('^\d{2}H\d$')]
    [string]$OsVer,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ImagePath,

    [Parameter(Mandatory)]
    [ValidateRange(1, 999)]
    [int]$Index,

    [Parameter(Mandatory)]
    [string]$OutputImagePath,

    [string]$WorkPath = 'C:\HPImageBuild',

    [ValidateSet('Auto', 'Build', 'Softpaq')]
    [string]$ContentSource = 'Auto',

    [string[]]$UnselectList,

    [string]$Regions = 'all',

    [switch]$AllowUnlicensed,
    [switch]$SkipDrivers,
    [switch]$SkipApps,
    [switch]$Force,
    [switch]$KeepWorkFiles,

    [string]$AppxScriptPath = (Join-Path $PSScriptRoot 'Add-AppxToOfflineImage.ps1')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

#region Helpers

function Resolve-FirstCommand {
    # CMSL 1.7+ uses HP-prefixed names; older releases (and aliases) use the unprefixed ones.
    param([string[]]$Names, [switch]$Optional)
    foreach ($n in $Names) {
        $c = Get-Command -Name $n -ErrorAction SilentlyContinue
        if ($c) { return $c }
    }
    if ($Optional) { return $null }
    throw "None of these HP CMSL commands were found: $($Names -join ', '). Is HP CMSL installed?"
}

function ConvertTo-ArchName {
    param($Value)
    switch -Regex ("$Value".Trim().ToLowerInvariant()) {
        '^(0|x86)$'       { return 'x86' }
        '^(5|arm)$'       { return 'arm' }
        '^(9|x64|amd64)$' { return 'x64' }
        '^(12|arm64)$'    { return 'arm64' }
    }
    return $null
}

function Reset-Folder {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Recurse -Force }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Get-FileCount {
    param([string]$Path, [string[]]$Extensions)
    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    @(Get-ChildItem -LiteralPath $Path -Recurse -File |
        Where-Object { $Extensions -contains $_.Extension.ToLowerInvariant() }).Count
}

function Get-SoftpaqNumber {
    param($Softpaq)
    return ("$($Softpaq.Id)" -replace '^(?i)sp', '')
}

function Expand-HPSoftpaqs {
    # Downloads and extracts each SoftPaq into <Destination>\sp<number>.
    param([object[]]$Softpaqs, [string]$Destination, [string]$DownloadPath)
    foreach ($sp in $Softpaqs) {
        $num = Get-SoftpaqNumber $sp
        $target = Join-Path $Destination "sp$num"
        Write-Output "  sp$num  $($sp.Name)  $($sp.Version)"
        & $script:cmdGetSoftpaq -Number $num -SaveAs (Join-Path $DownloadPath "sp$num.exe") `
            -Extract -DestinationPath $target -Overwrite yes -Quiet
    }
}

#endregion Helpers

#region Preflight

Import-Module HPCMSL -ErrorAction SilentlyContinue
$script:cmdSoftpaqList = Resolve-FirstCommand 'Get-HPSoftpaqList', 'Get-SoftpaqList'
$script:cmdGetSoftpaq  = Resolve-FirstCommand 'Get-HPSoftpaq', 'Get-Softpaq'
$cmdDriverPack    = Resolve-FirstCommand 'New-HPDriverPack' -Optional
$cmdUwpDriverPack = Resolve-FirstCommand 'New-HPUWPDriverPack' -Optional

if (-not $SkipApps -and -not (Test-Path -LiteralPath $AppxScriptPath -PathType Leaf)) {
    throw "Add-AppxToOfflineImage.ps1 not found at '$AppxScriptPath'. Use -AppxScriptPath."
}

$Platform = $Platform.ToUpperInvariant()
$OutputImagePath = [System.IO.Path]::GetFullPath($OutputImagePath)
if (Test-Path -LiteralPath $OutputImagePath) {
    if (-not $Force) { throw "'$OutputImagePath' already exists. Export-WindowsImage would append to it. Use -Force to replace it." }
}

# Model name is informational only; a failure here does not stop the build.
try {
    $cmdDetails = Resolve-FirstCommand 'Get-HPDeviceDetails' -Optional
    if ($cmdDetails) {
        $names = @(& $cmdDetails -Platform $Platform | ForEach-Object { $_.Name } | Select-Object -Unique)
        if ($names.Count) { Write-Output "Platform $Platform`: $($names -join '; ')" }
    }
} catch { Write-Verbose "Get-HPDeviceDetails: $($_.Exception.Message)" }

$img = Get-WindowsImage -ImagePath $ImagePath -Index $Index
$imageArch = ConvertTo-ArchName $img.Architecture
if ($imageArch -notin @('x64', 'arm64')) { throw "Image architecture '$($img.Architecture)' is not supported by HP driver packs." }
$bitness = $(if ($imageArch -eq 'arm64') { 'arm64' } else { '64' })
Write-Output "Source: $($img.ImageName), version $($img.Version), $imageArch"

# Compare the image build with -Os/-OsVer so a 23H2 driver set doesn't silently land in a 24H2 image.
$knownBuilds = @{
    19044 = 'win10 21H2'; 19045 = 'win10 22H2'
    22000 = 'win11 21H2'; 22621 = 'win11 22H2'; 22631 = 'win11 23H2'
    26100 = 'win11 24H2'; 26200 = 'win11 25H2'
}
$build = ([version]$img.Version).Build
if ($knownBuilds.ContainsKey($build) -and $knownBuilds[$build] -ne "$Os $OsVer") {
    Write-Warning "The image is build $build ($($knownBuilds[$build])) but you asked for $Os $OsVer. Content will be fetched for $Os $OsVer."
}

#endregion Preflight

#region Preview (-WhatIf)

if ($WhatIfPreference) {
    $common = @{ Platform = $Platform; Os = $Os; OSVer = $OsVer }
    if ($UnselectList) { $common.UnselectList = $UnselectList }

    if (-not $SkipDrivers) {
        Write-Output "`nDriver SoftPaqs:"
        if ($cmdDriverPack -and $ContentSource -ne 'Softpaq') { & $cmdDriverPack @common -WhatIf }
        else { & $script:cmdSoftpaqList -Platform $Platform -Os $Os -OsVer $OsVer -Bitness $bitness -Category Driverpack | Select-Object Id, Name, Version, ReleaseDate }
    }
    if (-not $SkipApps) {
        Write-Output "`nUWP SoftPaqs:"
        if ($cmdUwpDriverPack -and $ContentSource -ne 'Softpaq') { & $cmdUwpDriverPack @common -WhatIf }
        else { & $script:cmdSoftpaqList -Platform $Platform -Os $Os -OsVer $OsVer -Bitness $bitness -Characteristic UWP | Select-Object Id, Name, Version }
    }
    Write-Output "`nWould export index $Index to a staging WIM, inject the content, and write $OutputImagePath."
    return
}

#endregion Preview

#region Main

$root      = Join-Path $WorkPath "$($Platform)_$($Os)_$($OsVer)"
$driverDir = Join-Path $root 'Drivers'
$uwpDir    = Join-Path $root 'UWP'
$dlDir     = Join-Path $root 'Downloads'
$mountDir  = Join-Path $root 'Mount'
$logDir    = Join-Path $root ('Logs_{0:yyyyMMdd_HHmmss}' -f (Get-Date))
$staging   = Join-Path $root 'staging.wim'

New-Item -ItemType Directory -Path $root, $logDir -Force | Out-Null
Start-Transcript -LiteralPath (Join-Path $logDir 'transcript.log') | Out-Null

$mounted = $false
$fatal = $false
$appxResults = @()
$driverCount = 0

try {
    if ((Get-WindowsImage -Mounted | Where-Object { $_.Path -ieq $mountDir })) {
        throw "An image is already mounted at $mountDir. Dismount it first (Dismount-WindowsImage -Path '$mountDir' -Discard)."
    }
    Reset-Folder $dlDir

    $common = @{ Platform = $Platform; Os = $Os; OSVer = $OsVer; Format = 'NoCompressedFile'; Overwrite = $true }
    if ($UnselectList) { $common.UnselectList = $UnselectList }
    # -TempDownloadPath exists only in CMSL 1.7.2 and later.
    $supportsTemp = [bool]($cmdDriverPack -and $cmdDriverPack.Parameters.ContainsKey('TempDownloadPath'))
    if ($supportsTemp) { $common.TempDownloadPath = $dlDir }

    # --- 1. Drivers ------------------------------------------------------------------------
    if (-not $SkipDrivers) {
        Reset-Folder $driverDir
        $built = $false
        if ($ContentSource -ne 'Softpaq') {
            if (-not $cmdDriverPack) {
                if ($ContentSource -eq 'Build') { throw 'New-HPDriverPack is not available in this CMSL version.' }
            }
            else {
                try {
                    Write-Output 'Building driver pack with New-HPDriverPack...'
                    & $cmdDriverPack @common -Path $driverDir -RemoveOlder
                    $built = (Get-FileCount $driverDir '.inf') -gt 0
                    if (-not $built) { Write-Warning 'New-HPDriverPack produced no INF files.' }
                }
                catch {
                    if ($ContentSource -eq 'Build') { throw }
                    Write-Warning "New-HPDriverPack failed: $($_.Exception.Message)"
                }
            }
        }
        if (-not $built) {
            Write-Output 'Using HP published driver pack SoftPaq...'
            Reset-Folder $driverDir
            $dp = @(& $script:cmdSoftpaqList -Platform $Platform -Os $Os -OsVer $OsVer -Bitness $bitness -Category Driverpack |
                Sort-Object -Property @{ Expression = { "$($_.ReleaseDate)" }; Descending = $true }, @{ Expression = { "$($_.Version)" }; Descending = $true } |
                Select-Object -First 1)
            if (-not $dp.Count) { throw "HP lists no driver pack for $Platform $Os $OsVer." }
            Expand-HPSoftpaqs -Softpaqs $dp -Destination $driverDir -DownloadPath $dlDir
        }
        $infCount = Get-FileCount $driverDir '.inf'
        if (-not $infCount) { throw "No INF files found in $driverDir after download." }
        Write-Output "Driver content ready: $infCount INF file(s)."
    }

    # --- 2. UWP apps -----------------------------------------------------------------------
    if (-not $SkipApps) {
        Reset-Folder $uwpDir
        $appxExt = @('.appx', '.msix', '.appxbundle', '.msixbundle')
        $built = $false
        if ($ContentSource -ne 'Softpaq') {
            if (-not $cmdUwpDriverPack) {
                if ($ContentSource -eq 'Build') { throw 'New-HPUWPDriverPack is not available in this CMSL version (1.6.9 or later is required).' }
            }
            else {
                try {
                    Write-Output 'Building UWP pack with New-HPUWPDriverPack...'
                    & $cmdUwpDriverPack @common -Path $uwpDir
                    $built = (Get-FileCount $uwpDir $appxExt) -gt 0
                    if (-not $built) { Write-Warning 'New-HPUWPDriverPack produced no app packages.' }
                }
                catch {
                    if ($ContentSource -eq 'Build') { throw }
                    Write-Warning "New-HPUWPDriverPack failed: $($_.Exception.Message)"
                }
            }
        }
        if (-not $built) {
            Write-Output 'Downloading SoftPaqs with the UWP characteristic...'
            Reset-Folder $uwpDir
            $sps = @(& $script:cmdSoftpaqList -Platform $Platform -Os $Os -OsVer $OsVer -Bitness $bitness -Characteristic UWP)
            if ($UnselectList) {
                $sps = @($sps | Where-Object {
                    $sp = $_; -not ($UnselectList | Where-Object { "sp$(Get-SoftpaqNumber $sp)" -like "*$_*" -or $sp.Name -like "*$_*" })
                })
            }
            if ($sps.Count) { Expand-HPSoftpaqs -Softpaqs $sps -Destination $uwpDir -DownloadPath $dlDir }
        }
        $appCount = Get-FileCount $uwpDir $appxExt
        if (-not $appCount) { Write-Warning "No UWP app packages found for $Platform $Os $OsVer; skipping app provisioning." }
        else { Write-Output "UWP content ready: $appCount package file(s)." }
    }

    # --- 3. Stage and mount ----------------------------------------------------------------
    Write-Output "Exporting index $Index to staging WIM..."
    if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Force }
    Export-WindowsImage -SourceImagePath $ImagePath -SourceIndex $Index -DestinationImagePath $staging `
        -CompressionType max -LogPath (Join-Path $logDir 'dism_export_staging.log') | Out-Null

    Reset-Folder $mountDir
    Write-Output 'Mounting staging image...'
    Mount-WindowsImage -ImagePath $staging -Index 1 -Path $mountDir -LogPath (Join-Path $logDir 'dism_mount.log') | Out-Null
    $mounted = $true

    # --- 4. Inject drivers -----------------------------------------------------------------
    if (-not $SkipDrivers) {
        Write-Output 'Adding drivers...'
        $drv = @(Add-WindowsDriver -Path $mountDir -Driver $driverDir -Recurse -LogPath (Join-Path $logDir 'dism_drivers.log'))
        $driverCount = $drv.Count
        Write-Output "Added $driverCount driver(s)."
    }

    # --- 5. Provision UWP apps -------------------------------------------------------------
    if (-not $SkipApps -and (Get-FileCount $uwpDir $appxExt)) {
        Write-Output 'Provisioning UWP apps...'
        $appxArgs = @{
            MountPath    = $mountDir
            SourcePath   = $uwpDir
            Architecture = $imageArch
            LogDirectory = (Join-Path $logDir 'Appx')
            Confirm      = $false
        }
        if ($Regions) { $appxArgs.Regions = $Regions }
        if ($AllowUnlicensed) { $appxArgs.AllowUnlicensed = $true }
        $out = & $AppxScriptPath @appxArgs
        $appxResults = @($out | Where-Object { $_ -is [psobject] -and $_.PSObject.Properties['Status'] })
        $out | Where-Object { $_ -is [string] } | ForEach-Object { Write-Output $_ }
    }
}
catch {
    $fatal = $true
    Write-Error "Fatal: $($_.Exception.Message)" -ErrorAction Continue
}
finally {
    if ($mounted) {
        if ($fatal) {
            Write-Output 'Discarding changes and dismounting.'
            Dismount-WindowsImage -Path $mountDir -Discard -LogPath (Join-Path $logDir 'dism_dismount.log') | Out-Null
        }
        else {
            Write-Output 'Committing changes and dismounting.'
            try {
                Dismount-WindowsImage -Path $mountDir -Save -LogPath (Join-Path $logDir 'dism_dismount.log') | Out-Null
            }
            catch {
                $fatal = $true
                Write-Error "Dismount failed: $($_.Exception.Message). Run Dismount-WindowsImage -Path '$mountDir' -Discard, then Clear-WindowsCorruptMountPoint if needed." -ErrorAction Continue
            }
        }
    }
}

# --- 6. Final export ------------------------------------------------------------------------
if (-not $fatal) {
    try {
        Write-Output "Writing $OutputImagePath..."
        if (Test-Path -LiteralPath $OutputImagePath) { Remove-Item -LiteralPath $OutputImagePath -Force }
        New-Item -ItemType Directory -Path (Split-Path $OutputImagePath -Parent) -Force | Out-Null
        Export-WindowsImage -SourceImagePath $staging -SourceIndex 1 -DestinationImagePath $OutputImagePath `
            -CompressionType max -CheckIntegrity -LogPath (Join-Path $logDir 'dism_export_final.log') | Out-Null
    }
    catch {
        $fatal = $true
        Write-Error "Final export failed: $($_.Exception.Message)" -ErrorAction Continue
    }
}

if (-not $KeepWorkFiles) {
    foreach ($p in @($staging, $dlDir, $mountDir)) {
        if ($p -eq $mountDir -and $mounted -and $fatal) { continue }
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

$appxSummary = ($appxResults | Group-Object Status | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', '
if (-not $appxSummary) { $appxSummary = 'none' }
Write-Output ''
Write-Output "Result: $(if ($fatal) { 'FAILED' } else { 'Succeeded' })"
Write-Output "Drivers added: $driverCount"
Write-Output "UWP apps: $appxSummary"
if (-not $fatal) { Write-Output "Image: $OutputImagePath" }
Write-Output "Logs: $logDir"
Stop-Transcript | Out-Null
$appxResults

#endregion Main
