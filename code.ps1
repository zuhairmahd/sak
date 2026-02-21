
# Create the VS Code policy registry key if it doesn't exist
if (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\VSCode"))
{
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\VSCode" -Force -EA SilentlyContinue
}

# Set AllowedExtensions policy - allows Microsoft extensions and approved third-party extensions
$allowedExtensions = @{
    # Allow all extensions from trusted publishers
    "microsoft"          = $true
    "github"             = $true
    # Version-locked extensions (specific versions only)
    "charliermarsh.ruff" = @("2025.24.0")
    "eeyore.yapf"        = @("2025.5.107163247")
}

# Convert to JSON string
$allowedExtensionsJson = $allowedExtensions | ConvertTo-Json -Compress
try
{
    # Set the registry value
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\VSCode" -Name "AllowedExtensions" -PropertyType String -Value $allowedExtensionsJson -Force

    # Set UpdateMode to none (prevents update prompts for non-admin users)
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\VSCode" -Name "UpdateMode" -PropertyType String -Value "none" -Force

    # Exit with success code
    exit 0
}
catch
{
    Write-Error "Failed to set VS Code policies: $_"
    # Exit with error code
    exit 1
}