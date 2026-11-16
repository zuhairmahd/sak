---
name: AGENTS
applyTo: "**/*.ps1"
# This file tells Copilot to load the following content when working with any PowerShell script in the project.
# It provides a high‑level overview, useful links, and tool‑usage hints for this PowerShell toolkit.
---

# AGENTS for PowerShell Toolkit (SAK)

## Purpose
This instruction set gives Copilot quick context for the **SAK** PowerShell toolkit, a collection of helper scripts for Windows system administration. The instructions include:
1.  A short project description.
2.  Links to key documentation (`README.md`, `LICENSE`, `CONTRIBUTING.md`).
3.  Tool‑usage hints for the most common patterns used by the toolkit.
4.  A suggestion for a minimal `preToolUse` hook that ensures scripts run with strict mode enabled.

## Project Overview
- **Main entry point**: `sak.ps1` – provides an interactive menu or CLI options.
- **Utility functions**: `functions/*.ps1` – core helpers (registry, email, uninstall, service, network, ZIP, etc.).
- **Tools folder**: `tools/*.ps1` – higher‑level helpers (e.g., Autopilot diagnostics, graph API, offline image building).
- **Test harness**: `test.ps1` – runs simple syntax checks and demonstrates example usages.
- **Documentation**: `README.md`, `LICENSE`, `CONTRIBUTING.md`.
- **No unit‑test framework** – scripts include basic syntax validation.

## Useful Links
- [README](../README.md) – project overview and usage.
- [LICENSE](../LICENSE) – MIT license.
- [CONTRIBUTING](../CONTRIBUTING.md) – contribution guidelines.
- [functions](../functions/) – core helper scripts.
- [tools](../tools/) – advanced helpers.
- [test.ps1](../test.ps1) – quick sanity checks.

## Tool‑Usage Hints
- When opening a `.ps1` file, Copilot can suggest parameter sets such as `-CheckRegKeyExists`, `-GetUninstallCommands`, `-KillGuiltyProcesses` based on the script’s public functions.
- The `sak.ps1` script imports `functions/*.ps1` automatically; running it from the PowerShell terminal shows an interactive menu.
- For email extraction, use `Get-EmailAddresses` from a text blob; Copilot can suggest this helper when you see an address pattern.
- The `Invoke-ExternalProcess` helper demonstrates a pattern for launching external commands with timeout and retry logic.

## Optional Hook (preToolUse)
```json
{
  "name": "EnsureStrictMode",
  "type": "preToolUse",
  "description": "Activate PowerShell strict mode for every script in the project.",
  "command": "Set-StrictMode -Version Latest"
}
```
Add this JSON to `.github/hooks/EnsureStrictMode.json` if you want every script to run with `Set-StrictMode` enabled by default.

---
