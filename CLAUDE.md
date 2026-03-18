# CLAUDE.md - FirewallForge

## Overview
WPF GUI suite for Windows Firewall rule management. Two tools: live manager + offline backup editor. v1.0.

## Tech Stack
- PowerShell 5.1, WPF GUI
- Dark theme (#1E1E1E)
- Full ComboBox ControlTemplates for dark mode dropdowns

## Key Files
- `FirewallManager.ps1` (~926 lines) — Live firewall rule management
- `FirewallRulesEditor.ps1` (~976 lines) — Offline `.fwbackup` file editor

## Key Details
- **FirewallManager**: Browse/edit/delete live rules, backup (netsh export wrapped in JSON+base64), restore, CSV export, search/filter by name/program/port
- **FirewallRulesEditor**: Import `.fwbackup` files, edit rules offline without touching live firewall, export modified rulesets
- Shared dark theme and ComboBox ControlTemplate pattern

## Build/Run
```powershell
# Live management (requires Administrator)
.\FirewallManager.ps1

# Offline editing (no admin needed)
.\FirewallRulesEditor.ps1
```

## Version
1.0

## Gotchas
- `.fwbackup` format is netsh export wrapped in JSON with base64 payload + metadata
- FirewallRulesEditor explicitly does NOT modify system firewall
