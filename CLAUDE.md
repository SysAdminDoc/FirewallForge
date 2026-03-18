# CLAUDE.md - FirewallForge

## Overview
WPF GUI suite for Windows Firewall rule management. Two tools: live manager + offline backup editor. v1.1.0.

## Tech Stack
- PowerShell 5.1, WPF GUI
- Dark theme (#1E1E1E)
- Full ComboBox ControlTemplates for dark mode dropdowns

## Key Files
- `FirewallManager.ps1` (~1050 lines) - Live firewall rule management
- `FirewallRulesEditor.ps1` (~870 lines) - Offline `.fwbackup` file editor

## Key Details
- **FirewallManager**: Browse/edit/delete live rules, backup (netsh export wrapped in JSON+base64), restore, CSV export, search/filter by name/program/port, group filter, duplicate detection, quick-block (program/port/IP), collapsible stats panel, search highlighting
- **FirewallRulesEditor**: Import `.fwbackup` files, edit rules offline without touching live firewall, export modified rulesets, merge second backup into current set, diff view (show changes since import), rule count in title bar
- Shared dark theme and ComboBox ControlTemplate pattern

## Build/Run
```powershell
# Live management (requires Administrator)
.\FirewallManager.ps1

# Offline editing (no admin needed)
.\FirewallRulesEditor.ps1
```

## Version History
- **1.1.0** - Duplicate detection, stats panel, quick-block menu (program/port/IP), group filter, search highlighting, diff view, merge import, rule count in title bar
- **1.0** - Initial release. Backup/restore/edit/delete/search/CSV export

## Gotchas
- `.fwbackup` format is netsh export wrapped in JSON with base64 payload + metadata
- FirewallRulesEditor explicitly does NOT modify system firewall
- Diff tracking snapshots at import time; adding rules manually after import counts as "added" in diff
- Group column in FirewallManager pulled from rule metadata (may be empty for custom rules)
