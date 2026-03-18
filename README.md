<p align="center"><img src="icon.svg" width="128" height="128" alt="FirewallForge"></p>

# FirewallForge

WPF GUI suite for Windows Firewall rule management. Two tools: a **live manager** for your system firewall and an **offline editor** for safe rule manipulation on backup files.

## Tools

### FirewallManager.ps1
Live firewall rule management:
- Browse all rules in a searchable DataGrid with Group column
- Inline editing (enabled, action, profile, protocol, ports)
- Backup rules to `.fwbackup` files (netsh export wrapped in JSON with metadata)
- Restore from backup
- Delete rules (multi-select)
- Export to CSV
- Filter by name, program, port, description
- **Filter by Group** via dropdown
- **Find Duplicates** - scans rules for matching program + direction + action + ports, shows report and filters grid
- **Quick Block** menu - one-click blocking for programs (file browser), ports (text input), or IP addresses (text input)
- **Collapsible Stats Panel** - total/inbound/outbound, enabled/disabled, allow/block, by profile counts
- **Search Highlighting** - bold + colored rows when search filter is active

### FirewallRulesEditor.ps1
Offline backup editor:
- Import `.fwbackup` files created by FirewallManager
- Browse, edit, and delete rules without touching your live firewall
- Export modified rulesets back to `.fwbackup` or CSV
- **Merge Import** - import a second backup and merge new rules into the current set (deduplicates by name)
- **Show Changes** - diff view comparing current state to the original import (added, deleted, modified rules)
- **Rule count in title bar** - always shows current rule count
- Safe sandbox for planning firewall changes

## Usage

```powershell
# Live firewall management (requires Administrator)
.\FirewallManager.ps1

# Offline backup editing (no admin needed)
.\FirewallRulesEditor.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges (FirewallManager only)
