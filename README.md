# FirewallForge

WPF GUI suite for Windows Firewall rule management. Two tools: a **live manager** for your system firewall and an **offline editor** for safe rule manipulation on backup files.

## Tools

### FirewallManager.ps1
Live firewall rule management:
- Browse all rules in a searchable DataGrid
- Inline editing (enabled, action, profile, protocol, ports)
- Backup rules to `.fwbackup` files (netsh export wrapped in JSON with metadata)
- Restore from backup
- Delete rules (multi-select)
- Export to CSV
- Filter by name, program, port, description

### FirewallRulesEditor.ps1
Offline backup editor:
- Import `.fwbackup` files created by FirewallManager
- Browse, edit, and delete rules without touching your live firewall
- Export modified rulesets back to `.fwbackup` or CSV
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
