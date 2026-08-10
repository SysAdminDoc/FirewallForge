# Changelog

All notable changes to FirewallForge will be documented in this file.

## [v1.3.0] - 2026-08-09

- Added: Per-program rule wizard with existing-rule inspection and Block / Allow In / Allow Out presets.
- Added: Optional connection monitor for Security events 5156/5157 with Allow, Block, and Ignore decisions.
- Added: Reversible outbound lockdown profile with rollback snapshots and essential service exceptions.
- Added: Approximate rule-priority view with flow predicates and optional endpoint reachability testing.
- Added: Group operations for enabling, disabling, or deleting rules by group with GPO protection.
- Added: Tailable pfirewall.log viewer with ALLOW/DROP coloring and action, direction, and port filters.
- Added: Offline comparison of any two .fwbackup files with added, removed, and modified reports.
- Added: Merge conflict strategies for prefer-newer, prefer-imported, and per-rule manual decisions.
- Added: Offline rule template library with JSON save, browse, insert, and delete actions.
- Added: Policy export to netsh, PowerShell, and binary GPO Registry.pol formats.
- Added: Task Scheduler integration for daily/weekly backups with retention rotation and headless worker.
- Added: IPv6 coverage audit for IPv4-only rules without matching IPv6 twins.
- Added: Regex search mode to the offline editor's rule filter with invalid-pattern feedback.
- Added: Named manager views for filters, regex state, column layout, and sort descriptions.
- Added: Bulk tag replace, append, and remove operations with GPO protection.
- Added: Rule health score audit for broad rules, duplicates, and missing executable paths.
- Added: Headless FirewallForge.ps1 deployment wrapper with profile validation, idempotent apply, and managed-group pruning.
- Added: DNS correlation report linking recent blocked connections to Windows DNS Client event 3008 queries and resolved domains.
- Added: Multi-machine backup comparison with consistent, drifted, and missing rule classification across endpoints.

## [v1.1.0]

- Added: Add project icon to README
- v1.1.0 - Duplicate detection, stats panel, quick-block, diff view
- Initial commit - FirewallForge

## Roadmap archive — 2026-08-10 — ROADMAP.md

<details>
<summary>Original roadmap snapshot</summary>

```markdown
# ROADMAP

Backlog for FirewallForge. Two-tool suite (live manager + offline editor) for Windows Firewall
rules. Goal: close the gap with Windows Firewall Control (WFC) without adopting its paid model.

## Planned Features

## Competitive Research

- **Windows Firewall Control (Malwarebytes, paid)** — the feature leader: interactive mode,
  notifications, one-click outbound-block preset. FirewallForge should match the free feature set
  without the nag.
- **simplewall (henrypp/simplewall)** — open-source, uses WFP directly (not WF rule API). Very
  lightweight, outbound-block-by-default. Good reference for a "minimal mode".
- **Portmaster (Safing)** — modern UX, per-connection prompts, SPN integration. The new-connection
  prompt UX is worth borrowing conceptually.
- **TinyWall** — free, lockdown mode, zone-aware. Closest free competitor; match its allowlist
  simplicity.

## Nice-to-Haves

## Open-Source Research (Round 2)

### Related OSS Projects
- **metablaster/WindowsFirewallRuleset** — https://github.com/metablaster/WindowsFirewallRuleset — Full ruleset framework; auto-path detection, `Deploy-Firewall` command, experimental remote deployment via PS Remoting.
- **SteveUnderScoreN/WindowsFirewall** — https://github.com/SteveUnderScoreN/WindowsFirewall — Enterprise GPO-driven firewall hardening; domain/tier baselines; audit event ID 5156/5157 wiring.
- **MScholtes/Firewall-Manager** — https://github.com/MScholtes/Firewall-Manager — Lean PS module: `Export-FirewallRules`, `Import-FirewallRules`, `Remove-FirewallRules` with CSV/JSON + filter flags.
- **Windows Firewall Notifier** — https://github.com/wokhan/WFN — WPF GUI for outbound-connection prompts, live connections map, bandwidth monitor.
- **Z3R0th-13/FirewallRules** — https://github.com/Z3R0th-13/FirewallRules — Minimal quick-add script; worth reading for clean param set.
- **HoneyCheng/PowerShell-Scripts firewall log pretty-printer** — https://github.com/topics/windows-firewall — nicer `pfirewall.log` rendering (no canonical repo; topic aggregate).

### Features to Borrow
- Outbound-notification prompt UX from `WFN` — "New connection: chrome.exe → 1.2.3.4:443 — Allow / Block / Temp allow." Fills the biggest gap in native Windows Firewall.
- CSV + JSON round-trip export/import (`MScholtes`) — CSV for spreadsheet auditing, JSON for programmatic diff.
- Domain / Tier-x GPO baselines (`SteveUnderScoreN`) — ship preset bundles: Workstation / Server-DC / PAW.
- Audit-log ingest (event ID 5156 allowed, 5157 blocked) with a filterable grid — show "blocked in the last 5 min" live.
- Remote deploy (`metablaster` experimental) — WinRM / PSRemoting push to a fleet; status roll-up grid.
- Rule linter (`metablaster`) — warn on overlapping/shadowed rules, any-any rules, unused rules (never matched in 30d via audit log).
- Live connections map (`WFN`) — IP geolocation overlay for inbound/outbound sockets.

### Patterns & Architectures Worth Studying
- **WFP callout driver vs netsh wrapper** (`WFN`): real-time prompting needs a callout/filter driver, not polling `Get-NetFirewallRule`. Worth documenting even if out-of-scope.
- **Baseline diff deploy** (`metablaster`): hash a rule bundle, compare to active ruleset, apply only the delta — idempotent and fast.
- **ETW (Microsoft-Windows-Windows Firewall With Advanced Security) subscription** — more efficient than tailing `pfirewall.log`; streamable event source for live UI.
- **PowerShell module + signed manifest + PSGallery publish** (`MScholtes`): makes "live manager" installable via `Install-Module` with Authenticode trust.
```

</details>
