# ROADMAP

Backlog for FirewallForge. Two-tool suite (live manager + offline editor) for Windows Firewall
rules. Goal: close the gap with Windows Firewall Control (WFC) without adopting its paid model.

## Planned Features

### FirewallManager (live)
- **Per-program rule wizard** — pick an exe, show all existing rules matching its path, offer
  Block / Allow In / Allow Out presets in one click.
- **New-connection prompts** — optional interactive mode that watches
  `Microsoft-Windows-Security-Auditing` event 5157 (blocked) and 5156 (allowed) and pops a
  decision dialog (WFC-style). Default off.
- **Outbound-block-by-default profile** — one-click switch that flips outbound default to Block,
  adds essential OS allow rules, and snapshots prior state for rollback.
- **Rule ordering / priority view** — show which rules win for a given flow using
  `Test-NetConnection` + predicate walk; WF doesn't expose this clearly.
- **Service rule support** — create rules scoped to a service name (`Service=svchost-...`) not
  just program path.
- **Group operations** — enable/disable/delete by group match, not just individual rows.
- **Profile-aware editing** — separate Domain / Private / Public toggles per rule with a single
  row.
- **Log viewer tab** — tail `pfirewall.log` with coloring and filter by action/direction/port.

### FirewallRulesEditor (offline)
- **Diff between two backups** — not just against original import; load any two `.fwbackup` and
  show added/removed/modified.
- **Merge-apply strategy picker** — "prefer newer", "prefer imported", "manual conflict" like
  three-way merge tools.
- **Rule template library** — save rule rows as reusable templates (e.g. "SMB block inbound on
  Public"), insert with one click.
- **Policy export** — emit `netsh advfirewall` script, PowerShell `New-NetFirewallRule` script, or
  GPO `.pol` fragment.

### Cross-cutting
- **JSON schema for `.fwbackup`** so external tooling can validate.
- **Group Policy awareness** — detect and flag rules delivered by GPO (read-only, can't be
  deleted in the UI); currently the tool operates only on local store.
- **Hyper-V / WSL** profile awareness — call out rules scoped to those interfaces since WF
  interface filtering is non-obvious.
- **Scheduled backups** via Task Scheduler — rotate daily/weekly with retention count.
- **IPv6 coverage audit** — flag IPv4-only rules that should also have v6 twins.

### UI
- **Dark theme across both tools** (currently only partial).
- **Regex search** in the DataGrid filter.
- **Saved views** — named column sets + sort + filter combos.
- **Bulk comment/tag column** — user-defined tags stored as comment text in the rule, surfaced as
  a column.

### Distribution
- **Code-signed script bundle** — Authenticode on both `.ps1` files so they run without
  ExecutionPolicy bypass on managed machines.
- **Signed module format (`.psm1`)** as an alternative invocation.

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

- **DNS-level complement** — show Windows DNS client log correlated with blocked rules so users
  see what domain a blocked app was trying to resolve.
- **Geofencing** — auto-add remote-address allowlists scoped to a country CIDR set (download from
  IP2Country), opt-in only.
- **Rule health score** — flag overly-broad rules (Any/Any/Any), duplicates, rules pointing at
  nonexistent exes.
- **Command-line wrapper** (`FirewallForge.ps1 apply profile.json`) for Intune deployment.
- **Multi-machine compare** — diff `.fwbackup` from several endpoints to find drift.
- **MSIX / Winget** distribution.

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
