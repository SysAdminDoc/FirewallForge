# Changelog

All notable changes to FirewallForge will be documented in this file.

## Unreleased

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

## [v1.1.0] - %Y->- (HEAD -> main, tag: v1.1.0, origin/main)

- Added: Add project icon to README
- v1.1.0 - Duplicate detection, stats panel, quick-block, diff view
- Initial commit - FirewallForge
