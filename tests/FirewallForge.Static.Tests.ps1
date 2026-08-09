$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        $failures.Add($Message)
    }
}

function Assert-PowerShellParses {
    param([string]$Path)

    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) | Out-Null
    Assert-True ($errors.Count -eq 0) "$([System.IO.Path]::GetFileName($Path)) has PowerShell parse errors: $($errors -join '; ')"
}

$managerPath = Join-Path $repoRoot "FirewallManager.ps1"
$editorPath = Join-Path $repoRoot "FirewallRulesEditor.ps1"
$schemaPath = Join-Path $repoRoot "fwbackup.schema.json"
$scheduledWorkerPath = Join-Path $repoRoot "FirewallForgeScheduledBackup.ps1"

Assert-PowerShellParses $managerPath
Assert-PowerShellParses $editorPath
Assert-PowerShellParses $scheduledWorkerPath

$managerSource = Get-Content -LiteralPath $managerPath -Raw
Assert-True ($managerSource -match 'function Show-ProgramRuleWizard') "Program wizard function is missing."
Assert-True ($managerSource -match '\$btnProgramWizard\.Add_Click') "Program wizard button is not wired."
Assert-True ($managerSource -match 'Block In \+ Out') "Program wizard block preset is missing."
Assert-True ($managerSource -match 'Allow In') "Program wizard inbound preset is missing."
Assert-True ($managerSource -match 'Allow Out') "Program wizard outbound preset is missing."
Assert-True ($managerSource -match 'function Toggle-ConnectionMonitor') "Connection monitor toggle is missing."
Assert-True ($managerSource -match 'Id = @\(5156, 5157\)') "Connection monitor event query is missing 5156/5157."
Assert-True ($managerSource -match '\$btnConnectionMonitor\.Add_Click') "Connection monitor button is not wired."
Assert-True ($managerSource -match 'function Enable-OutboundLockdown') "Outbound lockdown function is missing."
Assert-True ($managerSource -match 'function Restore-LockdownSnapshot') "Lockdown rollback function is missing."
Assert-True ($managerSource -match 'DefaultOutboundAction Block') "Outbound lockdown does not change the default action."
Assert-True ($managerSource -match 'function New-LockdownAllowRule') "Essential lockdown allow-rule helper is missing."
Assert-True ($managerSource -match 'function Test-FirewallRulePriority') "Rule priority evaluator is missing."
Assert-True ($managerSource -match 'function Test-FirewallRuleFlowMatch') "Rule predicate walk is missing."
Assert-True ($managerSource -match 'Test-NetConnection') "Rule priority view does not expose endpoint testing."
Assert-True ($managerSource -match '\$btnRulePriority\.Add_Click') "Rule priority button is not wired."
Assert-True ($managerSource -match 'function Show-GroupOperations') "Group operations function is missing."
Assert-True ($managerSource -match 'Remove-NetFirewallRule -Name \$rule\.Name') "Group delete operation is missing."
Assert-True ($managerSource -match '\$btnGroupOps\.Add_Click') "Group operations button is not wired."
Assert-True ($managerSource -match 'function Show-FirewallLogViewer') "Firewall log viewer is missing."
Assert-True ($managerSource -match 'function Read-FirewallLogEntries') "Firewall log parser is missing."
Assert-True ($managerSource -match 'pfirewall\.log') "Standard firewall log path is missing."
Assert-True ($managerSource -match '\$btnLogViewer\.Add_Click') "Firewall log viewer button is not wired."
Assert-True ($managerSource -match 'function Show-ScheduledBackupDialog') "Scheduled backup dialog is missing."
Assert-True ($managerSource -match 'Register-ScheduledTask') "Scheduled backup registration is missing."
Assert-True ($managerSource -match '\$btnScheduleBackups\.Add_Click') "Scheduled backup button is not wired."
Assert-True ($managerSource -match 'function Show-IPv6CoverageAudit') "IPv6 coverage audit is missing."
Assert-True ($managerSource -match 'function Test-IsIPv4OnlyRule') "IPv4-only rule heuristic is missing."
Assert-True ($managerSource -match '\$btnAuditRules\.Add_Click') "IPv6 audit button is not wired."
Assert-True ($managerSource -match 'function Show-SavedViews') "Saved views window is missing."
Assert-True ($managerSource -match 'function Get-CurrentSavedViewDefinition') "Saved view capture is missing."
Assert-True ($managerSource -match 'function Apply-SavedView') "Saved view application is missing."
Assert-True ($managerSource -match '\$btnSavedViews\.Add_Click') "Saved views button is not wired."

$editorSource = Get-Content -LiteralPath $editorPath -Raw
Assert-True ($editorSource -match 'function Compare-FWBackups') "Two-backup comparison function is missing."
Assert-True ($editorSource -match 'function Compare-BackupRuleSets') "Two-backup diff engine is missing."
Assert-True ($editorSource -match '\$btnCompareBackups\.Add_Click') "Two-backup comparison button is not wired."
Assert-True ($editorSource -match 'function Show-MergeStrategyPicker') "Merge strategy picker is missing."
Assert-True ($editorSource -match 'Prefer newer') "Prefer-newer merge strategy is missing."
Assert-True ($editorSource -match 'Manual conflict') "Manual merge strategy is missing."
Assert-True ($editorSource -match 'function Resolve-MergeConflict') "Manual merge conflict resolver is missing."
Assert-True ($editorSource -match 'function Show-TemplateLibrary') "Template library window is missing."
Assert-True ($editorSource -match 'function Save-RuleTemplate') "Template save function is missing."
Assert-True ($editorSource -match 'function Insert-RuleTemplate') "Template insertion function is missing."
Assert-True ($editorSource -match 'FirewallForge_Templates') "Template storage folder is missing."
Assert-True ($editorSource -match '\$btnTemplates\.Add_Click') "Template library button is not wired."
Assert-True ($editorSource -match 'function Export-SelectedPolicy') "Policy export function is missing."
Assert-True ($editorSource -match 'function ConvertTo-NetshFirewallScript') "netsh policy export is missing."
Assert-True ($editorSource -match 'function ConvertTo-PowerShellFirewallScript') "PowerShell policy export is missing."
Assert-True ($editorSource -match 'function Write-GpoRegistryPol') "Registry.pol policy export is missing."
Assert-True ($editorSource -match 'REGFILE_SIGNATURE') "Registry.pol signature is missing."
Assert-True ($editorSource -match '\$btnExportPolicy\.Add_Click') "Policy export button is not wired."
Assert-True ($editorSource -match 'function Test-EditorRuleSearchMatch') "Editor regex search matcher is missing."
Assert-True ($editorSource -match '\$chkRegexSearch\.Add_Click') "Editor regex search toggle is not wired."

# Exercise the formatters without starting WPF or changing firewall state.
$editorTokens = $null
$editorParseErrors = $null
$editorAst = [System.Management.Automation.Language.Parser]::ParseFile($editorPath, [ref]$editorTokens, [ref]$editorParseErrors)
$policyFunctionNames = @(
    "ConvertTo-NetshFirewallScript",
    "ConvertTo-PowerShellFirewallScript",
    "New-GpoFirewallRuleData",
    "Write-RegistryPolString",
    "Write-RegistryPolRecord",
    "Write-GpoRegistryPol"
)
$policyFunctionText = foreach ($name in $policyFunctionNames) {
    $functionAst = $editorAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($functionAst) { $functionAst.Extent.Text }
}
try {
    . ([scriptblock]::Create(($policyFunctionText -join [Environment]::NewLine)))
    $sampleRule = [PSCustomObject]@{
        Name = "{11111111-1111-1111-1111-111111111111}"
        DisplayName = "FirewallForge smoke rule"
        Description = "Smoke test"
        Direction = "Outbound"
        Action = "Allow"
        Enabled = "True"
        Profile = "Private"
        Protocol = "TCP"
        LocalPort = "Any"
        RemotePort = "443"
        Program = "C:\\Windows\\System32\\smoke.exe"
    }
    $netshOutput = ConvertTo-NetshFirewallScript -Rules @($sampleRule)
    $powerShellOutput = ConvertTo-PowerShellFirewallScript -Rules @($sampleRule)
    Assert-True ($netshOutput -match "netsh advfirewall firewall add rule") "netsh formatter did not emit a rule."
    Assert-True ($powerShellOutput -match "New-NetFirewallRule") "PowerShell formatter did not emit a rule."
    $policyPath = Join-Path ([System.IO.Path]::GetTempPath()) ("FirewallForgePolicy_{0}.pol" -f [guid]::NewGuid().ToString("N"))
    try {
        Write-GpoRegistryPol -Rules @($sampleRule) -Path $policyPath
        $policyBytes = [System.IO.File]::ReadAllBytes($policyPath)
        Assert-True ($policyBytes.Length -gt 16) "Registry.pol output is unexpectedly short."
        Assert-True ($policyBytes[0] -eq 0x50 -and $policyBytes[1] -eq 0x52 -and $policyBytes[2] -eq 0x65 -and $policyBytes[3] -eq 0x67) "Registry.pol signature is invalid."
        $policyText = [System.Text.Encoding]::Unicode.GetString($policyBytes)
        Assert-True ($policyText -match "FirewallRules") "Registry.pol key is missing."
        Assert-True ($policyText -match "FirewallForge smoke rule") "Registry.pol rule payload is missing."
    }
    finally {
        if (Test-Path -LiteralPath $policyPath) { Remove-Item -LiteralPath $policyPath -Force }
    }
}
catch {
    Assert-True $false "Policy formatter smoke test failed: $($_.Exception.Message)"
}

try {
    $null = Get-Content -LiteralPath $schemaPath -Raw | ConvertFrom-Json
    Assert-True $true "JSON schema parsed."
}
catch {
    Assert-True $false "fwbackup.schema.json is invalid: $($_.Exception.Message)"
}

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Output "PASS FirewallForge static smoke tests"
