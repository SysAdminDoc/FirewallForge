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

Assert-PowerShellParses $managerPath
Assert-PowerShellParses $editorPath

$managerSource = Get-Content -LiteralPath $managerPath -Raw
Assert-True ($managerSource -match 'function Show-ProgramRuleWizard') "Program wizard function is missing."
Assert-True ($managerSource -match '\$btnProgramWizard\.Add_Click') "Program wizard button is not wired."
Assert-True ($managerSource -match 'Block In \+ Out') "Program wizard block preset is missing."
Assert-True ($managerSource -match 'Allow In') "Program wizard inbound preset is missing."
Assert-True ($managerSource -match 'Allow Out') "Program wizard outbound preset is missing."

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
