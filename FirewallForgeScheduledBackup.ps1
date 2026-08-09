<#
.SYNOPSIS
    Creates a non-interactive FirewallForge .fwbackup package and rotates scheduled backups.
.NOTES
    This worker is launched by Task Scheduler and does not open the WPF manager.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$BackupPath,

    [ValidateRange(1, 3650)]
    [int]$RetentionCount = 14
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $BackupPath)) {
    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
}

$tempFile = Join-Path $env:TEMP ("FirewallForgeScheduled_{0}.wfw" -f [guid]::NewGuid().ToString("N"))
try {
    & netsh advfirewall export $tempFile 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tempFile)) {
        throw "netsh advfirewall export failed."
    }

    $rules = @(Get-NetFirewallRule -ErrorAction Stop | ForEach-Object {
        [ordered]@{
            Name = $_.Name
            DisplayName = $_.DisplayName
            Description = $_.Description
            Direction = $_.Direction.ToString()
            Action = $_.Action.ToString()
            Enabled = $_.Enabled.ToString()
            Profile = $_.Profile.ToString()
        }
    })

    $backup = [ordered]@{
        BackupDate = (Get-Date).ToString("o")
        ComputerName = $env:COMPUTERNAME
        RuleCount = $rules.Count
        NetshBackup = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tempFile))
        RuleDetails = $rules
        ScheduledBackup = $true
    }
    $outputPath = Join-Path $BackupPath ("FirewallScheduledBackup_{0}.fwbackup" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $backup | ConvertTo-Json -Depth 20 -Compress | Set-Content -LiteralPath $outputPath -Encoding UTF8

    $scheduledBackups = @(Get-ChildItem -LiteralPath $BackupPath -Filter "FirewallScheduledBackup_*.fwbackup" -File | Sort-Object LastWriteTime -Descending)
    if ($scheduledBackups.Count -gt $RetentionCount) {
        foreach ($oldBackup in $scheduledBackups | Select-Object -Skip $RetentionCount) {
            Remove-Item -LiteralPath $oldBackup.FullName -Force
        }
    }

    Write-Output $outputPath
}
finally {
    if (Test-Path -LiteralPath $tempFile) {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}
