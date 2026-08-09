<#[
.SYNOPSIS
    Applies a declarative FirewallForge firewall profile without opening the GUI.
.DESCRIPTION
    The apply command is intended for Intune, Task Scheduler, and other unattended
    deployment workflows. It creates or updates only the rules named by the profile.
    Use -Prune to remove rules in the profile's managed group that are not present
    in the desired profile.
.EXAMPLE
    .\FirewallForge.ps1 validate .\profiles\workstation.json
.EXAMPLE
    .\FirewallForge.ps1 apply .\profiles\workstation.json -Prune
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("apply", "validate")]
    [string]$Command,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$ProfilePath,

    [switch]$Prune
)

$ErrorActionPreference = "Stop"

function Get-ProfileProperty {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [object]$DefaultValue = $null
    )

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $DefaultValue
    }
    return $property.Value
}

function ConvertTo-StringArray {
    param([object]$Value)

    if ($null -eq $Value) {
        return @()
    }
    if ($Value -is [System.Array]) {
        return @($Value | ForEach-Object { [string]$_ })
    }
    return @([string]$Value)
}

function ConvertTo-ProfileText {
    param(
        [object]$Value,
        [string]$Default = "Any"
    )

    $values = @(ConvertTo-StringArray -Value $Value | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($values.Count -eq 0) {
        return $Default
    }
    return ($values -join ",")
}

function Test-ProfileBoolean {
    param(
        [object]$Value,
        [bool]$Default = $true
    )

    if ($null -eq $Value) {
        return $Default
    }
    if ($Value -is [bool]) {
        return [bool]$Value
    }
    $parsed = $false
    if ([bool]::TryParse(([string]$Value), [ref]$parsed)) {
        return $parsed
    }
    throw "Expected a Boolean value, but received '$Value'."
}

function Assert-ProfileValue {
    param(
        [string]$Value,
        [string]$PropertyName,
        [string[]]$AllowedValues
    )

    if ($AllowedValues -notcontains $Value) {
        throw "$PropertyName must be one of: $($AllowedValues -join ', '). Received '$Value'."
    }
}

function Read-FirewallForgeProfile {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolvedPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path)
    $profileDocument = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $profileDocument -or $profileDocument -isnot [psobject]) {
        throw "Profile '$resolvedPath' did not contain a JSON object."
    }

    $version = Get-ProfileProperty -Object $profileDocument -Name "Version" -DefaultValue 1
    $versionNumber = 0
    if (-not [int]::TryParse(([string]$version), [ref]$versionNumber) -or $versionNumber -ne 1) {
        throw "Unsupported profile Version '$version'. The only supported version is 1."
    }

    $profileName = [string](Get-ProfileProperty -Object $profileDocument -Name "Name" -DefaultValue ([System.IO.Path]::GetFileNameWithoutExtension($resolvedPath)))
    if ([string]::IsNullOrWhiteSpace($profileName)) {
        throw "Profile Name cannot be empty."
    }

    $group = [string](Get-ProfileProperty -Object $profileDocument -Name "Group" -DefaultValue "FirewallForge:$profileName")
    if ([string]::IsNullOrWhiteSpace($group)) {
        throw "Profile Group cannot be empty."
    }

    $rawRules = Get-ProfileProperty -Object $profileDocument -Name "Rules"
    if ($null -eq $rawRules) {
        throw "Profile must contain a Rules array."
    }
    $rules = @($rawRules)
    $seenNames = @{}
    $normalizedRules = New-Object System.Collections.Generic.List[object]

    foreach ($rawRule in $rules) {
        if ($null -eq $rawRule) {
            throw "Rules cannot contain null entries."
        }

        $name = [string](Get-ProfileProperty -Object $rawRule -Name "Name")
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw "Every rule must have a non-empty Name."
        }
        if ($name.Length -gt 255) {
            throw "Rule '$name' exceeds the 255-character firewall rule name limit."
        }
        $nameKey = $name.ToUpperInvariant()
        if ($seenNames.ContainsKey($nameKey)) {
            throw "Rule name '$name' is duplicated in the profile."
        }
        $seenNames[$nameKey] = $true

        $displayName = [string](Get-ProfileProperty -Object $rawRule -Name "DisplayName" -DefaultValue $name)
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            throw "Rule '$name' must have a non-empty DisplayName."
        }

        $direction = ([string](Get-ProfileProperty -Object $rawRule -Name "Direction" -DefaultValue "Inbound")).Trim()
        $action = ([string](Get-ProfileProperty -Object $rawRule -Name "Action" -DefaultValue "Allow")).Trim()
        Assert-ProfileValue -Value $direction -PropertyName "Rule '$name' Direction" -AllowedValues @("Inbound", "Outbound")
        Assert-ProfileValue -Value $action -PropertyName "Rule '$name' Action" -AllowedValues @("Allow", "Block")

        $profileValue = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Profile")
        foreach ($profileToken in @($profileValue -split ",")) {
            Assert-ProfileValue -Value $profileToken.Trim() -PropertyName "Rule '$name' Profile" -AllowedValues @("Any", "Domain", "Private", "Public")
        }

        $protocol = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Protocol")
        if ($protocol -ne "Any" -and $protocol -notmatch '^(?i:TCP|UDP|ICMPv4|ICMPv6|\d{1,3})$') {
            throw "Rule '$name' Protocol must be Any, TCP, UDP, ICMPv4, ICMPv6, or a protocol number. Received '$protocol'."
        }

        $description = Get-ProfileProperty -Object $rawRule -Name "Description" -DefaultValue ""
        $enabled = Test-ProfileBoolean -Value (Get-ProfileProperty -Object $rawRule -Name "Enabled")

        $normalizedRules.Add([PSCustomObject]@{
                Name = $name
                DisplayName = $displayName
                Description = if ($null -eq $description) { "" } else { [string]$description }
                Direction = $direction
                Action = $action
                Enabled = $enabled
                Profile = $profileValue
                Protocol = $protocol
                LocalPort = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "LocalPort")
                RemotePort = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "RemotePort")
                LocalAddress = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "LocalAddress")
                RemoteAddress = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "RemoteAddress")
                Program = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Program")
                Service = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Service")
                InterfaceAlias = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "InterfaceAlias") -Default ""
                InterfaceType = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "InterfaceType") -Default ""
                IcmpType = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "IcmpType") -Default ""
                EdgeTraversalPolicy = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "EdgeTraversalPolicy") -Default ""
                LocalUser = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "LocalUser") -Default ""
                RemoteUser = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "RemoteUser") -Default ""
                RemoteMachine = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "RemoteMachine") -Default ""
                Platform = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Platform") -Default ""
                Package = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Package") -Default ""
                PackageFamilyName = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "PackageFamilyName") -Default ""
                Authentication = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Authentication") -Default ""
                Encryption = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Encryption") -Default ""
                Owner = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "Owner") -Default ""
                PolicyAppId = ConvertTo-ProfileText -Value (Get-ProfileProperty -Object $rawRule -Name "PolicyAppId") -Default ""
            })
    }

    [PSCustomObject]@{
        Path = $resolvedPath
        Version = $versionNumber
        Name = $profileName
        Group = $group
        Prune = Test-ProfileBoolean -Value (Get-ProfileProperty -Object $profileDocument -Name "Prune") -Default $false
        Rules = $normalizedRules.ToArray()
    }
}

function ConvertTo-FirewallRuleParameter {
    param(
        [Parameter(Mandatory = $true)][object]$Rule,
        [Parameter(Mandatory = $true)][string]$Group
    )

    $parameters = @{
        Name = $Rule.Name
        DisplayName = $Rule.DisplayName
        Description = $Rule.Description
        Direction = $Rule.Direction
        Action = $Rule.Action
        Enabled = $Rule.Enabled
        Profile = $Rule.Profile
        Group = $Group
    }

    $optionalProperties = @(
        "Protocol", "LocalPort", "RemotePort", "LocalAddress", "RemoteAddress",
        "Program", "Service", "InterfaceAlias", "InterfaceType", "IcmpType",
        "EdgeTraversalPolicy", "LocalUser", "RemoteUser", "RemoteMachine", "Platform",
        "Package", "PackageFamilyName", "Authentication", "Encryption", "Owner", "PolicyAppId"
    )
    foreach ($propertyName in $optionalProperties) {
        $value = [string]$Rule.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value) -and $value -ne "Any") {
            $parameters[$propertyName] = $value
        }
    }
    return $parameters
}

function Test-IsGpoFirewallRule {
    param([object]$Rule)

    $sourceType = [string](Get-ProfileProperty -Object $Rule -Name "PolicyStoreSourceType" -DefaultValue "")
    $source = [string](Get-ProfileProperty -Object $Rule -Name "PolicyStoreSource" -DefaultValue "")
    return $sourceType -match "GroupPolicy|GPO" -or ($source -and $source -ne "PersistentStore" -and $source -ne "Local")
}

function Invoke-FirewallForgeApply {
    param(
        [Parameter(Mandatory = $true)][object]$Profile,
        [switch]$PruneRules
    )

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "The apply command requires Administrator privileges. Use validate for non-privileged profile checks."
    }

    $desiredNames = @{}
    $created = 0
    $updated = 0
    $removed = 0

    foreach ($rule in $Profile.Rules) {
        $desiredNames[$rule.Name.ToUpperInvariant()] = $true
        $parameters = ConvertTo-FirewallRuleParameter -Rule $rule -Group $Profile.Group
        $existing = @(Get-NetFirewallRule -Name $rule.Name -ErrorAction SilentlyContinue)
        if ($existing.Count -gt 1) {
            throw "More than one firewall rule matched name '$($rule.Name)'."
        }

        if ($existing.Count -eq 0) {
            New-NetFirewallRule @parameters -ErrorAction Stop -WhatIf:$WhatIfPreference | Out-Null
            $created++
        }
        else {
            if (Test-IsGpoFirewallRule -Rule $existing[0]) {
                throw "Rule '$($rule.Name)' is controlled by Group Policy and cannot be managed by the local profile."
            }
            Set-NetFirewallRule @parameters -ErrorAction Stop -WhatIf:$WhatIfPreference | Out-Null
            $updated++
        }
    }

    if ($PruneRules) {
        $managedRules = @(Get-NetFirewallRule -Group $Profile.Group -ErrorAction SilentlyContinue)
        foreach ($managedRule in $managedRules) {
            if (-not $desiredNames.ContainsKey(([string]$managedRule.Name).ToUpperInvariant())) {
                if (Test-IsGpoFirewallRule -Rule $managedRule) {
                    throw "Managed group contains Group Policy rule '$($managedRule.Name)'; pruning was stopped before removing it."
                }
                Remove-NetFirewallRule -Name $managedRule.Name -Confirm:$false -ErrorAction Stop -WhatIf:$WhatIfPreference
                $removed++
            }
        }
    }

    [ordered]@{
        Status = "Applied"
        Profile = $Profile.Name
        Group = $Profile.Group
        Created = $created
        Updated = $updated
        Removed = $removed
        WhatIf = [bool]$WhatIfPreference
    }
}

try {
    $firewallProfile = Read-FirewallForgeProfile -Path $ProfilePath
    if ($Command -eq "validate") {
        [ordered]@{
            Status = "Valid"
            Profile = $firewallProfile.Name
            Group = $firewallProfile.Group
            RuleCount = $firewallProfile.Rules.Count
            Path = $firewallProfile.Path
        } | ConvertTo-Json -Compress
        exit 0
    }

    $shouldPrune = $Prune -or $firewallProfile.Prune
    $summary = Invoke-FirewallForgeApply -Profile $firewallProfile -PruneRules:$shouldPrune
    $summary | ConvertTo-Json -Compress
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
