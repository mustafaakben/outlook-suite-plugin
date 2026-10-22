param(
    [string]$Mailbox = "",
    [string]$Name = "",
    [switch]$Enabled,
    [switch]$Detailed,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Rules List: rules of a mailbox in execution order. -Detailed shows every condition,
# exception and action (including ones made in the Rules Wizard that scripts cannot create).
# -Mailbox picks another mailbox in the profile (default: your default mailbox).
# -Name filters by name (substring).
# Usage: .\outlook-rules-list.ps1 -Detailed
# Usage: .\outlook-rules-list.ps1 -Mailbox "someone@x.com" -Enabled

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
    $rules = Get-StoreRules $store
    $records = @(for ($i = 1; $i -le $rules.Count; $i++) {
        $r = $rules.Item($i)
        if ($Enabled -and -not $r.Enabled) { continue }
        if ($Name -and $r.Name -notlike "*$Name*") { continue }
        ConvertTo-RuleRecord $r
    }) | Sort-Object { $_.Order }
    $on = @(for ($i = 1; $i -le $rules.Count; $i++) { if ($rules.Item($i).Enabled) { 1 } }).Count

    if ($Json) { Write-Json ([ordered]@{ mailbox = $store.DisplayName; total = $rules.Count; enabled = $on; count = @($records).Count; rules = @($records) }); exit 0 }
    Write-Host "=== RULES: $($store.DisplayName) ===" -ForegroundColor Cyan
    Write-Host "$($rules.Count) rules, $on on$(if ($Enabled -or $Name) { "; showing $(@($records).Count)" })" -ForegroundColor Gray
    foreach ($r in $records) { Write-RuleRecord $r -Detailed:$Detailed }
    if ($rules.Count -eq 0) { Write-Host "No rules. Create one with outlook-rules-create.ps1." -ForegroundColor Gray }
    elseif (-not $Detailed) { Write-Host "Add -Detailed for conditions and actions. Change rules with outlook-rules-manage.ps1." -ForegroundColor DarkGray }
}
catch {
    Stop-WithError $_
}
