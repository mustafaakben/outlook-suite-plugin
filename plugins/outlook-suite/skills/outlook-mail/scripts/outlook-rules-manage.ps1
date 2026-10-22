param(
    [string]$Name = "",
    [int]$Order = 0,
    [string]$Mailbox = "",
    [switch]$Enable,
    [switch]$Disable,
    [string]$NewName = "",
    [int]$MoveTo = 0,
    [switch]$Delete,
    [switch]$Run,
    [string]$Folder = "Inbox",
    [switch]$IncludeSubfolders,
    [ValidateSet("All", "Read", "Unread")]
    [string]$Messages = "All",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Rules Manage: change one existing rule, chosen by -Name (exact) or -Order (position).
#   -Enable / -Disable     turn it on or off
#   -NewName "..."         rename
#   -MoveTo <n>            change its position in the run order (1 = first)
#   -Delete                remove it
#   -Run                   run it now on -Folder (default Inbox; -IncludeSubfolders;
#                          -Messages All|Read|Unread). Works even when the rule is off.
# Every change is gated: previews unless -Confirm (confirmation for "rules" can be turned off).
# Usage: .\outlook-rules-manage.ps1 -Name "Newsletters" -Disable -Confirm
# Usage: .\outlook-rules-manage.ps1 -Name "Newsletters" -Run -Folder "Inbox" -Messages Unread -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    if (-not ($Name -or $Order)) { Stop-WithError "Choose the rule with -Name or -Order (see outlook-rules-list.ps1)." }
    if (-not ($Enable -or $Disable -or $NewName -or $MoveTo -or $Delete -or $Run)) { Stop-WithError "Nothing to do: give -Enable, -Disable, -NewName, -MoveTo, -Delete or -Run." }
    if ($Enable -and $Disable) { Stop-WithError "Choose -Enable or -Disable, not both." }

    $null = Connect-Outlook
    $store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
    $rules = Get-StoreRules $store
    $rule = Find-Rule -Rules $rules -Name $Name -Order $Order
    $before = ConvertTo-RuleRecord $rule

    $plan = @()
    if ($Enable) { $plan += "turn ON" }
    if ($Disable) { $plan += "turn OFF" }
    if ($NewName) { $plan += "rename to '$NewName'" }
    if ($MoveTo) { $plan += "move to position $MoveTo of $($rules.Count)" }
    if ($Run) { $plan += "run now on $Folder ($Messages messages$(if ($IncludeSubfolders) { ', with subfolders' }))" }
    if ($Delete) { $plan += "DELETE" }

    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "rules"))) {
        Write-Host "=== RULE CHANGE (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Mailbox: $($store.DisplayName)"
        Write-RuleRecord $before -Detailed
        Write-Host "  Change: $($plan -join '; ')" -ForegroundColor Cyan
        Approve-Action -Action "rules" -Confirmed $false -What "this rule change" | Out-Null
        exit 0
    }

    $ruleName = $rule.Name
    if ($Run) {
        $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
        $opt = @{ "All" = 0; "Read" = 1; "Unread" = 2 }[$Messages]
        $rule.Execute($false, $target, [bool]$IncludeSubfolders, $opt)
    }
    $saveNeeded = $false
    if ($Enable) { $rule.Enabled = $true; $saveNeeded = $true }
    if ($Disable) { $rule.Enabled = $false; $saveNeeded = $true }
    if ($NewName) { $rule.Name = $NewName; $ruleName = $NewName; $saveNeeded = $true }
    if ($MoveTo) {
        if ($MoveTo -lt 1 -or $MoveTo -gt $rules.Count) { Stop-WithError "-MoveTo must be between 1 and $($rules.Count)." }
        $rule.ExecutionOrder = $MoveTo; $saveNeeded = $true
    }
    if ($Delete) { $rules.Remove($ruleName); $saveNeeded = $true }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($saveNeeded) { $rules.Save($false) }
    $saveMs = $sw.ElapsedMilliseconds

    $after = $null
    if (-not $Delete) { $after = ConvertTo-RuleRecord (Find-Rule -Rules (Get-StoreRules $store) -Name $ruleName) }
    if ($Json) { Write-Json ([ordered]@{ Result = "done"; Mailbox = $store.DisplayName; Changes = $plan; Before = $before; After = $after; SaveMs = $saveMs }); exit 0 }
    Write-Host "=== RULE UPDATED: $($plan -join '; ') ===" -ForegroundColor Green
    Write-Host "  Mailbox: $($store.DisplayName)   (saved in $saveMs ms)"
    if ($after) { Write-RuleRecord $after -Detailed } else { Write-Host "  Deleted: $($before.Name)" }
}
catch {
    Stop-WithError $_
}
