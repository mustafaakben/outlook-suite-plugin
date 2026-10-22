param(
    [ValidateSet("List", "Snooze", "Dismiss")]
    [string]$Action = "List",
    [string]$Caption = "",
    [int]$Index = 0,
    [int]$Minutes = 15,
    [int]$Days = 7,
    [switch]$All,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Reminders: reminders that are due now or come up within -Days (default 7);
# -All lists every reminder (recurring meetings can add up to hundreds).
#   List                         due and upcoming reminders, with the item's EntryID
#   Snooze -Index n -Minutes 30  remind again later (-All for every due reminder)
#   Dismiss -Index n             dismiss it (gated: previews unless -Confirm)
# -Caption matches by the reminder's text instead of -Index.
# Usage: .\outlook-reminders.ps1
# Usage: .\outlook-reminders.ps1 -Action Snooze -Index 1 -Minutes 60

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $rem = $Outlook.Reminders
    $horizon = (Get-Date).AddDays($Days)
    $listAll = ($Action -eq "List" -and $All)
    $list = @(for ($i = 1; $i -le $rem.Count; $i++) {
        $r = $rem.Item($i)
        $visible = [bool]$r.IsVisible
        if (-not $listAll -and -not $visible) {
            $next = $null; try { $next = [datetime]$r.NextReminderDate } catch { }
            # Not shown as due yet: keep ones that became due in the last 24 h (Outlook may not
            # have popped them up yet) or come up before the horizon. Older ones are stale.
            if (-not $next -or $next -lt (Get-Date).AddDays(-1) -or $next -gt $horizon) { continue }
        }
        $o = [ordered]@{ Index = $i; Caption = $r.Caption; Due = $null; NextReminder = $null; Visible = $visible; ItemType = ""; EntryID = "" }
        try { $o.NextReminder = Format-MailDateIso $r.NextReminderDate } catch { }
        try { $o.Due = Format-MailDateIso $r.OriginalReminderDate } catch { }
        try { $it = $r.Item; $o.EntryID = $it.EntryID; $o.ItemType = @{ 43 = "email"; 26 = "meeting"; 48 = "task"; 40 = "contact" }[[int]$it.Class] } catch { }
        $o
    })

    if ($Action -eq "List") {
        if ($Json) { Write-Json ([ordered]@{ count = $list.Count; reminders = $list }); exit 0 }
        Write-Host "=== REMINDERS: $($list.Count) $(if ($All) { 'in total' } else { "due or within $Days days (of $($rem.Count); -All for every one)" }) ===" -ForegroundColor Cyan
        foreach ($o in $list) {
            $state = if ($o.Visible) { "DUE" } else { "pending" }
            Write-Host "  [$($o.Index)] $state  $($o.Caption)  ($($o.ItemType), next $($o.NextReminder))"
            if ($o.EntryID) { Write-Host "      EntryID: $($o.EntryID)" -ForegroundColor DarkGray }
        }
        if ($list.Count -eq 0) { Write-Host "No reminders." -ForegroundColor Gray }
        exit 0
    }

    $targets = @(if ($All) { $list | Where-Object { $_.Visible } }
        elseif ($Index) { $list | Where-Object { $_.Index -eq $Index } }
        elseif ($Caption) { $list | Where-Object { $_.Caption -like "*$Caption*" } })
    if ($targets.Count -eq 0) { Stop-WithError "No matching reminder. Use -Index (see -Action List), -Caption, or -All (due ones)." }

    if ($Action -eq "Dismiss" -and -not ($Confirm -or -not (Test-ConfirmRequired -Action "reminders"))) {
        Write-Host "=== DISMISS REMINDERS (PREVIEW) ===" -ForegroundColor Yellow
        foreach ($t in $targets) { Write-Host "  [$($t.Index)] $($t.Caption)" }
        Approve-Action -Action "reminders" -Confirmed $false -What "dismissing" | Out-Null
        exit 0
    }
    $done = @(); $failed = @()
    foreach ($t in ($targets | Sort-Object Index -Descending)) {
        try {
            $r = $rem.Item($t.Index)
            if ($Action -eq "Snooze") { $r.Snooze($Minutes) } else { $r.Dismiss() }
            $done += $t.Caption
        } catch { $failed += "$($t.Caption): $($_.Exception.Message)" }
    }
    if ($Json) { Write-Json ([ordered]@{ Action = $Action; Done = $done; Failed = $failed }); if ($failed) { exit 1 } else { exit 0 } }
    Write-Host "=== $($Action.ToUpper()) DONE ($($done.Count)) ===" -ForegroundColor Green
    foreach ($d in $done) { Write-Host "  $d" }
    foreach ($f in $failed) { Write-Host "  Failed: $f" -ForegroundColor Red }
    if ($failed) { Write-Host "  (Outlook only snoozes/dismisses reminders that are currently due.)" -ForegroundColor Gray; exit 1 }
}
catch {
    Stop-WithError $_
}
