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
    # The due/upcoming window only shapes the List display. Snooze/Dismiss of an explicit
    # target (-Caption/-Index) look through every reminder.
    $listAll = ($Action -eq "List" -and $All) -or ($Action -ne "List" -and -not $All)
    function Get-ReminderList {
      @(for ($i = 1; $i -le $rem.Count; $i++) {
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
    }
    $list = Get-ReminderList

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

    function Select-Targets($l) {
        @(if ($All) { $l | Where-Object { $_.Visible } }
            elseif ($Index) { $l | Where-Object { $_.Index -eq $Index } }
            elseif ($Caption) { $l | Where-Object { $_.Caption -like "*$Caption*" } })
    }
    $targets = Select-Targets $list
    if ($targets.Count -eq 0) {
        # Outlook changes the Reminders collection while it fires reminders; look once more.
        Start-Sleep -Seconds 2
        $list = Get-ReminderList
        $targets = Select-Targets $list
    }
    if ($targets.Count -eq 0) { Stop-WithError "No matching reminder. Use -Index (see -Action List), -Caption, or -All (due ones)." }

    if ($Action -eq "Dismiss" -and -not ($Confirm -or -not (Test-ConfirmRequired -Action "reminders"))) {
        Write-Host "=== DISMISS REMINDERS (PREVIEW) ===" -ForegroundColor Yellow
        foreach ($t in $targets) { Write-Host "  [$($t.Index)] $($t.Caption)" }
        Approve-Action -Action "reminders" -Confirmed $false -What "dismissing" | Out-Null
        exit 0
    }
    # Outlook reorders the Reminders collection as it fires reminders, so a position from the
    # listing may now point at another reminder. Re-find each target by caption + item EntryID
    # right before acting, and never act on anything else.
    function Find-LiveReminder($t) {
        for ($i = 1; $i -le $rem.Count; $i++) {
            $r = $rem.Item($i)
            if ($r.Caption -ne $t.Caption) { continue }
            $eid = ""; try { $eid = $r.Item.EntryID } catch { }
            if (-not $t.EntryID -or $eid -eq $t.EntryID) { return $r }
        }
        return $null
    }
    $done = @(); $failed = @()
    foreach ($t in $targets) {
        $ok = $false; $err = "not found any more"
        foreach ($attempt in 1, 2) {
            $r = Find-LiveReminder $t
            if ($r) {
                try { if ($Action -eq "Snooze") { $r.Snooze($Minutes) } else { $r.Dismiss() }; $ok = $true; break }
                catch { $err = $_.Exception.Message }
            }
            Start-Sleep -Seconds 2
        }
        if ($ok) { $done += $t.Caption } else { $failed += "$($t.Caption): $err" }
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
