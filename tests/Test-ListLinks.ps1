<#
.SYNOPSIS
  Every list, search and write result that names an email carries its Message-ID and WebLink.
  The reference values come from outlook-triage.ps1 (checked against Microsoft Graph in 3.2.0).
  Read-only for mail lists. Writes only to a test draft (subject "[suite-test] ...", addressed to
  your own mailbox, never sent), its copy and forward, all ending in the hidden folder
  "Deleted Items/_suite-test", plus one task that is deleted again. Prints counts and PASS/FAIL only.
#>
param(
    [ValidateSet("pwsh", "powershell")][string]$Exe = "pwsh",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$TasksDir = (Resolve-Path (Join-Path $ScriptsDir "../../outlook-tasks/scripts")).Path
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$dir, [string]$script, [string[]]$a, [switch]$Text) {
    $argv = @("-NoProfile", "-File", (Join-Path $dir $script)) + $a
    if (-not $Text) { $argv += "-Json" }
    $out = & $Exe @argv 2>&1 | Out-String
    if ($Text) { return $out }
    try { return ($out | ConvertFrom-Json) } catch { Write-Host "  (not JSON from $script)" -ForegroundColor DarkGray; return $null }
}

# Reference: EntryID -> Message-ID and WebLink for every Inbox email of the last 7 days.
$truth = @{}
foreach ($i in (Run $ScriptsDir "outlook-triage.ps1" @("all", "-Days", "7", "-Limit", "100000")).items) { $truth[$i.EntryID] = $i }
Check "reference: triage returned emails" ($truth.Count -gt 0) "($($truth.Count))"

# Each row has both fields, and rows that are in the reference match it exactly.
function CheckRows([string]$name, $rows) {
    $rows = @($rows | Where-Object { $_.EntryID })
    $noLink = @($rows | Where-Object { $_.WebLink -notmatch '^https://' }).Count
    $noMid = @($rows | Where-Object { -not $_.MessageID }).Count
    $known = @($rows | Where-Object { $truth.ContainsKey($_.EntryID) })
    $wrong = @($known | Where-Object { $_.WebLink -ne $truth[$_.EntryID].WebLink -or $_.MessageID -ne $truth[$_.EntryID].MessageID }).Count
    if ($known.Count -eq 0 -and $rows.Count) {
        # None in the Inbox reference (other folders): compare the first row with outlook-read-body.ps1.
        $one = Run $ScriptsDir "outlook-read-body.ps1" @("-EntryID", $rows[0].EntryID)
        $known = @($rows[0]); $wrong = [int]($one.WebLink -ne $rows[0].WebLink -or $one.MessageID -ne $rows[0].MessageID)
    }
    Check "${name}: every row has a WebLink and a Message-ID" ($rows.Count -gt 0 -and $noLink -eq 0 -and $noMid -eq 0) "($($rows.Count) rows, $noLink without link, $noMid without Message-ID)"
    Check "${name}: both match the reference" ($known.Count -gt 0 -and $wrong -eq 0) "($($known.Count) compared, $wrong different)"
}

CheckRows "mail_read" (Run $ScriptsDir "outlook-read.ps1" @("-Days", "7", "-Limit", "25")).items
CheckRows "mail_find" (Run $ScriptsDir "outlook-find.ps1" @("-Days", "7", "-Limit", "25")).items
CheckRows "mail_search" (Run $ScriptsDir "outlook-search.ps1" @("-Days", "7", "-Subject", "e", "-Limit", "25")).items
CheckRows "mail_search_advanced" (Run $ScriptsDir "outlook-search-advanced.ps1" @("-Days", "7", "-Subject", "e", "-Limit", "25")).items
$prio = Run $ScriptsDir "outlook-priority.ps1" @("-Days", "7", "-Limit", "25")
CheckRows "mail_priority" $prio.items
Check "mail_priority: JSON rows do not carry the internal StoreID" (@($prio.items | Where-Object { $_.PSObject.Properties.Name -contains "StoreID" }).Count -eq 0)

# mail_everything: search the index for a word from a recent subject.
$word = $null
foreach ($i in $truth.Values) { $w = @(([string]$i.Subject) -split '[^A-Za-z]+' | Where-Object { $_.Length -ge 7 }); if ($w.Count) { $word = $w[0]; break } }
if ($word) {
    $ev = Run $ScriptsDir "outlook-everything.ps1" @($word, "-Kind", "Mail", "-Days", "7", "-Top", "10")
    $mail = @($ev.groups.Mail | Where-Object { $_.EntryID })
    if ($mail.Count) {
        # The index covers every mailbox. Only Exchange mailboxes have a web address; IMAP/POP mail
        # (e.g. a Gmail account) gets an empty WebLink by design, but still its Message-ID.
        $exchange = @{}
        foreach ($a in (New-Object -ComObject Outlook.Application).Session.Accounts) { if ([int]$a.AccountType -eq 0) { $exchange[$a.DeliveryStore.DisplayName] = $true } }
        $isExchange = { $exchange.ContainsKey(([string]$_.Folder -split '/')[0]) }
        CheckRows "mail_everything (Exchange mailboxes)" @($mail | Where-Object $isExchange)
        $other = @($mail | Where-Object { -not (& $isExchange) })
        if ($other.Count) { Check "mail_everything (other mailboxes): no WebLink, Message-ID kept" (@($other | Where-Object { $_.WebLink -or -not $_.MessageID }).Count -eq 0) "($($other.Count) rows)" }
    }
    else { Write-Host "SKIP  mail_everything: the search index returned no openable mail for the test word" -ForegroundColor Yellow }
}

$text = Run $ScriptsDir "outlook-read.ps1" @("-Days", "7", "-Limit", "3") -Text
Check "text output shows WebLink and Message-ID lines" ($text -match "WebLink: https://" -and $text -match "Message-ID: <")

# ---------------------------------------------------------------------------
# Write results: draft, move, copy, forward, task from mail
# ---------------------------------------------------------------------------
$app = New-Object -ComObject Outlook.Application
$ns = $app.Session
$me = ""; try { $me = $ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { }
if (-not $me) { $me = $ns.Accounts.Item(1).SmtpAddress }
$deletedRoot = $ns.GetDefaultFolder(3)
$sandbox = $null
foreach ($f in $deletedRoot.Folders) { if ($f.Name -eq "_suite-test") { $sandbox = $f } }
if (-not $sandbox) { $sandbox = $deletedRoot.Folders.Add("_suite-test") }
$sandboxPath = "$($deletedRoot.Name)/_suite-test"
function LinkOf([string]$eid) { return (Run $ScriptsDir "outlook-read-body.ps1" @("-EntryID", $eid)).WebLink }

$tag = [guid]::NewGuid().ToString("N").Substring(0, 8)
$draft = Run $ScriptsDir "outlook-draft.ps1" @("-To", $me, "-Subject", "[suite-test] list links $tag", "-Body", "Test item for Test-ListLinks.ps1. Never sent.")
Check "mail_draft returns the draft's WebLink" ($draft.WebLink -match '^https://' -and $draft.WebLink -eq (LinkOf $draft.EntryID))

$movedText = Run $ScriptsDir "outlook-move.ps1" @("-EntryID", $draft.EntryID, "-Folder", $sandboxPath) -Text
$movedId = [regex]::Match($movedText, 'EntryID: ([0-9A-Fa-f]+)').Groups[1].Value
$moved = Run $ScriptsDir "outlook-read-body.ps1" @("-EntryID", $movedId)
Check "mail_move shows the new WebLink" ($moved.WebLink -match '^https://' -and $moved.WebLink -ne $draft.WebLink -and $movedText -match [regex]::Escape("WebLink: $($moved.WebLink)"))

$copy = Run $ScriptsDir "outlook-copy.ps1" @("-EntryID", $moved.EntryID, "-Folder", $sandboxPath)
Check "mail_copy returns the copy's own WebLink" ($copy.WebLink -match '^https://' -and $copy.WebLink -ne $moved.WebLink -and $copy.WebLink -eq (LinkOf $copy.EntryID))

$fwd = Run $ScriptsDir "outlook-forward.ps1" @("-EntryID", $moved.EntryID, "-To", $me, "-Body", "suite test, never sent")
Check "mail_forward (draft) returns the draft's WebLink" ($fwd.WebLink -match '^https://' -and $fwd.WebLink -eq (LinkOf $fwd.EntryID))
try { $null = $ns.GetItemFromID($fwd.EntryID).Move($sandbox) } catch { }

$task = Run $TasksDir "outlook-tasks-from-mail.ps1" @("-EntryID", $moved.EntryID, "-Force")
Check "tasks_from_mail returns the email's WebLink and Message-ID" ($task.Changes."Email WebLink" -eq $moved.WebLink -and [string]$task.Changes."Email Message-ID" -eq [string]$moved.MessageID)
$taskItem = $null; try { $taskItem = $ns.GetItemFromID($task.EntryID) } catch { }
Check "the task body has the web link" ($taskItem -and $taskItem.Body -match [regex]::Escape("Open it on the web: $($moved.WebLink)"))
$null = Run $TasksDir "outlook-tasks-delete.ps1" @("-EntryID", $task.EntryID, "-Confirm")

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
