<#
.SYNOPSIS
  Checks outlook-triage.ps1 against independent counts read item by item through the object
  model (not the table it uses), plus JSON, limits, errors, the host and Windows PowerShell 5.1.
  3.2: -ExcludeCategory, Message-ID, WebLink (decoded independently), -IncludeBody/-BodyChars,
  and read-body's Message-ID, WebLink and -MessageID lookup.
  Read-only. Prints counts only.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"))
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$Triage = Join-Path $ScriptsDir "outlook-triage.ps1"
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Triage([string[]]$a, [string]$exe = "pwsh") {
    $out = & $exe -NoProfile -File $Triage @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function TriageJson([string[]]$a, [string]$exe = "pwsh") { return ((Triage ($a + "-Json") $exe).Out | ConvertFrom-Json) }

# Independent truth: every item since a fixed start, read one by one.
$d3 = (Get-Date).Date.AddDays(-3); $d7 = (Get-Date).Date.AddDays(-7)
$app = New-Object -ComObject Outlook.Application
$inbox = $app.Session.GetDefaultFolder(6)
$verbTag = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
$scoreTag = "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}/CopilotInboxScore"
$midTag = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$truth = foreach ($it in $inbox.Items.Restrict("[ReceivedTime] >= '$($d7.ToString('MM/dd/yyyy HH:mm'))'")) {
    $verb = 0; try { $verb = [int]$it.PropertyAccessor.GetProperty($verbTag) } catch { }
    $score = 0; try { $score = [int]$it.PropertyAccessor.GetProperty($scoreTag) } catch { }
    $mid = ""; try { $mid = [string]$it.PropertyAccessor.GetProperty($midTag) } catch { }
    [pscustomobject]@{
        Received = $it.ReceivedTime; Unread = [bool]$it.UnRead
        Tags = @(([string]$it.Categories) -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        Replied = $verb -in 102, 103; Score = $score; EntryID = $it.EntryID; MessageID = $mid
    }
}
# Mailbox GUID of the default mailbox, read independently of _common.ps1 (empty if not Exchange).
$mailboxGuid = ""
foreach ($a in $app.Session.Accounts) {
    try { if ($a.DeliveryStore.StoreID -eq $inbox.StoreID -and [string]$a.ExchangeMailboxServerName -match 'MailboxId=([0-9a-fA-F-]{36})@') { $mailboxGuid = $Matches[1].ToLowerInvariant() } } catch { }
}
# Independent decoder for a WebLink: ItemID -> (GUID text, EntryID hex). Undoes the run-length step.
function Read-WebLink([string]$link) {
    if ($link -notmatch '^https://outlook\.office365\.com/owa/\?ItemID=([^&]+)&exvsurl=1&viewmodel=ReadMessageItem$') { return $null }
    $b = [Convert]::FromBase64String([uri]::UnescapeDataString($Matches[1]))
    $raw = New-Object System.Collections.Generic.List[byte]
    if ($b[0] -eq 1) {
        $i = 1
        while ($i -lt $b.Length) {
            $raw.Add($b[$i])
            if ($i + 1 -lt $b.Length -and $b[$i + 1] -eq $b[$i]) { $raw.Add($b[$i]); for ($k = 0; $k -lt $b[$i + 2]; $k++) { $raw.Add($b[$i]) }; $i += 3 } else { $i++ }
        }
    } else { for ($i = 1; $i -lt $b.Length; $i++) { $raw.Add($b[$i]) } }
    $r = $raw.ToArray()
    if ($r[0] -ne 3) { return $null }
    $gl = $r[1] + 256 * $r[2]; $guid = [Text.Encoding]::ASCII.GetString($r, 3, $gl)
    $p = 3 + $gl + 1; $el = $r[$p] + 256 * $r[$p + 1]
    $eid = [BitConverter]::ToString($r, $p + 2, $el).Replace("-", "")
    return [pscustomobject]@{ Guid = $guid; EntryID = $eid; Extra = $r.Length - ($p + 2 + $el) }
}
$in3 = @($truth | Where-Object { $_.Received -ge $d3 })
$from3 = $d3.ToString("yyyy-MM-dd"); $from7 = $d7.ToString("yyyy-MM-dd")
Write-Host "Independent count: $(@($truth).Count) items since $from7, $($in3.Count) since $from3" -ForegroundColor Gray

# Views against the independent counts
$want = @($in3 | Where-Object { $_.Unread -and $_.Tags.Count -eq 0 }).Count
$j = TriageJson @("untagged", "-DateFrom", $from3)
Check "untagged = unread with no tag since $from3" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

$want = @($in3 | Where-Object { $_.Tags.Count -gt 0 }).Count
$j = TriageJson @("tagged", "-DateFrom", $from3)
Check "tagged = any tag, read and unread" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

$want = @($truth | Where-Object { -not $_.Replied }).Count
$j = TriageJson @("reply", "-DateFrom", $from7, "-Limit", "1000")
Check "reply = not replied since $from7" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"
$dates = @($j.items | ForEach-Object { [datetime]$_.Date })
Check "reply is oldest first" (($dates.Count -lt 2) -or ($dates[0] -le $dates[-1])) "($($dates.Count) items)"

$want = @($truth | Where-Object { -not $_.Replied -and ($_.Tags | Where-Object { $_ -in "@Action", "@Review" }) }).Count
$j = TriageJson @("reply", "-DateFrom", $from7, "-Category", "@Action,@Review")
Check "reply -Category @Action,@Review" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

$want = @($truth | Where-Object { $_.Tags.Count -gt 0 -and $_.Score -eq 5 }).Count
$j = TriageJson @("tagged", "-DateFrom", $from7, "-Priority", "High")
Check "tagged -Priority High (CopilotInboxScore 5)" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

$want = @($truth | Where-Object { $_.Score -eq 1 }).Count
$j = TriageJson @("untagged", "-DateFrom", $from7, "-All", "-Priority", "Low")
$want = @($truth | Where-Object { $_.Tags.Count -eq 0 -and $_.Score -eq 1 }).Count
Check "untagged -All -Priority Low" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

$want = @($truth | Where-Object { $_.Replied -and -not $_.Unread }).Count
$j = TriageJson @("tagged", "-DateFrom", $from7, "-Replied", "-Read")
$want = @($truth | Where-Object { $_.Replied -and -not $_.Unread -and $_.Tags.Count -gt 0 }).Count
Check "tagged -Replied -Read" ($j.summary.matched -eq $want) "($($j.summary.matched)/$want)"

# Output contract
$j = TriageJson @("tagged", "-DateFrom", $from7, "-Limit", "2")
Check "-Limit 2 shows 2 of the matches" ($j.items.Count -eq [Math]::Min(2, $j.summary.matched)) "($($j.items.Count) of $($j.summary.matched))"
Check "every item has an EntryID" (@($j.items | Where-Object { $_.EntryID -notmatch '^[0-9A-F]{40,}$' }).Count -eq 0)
$t = Triage @("tagged", "-DateFrom", $from7)
Check "text view exits 0 with a summary line" ($t.Exit -eq 0 -and $t.Out -match "=== \d+ shown of \d+")
foreach ($bad in @(@("-Unread", "-Read"), @("-Priority", "Urgent"), @("-Folder", "No-Such-Folder-xyz"), @("-Untagged", "-Category", "@Action"),
        @("-Category", "@Action,@Review", "-ExcludeCategory", "@Review"), @("-IncludeBody", "-BodyChars", "-1"))) {
    $r = Triage $bad
    Check "error exits 1: $($bad -join ' ')" ($r.Exit -eq 1 -and $r.Out -match "^Error:")
}

# -ExcludeCategory against the independent counts: the two most used tags, one and both excluded
$top = @($truth | ForEach-Object { $_.Tags } | Group-Object | Sort-Object Count -Descending | Select-Object -First 2 | ForEach-Object { $_.Name })
if ($top.Count -gt 0) {
    foreach ($skip in @(, @($top[0])) + @(, $top)) {
        $want = @($truth | Where-Object { -not $_.Replied -and -not ($_.Tags | Where-Object { $_ -in $skip }) }).Count
        $j = TriageJson @("reply", "-DateFrom", $from7, "-Limit", "1000", "-ExcludeCategory", ($skip -join ","))
        $leak = @($j.items | Where-Object { $_.Categories | Where-Object { $_ -in $skip } }).Count
        Check "reply -ExcludeCategory ($($skip.Count) tag(s))" ($j.summary.matched -eq $want -and $leak -eq 0) "($($j.summary.matched)/$want, $leak excluded tags shown)"
    }
} else { Write-Host "SKIP  -ExcludeCategory: no tagged mail in the last 7 days" -ForegroundColor Yellow }

# Message-ID and WebLink on every row, against item-by-item reads
$j = TriageJson @("reply", "-DateFrom", $from7, "-Limit", "1000")
$byEid = @{}; foreach ($t in $truth) { $byEid[$t.EntryID] = $t }
$midBad = @($j.items | Where-Object { $byEid.ContainsKey($_.EntryID) -and $_.MessageID -ne $byEid[$_.EntryID].MessageID }).Count
Check "MessageID equals the item's Internet Message-ID" ($j.items.Count -gt 0 -and $midBad -eq 0) "($($j.items.Count - $midBad)/$($j.items.Count))"
if ($mailboxGuid) {
    $linkBad = 0
    foreach ($it in $j.items) {
        $d = Read-WebLink $it.WebLink
        if (-not $d -or $d.Guid -ne $mailboxGuid -or $d.EntryID -ne $it.EntryID -or $d.Extra -ne 0) { $linkBad++ }
    }
    Check "WebLink decodes to the mailbox GUID and the row's EntryID" ($linkBad -eq 0) "($($j.items.Count - $linkBad)/$($j.items.Count))"
} else { Write-Host "SKIP  WebLink decode: the default mailbox is not an Exchange account" -ForegroundColor Yellow }

# -IncludeBody: capped, flagged, and uncapped with -BodyChars 0
$j = TriageJson @("untagged", "-DateFrom", $from7, "-All", "-Limit", "5", "-IncludeBody", "-BodyChars", "300")
$capBad = @($j.items | Where-Object { $null -eq $_.Body -or ($_.BodyTruncated -and $_.Body -notmatch "\[\.\.\. cut at 300 of \d+ characters") -or (-not $_.BodyTruncated -and $_.Body.Length -gt 300) }).Count
Check "-IncludeBody -BodyChars 300: every body present, capped and flagged" ($j.items.Count -gt 0 -and $capBad -eq 0) "($($j.items.Count - $capBad)/$($j.items.Count))"
$j0 = TriageJson @("untagged", "-DateFrom", $from7, "-All", "-Limit", "5", "-IncludeBody", "-BodyChars", "0")
Check "-BodyChars 0: no body cut" (@($j0.items | Where-Object { $_.BodyTruncated }).Count -eq 0 -and @($j0.items | Where-Object { $null -eq $_.Body }).Count -eq 0)
$jn = TriageJson @("untagged", "-DateFrom", $from7, "-All", "-Limit", "5")
Check "no Body field without -IncludeBody" (@($jn.items | Where-Object { $_.PSObject.Properties.Name -contains "Body" }).Count -eq 0)

# read-body shows the same Message-ID and WebLink, and -MessageID finds the same email
$first = @($jn.items)[0]
if ($first) {
    $rb = & pwsh -NoProfile -File (Join-Path $ScriptsDir "outlook-read-body.ps1") -EntryID $first.EntryID -Json 2>&1 | Out-String | ConvertFrom-Json
    Check "read-body MessageID and WebLink equal triage's" ($rb.MessageID -eq $first.MessageID -and $rb.WebLink -eq $first.WebLink)
    if ($first.MessageID) {
        $rm = & pwsh -NoProfile -File (Join-Path $ScriptsDir "outlook-read-body.ps1") -MessageID $first.MessageID -Json 2>&1 | Out-String | ConvertFrom-Json
        Check "read-body -MessageID finds the same email" ($rm.EntryID -eq $first.EntryID -and $rm.WebLink -eq $first.WebLink)
    }
}

# Same EntryIDs through the warm host and under Windows PowerShell 5.1
foreach ($v in @(@("untagged", "-DateFrom", $from7, "-All"), @("tagged", "-DateFrom", $from3), @("reply", "-DateFrom", $from3))) {
    $key = { param($items) (@($items | ForEach-Object { "$($_.EntryID)|$($_.MessageID)|$($_.WebLink)" }) -join ",") }
    $a = & $key (TriageJson $v).items
    $h = & $key (& pwsh -NoProfile -File (Join-Path $ScriptsDir "outlook-call.ps1") triage @v -Json 2>&1 | Out-String | ConvertFrom-Json).items
    $p = & $key (TriageJson $v "powershell").items
    Check "same EntryIDs, Message-IDs and WebLinks direct / host / 5.1: $($v[0])" ($a -eq $h -and $a -eq $p) "($(@($a -split ',' | Where-Object { $_ }).Count) items)"
}

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
