<#
.SYNOPSIS
  Checks outlook-triage.ps1 against independent counts read item by item through the object
  model (not the table it uses), plus JSON, limits, errors, the host and Windows PowerShell 5.1.
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
$truth = foreach ($it in $inbox.Items.Restrict("[ReceivedTime] >= '$($d7.ToString('MM/dd/yyyy HH:mm'))'")) {
    $verb = 0; try { $verb = [int]$it.PropertyAccessor.GetProperty($verbTag) } catch { }
    $score = 0; try { $score = [int]$it.PropertyAccessor.GetProperty($scoreTag) } catch { }
    [pscustomobject]@{
        Received = $it.ReceivedTime; Unread = [bool]$it.UnRead
        Tags = @(([string]$it.Categories) -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        Replied = $verb -in 102, 103; Score = $score
    }
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
foreach ($bad in @(@("-Unread", "-Read"), @("-Priority", "Urgent"), @("-Folder", "No-Such-Folder-xyz"), @("-Untagged", "-Category", "@Action"))) {
    $r = Triage $bad
    Check "error exits 1: $($bad -join ' ')" ($r.Exit -eq 1 -and $r.Out -match "^Error:")
}

# Same EntryIDs through the warm host and under Windows PowerShell 5.1
foreach ($v in @(@("untagged", "-DateFrom", $from7, "-All"), @("tagged", "-DateFrom", $from3), @("reply", "-DateFrom", $from3))) {
    $a = (TriageJson $v).items.EntryID -join ","
    $h = (& pwsh -NoProfile -File (Join-Path $ScriptsDir "outlook-call.ps1") triage @v -Json 2>&1 | Out-String | ConvertFrom-Json).items.EntryID -join ","
    $p = (TriageJson $v "powershell").items.EntryID -join ","
    Check "same EntryIDs direct / host / 5.1: $($v[0])" ($a -eq $h -and $a -eq $p) "($(@($a -split ',' | Where-Object { $_ }).Count) items)"
}

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
