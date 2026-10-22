<#
.SYNOPSIS
  Live test for the compose scripts (draft, draft-update, reply, forward, send gates). Sends NOTHING.

.DESCRIPTION
  Every message is addressed to your own mailbox and saved as a draft; replies and forwards
  are made from one of your own Sent Items. Send scripts are exercised in preview mode only.
  Every test body contains the marker "suite-test-marker", and everything carrying it is
  deleted (and purged from Deleted Items) at the end. Prints PASS/FAIL lines only.
#>
param(
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$null = Connect-Outlook

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $out = & pwsh -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}
function Item($id) { return $Namespace.GetItemFromID($id) }
function JsonOf($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }

$me = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
$marker = "suite-test-marker"
$tmpDir = Join-Path $env:TEMP "suite-compose-test"; New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
$f1 = Join-Path $tmpDir "attach-one.txt"; "one" | Set-Content $f1
$f2 = Join-Path $tmpDir "attach-two.txt"; "two" | Set-Content $f2
$sentBefore = $Namespace.GetDefaultFolder(5).Items.Count

try {
    # --- New draft ---
    $r = Run "outlook-draft.ps1" @("-To", $me, "-Subject", "[suite-test] compose", "-Body", "Hello $marker`nSecond line https://example.com/x", "-Native", "-Attachment", $f1, "-Json")
    $d = JsonOf $r; $it = Item $d.EntryID
    Check "draft -Native with attachment" ($r.Exit -eq 0 -and $it.HTMLBody -match "elementToProof" -and $it.Attachments.Count -eq 1 -and -not $it.Sent)
    $draftId = $d.EntryID

    # --- draft-update ---
    $r = Run "outlook-draft-update.ps1" @("-EntryID", $draftId, "-Subject", "[suite-test] compose v2", "-AddAttachment", $f2)
    $it = Item $draftId; Check "draft-update subject + add attachment (no confirm needed)" ($r.Exit -eq 0 -and $it.Subject -eq "[suite-test] compose v2" -and $it.Attachments.Count -eq 2)
    $r = Run "outlook-draft-update.ps1" @("-EntryID", $draftId, "-RemoveAttachment", "attach-one.txt")
    $it = Item $draftId; Check "draft-update remove attachment previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and $it.Attachments.Count -eq 2)
    $r = Run "outlook-draft-update.ps1" @("-EntryID", $draftId, "-RemoveAttachment", "attach-one.txt", "-Confirm")
    $it = Item $draftId; Check "draft-update remove attachment -Confirm" ($r.Exit -eq 0 -and $it.Attachments.Count -eq 1)
    $r = Run "outlook-draft-update.ps1" @("-EntryID", $draftId, "-Body", "Replaced $marker", "-Confirm")
    $it = Item $draftId; Check "draft-update body replace -Confirm" ($r.Exit -eq 0 -and $it.Body -match "Replaced" -and $it.Body -notmatch "Second line")
    $r = Run "outlook-draft-update.ps1" @("-EntryID", $draftId, "-Body", "Appended line", "-AppendBody")
    $it = Item $draftId; Check "draft-update -AppendBody (no confirm needed)" ($r.Exit -eq 0 -and $it.Body -match "Replaced" -and $it.Body -match "Appended line")

    # --- Send gates (preview only) ---
    $r = Run "outlook-send.ps1" @("-To", $me, "-Subject", "[suite-test] send preview", "-Body", $marker)
    Check "send previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")
    $r = Run "outlook-send.ps1" @("-To", "zzq-no-such-person-xq", "-Subject", "[suite-test] bad", "-Body", $marker, "-Confirm")
    Check "send to unknown recipient fails before sending" ($r.Exit -eq 1 -and $r.Out -match "Unresolved")
    $r = Run "outlook-send-draft.ps1" @("-EntryID", $draftId)
    Check "send-draft previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and -not (Item $draftId).Sent)
    $r = Run "outlook-schedule-send.ps1" @("-To", $me, "-Subject", "[suite-test] past", "-Body", $marker, "-SendAt", (Get-Date).AddHours(-1).ToString("yyyy-MM-dd HH:mm"), "-Confirm")
    Check "schedule-send rejects a past time" ($r.Exit -eq 1)
    $r = Run "outlook-voting.ps1" @("-To", $me, "-Subject", "[suite-test] vote", "-Body", $marker, "-Options", "A;B")
    Check "voting previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW")

    # --- Reply in thread (from one of my own Sent Items; Reply addresses me) ---
    $sentItems = $Namespace.GetDefaultFolder(5).Items; $sentItems.Sort("[SentOn]", $true)
    $src = $null; foreach ($s in $sentItems) { if ($s.Class -eq 43 -and [int]$s.BodyFormat -eq 2) { $src = $s; break } }
    $r = Run "outlook-reply.ps1" @("-EntryID", $src.EntryID, "-Body", "Reply text $marker", "-Json")
    $rep = JsonOf $r; $h = (Item $rep.EntryID).HTMLBody
    $blk = $h.IndexOf("suite-reply-start"); $orig = $h.IndexOf("_MailOriginal")
    Check "reply draft: block above quoted thread" ($r.Exit -eq 0 -and $blk -ge 0 -and ($orig -lt 0 -or $blk -lt $orig) -and $h -match "Reply text")
    Check "reply draft: quoted thread kept" ($orig -gt 0 -or $h.Length -gt 2000)
    $replyId = $rep.EntryID

    $r = Run "outlook-reply.ps1" @("-DraftEntryID", $replyId, "-Body", "Second wording $marker")
    Check "reply update previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and (Item $replyId).HTMLBody -match "Reply text")
    $r = Run "outlook-reply.ps1" @("-DraftEntryID", $replyId, "-Body", "Second wording $marker", "-Confirm")
    $h = (Item $replyId).HTMLBody
    Check "reply update -Confirm replaces the block (one block, old text gone)" ($r.Exit -eq 0 -and $h -match "Second wording" -and $h -notmatch "Reply text" -and ([regex]::Matches($h, "suite-reply-start")).Count -eq 1)
    $r = Run "outlook-reply.ps1" @("-DraftEntryID", $replyId, "-Send")
    Check "reply -Send without -Confirm does not send" ($r.Exit -eq 0 -and $r.Out -match "not sent" -and -not (Item $replyId).Sent)

    $r = Run "outlook-reply.ps1" @("-EntryID", $src.EntryID, "-Body", "No sig $marker", "-RemoveAutoSignature", "-Json")
    $h2 = (Item (JsonOf $r).EntryID).HTMLBody
    Check "reply -RemoveAutoSignature drops _MailAutoSig, keeps thread" ($r.Exit -eq 0 -and $h2 -notmatch "_MailAutoSig" -and $h2 -match "No sig")

    # --- Forward ---
    $r = Run "outlook-forward.ps1" @("-EntryID", $src.EntryID, "-To", $me, "-Body", "FYI $marker", "-Json")
    $fw = JsonOf $r; $fi = Item $fw.EntryID
    Check "forward draft: note on top, attachments carried" ($r.Exit -eq 0 -and $fi.HTMLBody -match "FYI" -and $fi.Attachments.Count -ge $src.Attachments.Count -and $fi.To -ne "")

    Check "nothing was sent during the test" ($Namespace.GetDefaultFolder(5).Items.Count -eq $sentBefore)
}
finally {
    Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
    $q = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%$marker%' OR ""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    foreach ($pass in 1, 2) {
        foreach ($f in @($Namespace.GetDefaultFolder(16), $Namespace.GetDefaultFolder(3), $Namespace.GetDefaultFolder(4))) {
            try { foreach ($i in @($f.Items.Restrict($q))) { if (-not $i.Sent) { try { $i.Delete() } catch { } } } } catch { }
        }
    }
}

Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
