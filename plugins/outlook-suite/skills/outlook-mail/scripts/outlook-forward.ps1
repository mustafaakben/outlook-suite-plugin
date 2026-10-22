param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [string]$To = "",
    [string]$CC = "",
    [string]$BCC = "",
    [string]$Body = "",
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$RemoveAutoSignature,
    [string]$DraftEntryID = "",
    [string[]]$Attachment = @(),
    [switch]$Send,
    [switch]$Confirm,
    [switch]$Display,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Forward: forward an email (with its attachments and the original message kept intact).
# Optional note at the top in Outlook's HTML style (-Body), or your own -HTML / -HtmlFile.
#
# Default: saves ONE forward draft and prints its EntryID (nothing is sent).
#   Update it: -DraftEntryID <draft EntryID> -Body "..." (an edit: previews unless -Confirm).
#   Send: -Send -Confirm (or -Confirm alone, as before). Outward-facing: previews unless confirmed.
#
# Usage: .\outlook-forward.ps1 -EntryID "00000000..." -To "colleague@x.com" -Body "FYI, see below."
# Usage: .\outlook-forward.ps1 -EntryID "00000000..." -To "colleague@x.com" -Send -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $wantSend = $Send -or ($Confirm -and -not $DraftEntryID -and -not $PSBoundParameters.ContainsKey("Send"))
    $original = ""

    if ($DraftEntryID) {
        $mail = Get-OutlookItem -Namespace $Namespace -EntryID $DraftEntryID
        if ($mail.Sent) { Stop-WithError "-DraftEntryID points to a sent item, not a draft." }
        $isNew = $false
    } else {
        if (-not $To) { Stop-WithError "Give the recipient(s) with -To." }
        $target = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
        $original = "$($target.Subject) ($($target.EntryID))"
        $mail = $target.Forward()
        $isNew = $true
    }

    if ($Body -or $HtmlFile) {
        if (-not $isNew -and -not ($Confirm -or -not (Test-ConfirmRequired -Action "edit-draft"))) {
            Write-ComposeSummary -Mail $mail -Title "PREVIEW (draft not changed)" -Json:$Json -Action "edit-draft"
            if (-not $Json) { Approve-Action -Action "edit-draft" -Confirmed $false -What "replacing the forward note" | Out-Null }
            exit 0
        }
        if ([int]$mail.BodyFormat -eq 2) {
            $content = Get-ComposeBody -Body $Body -HTML:$HTML -HtmlFile $HtmlFile -Native:(-not $HTML -and -not $HtmlFile)
            Set-ReplyBlock -Item $mail -Html $content.Html -RemoveAutoSignature:$RemoveAutoSignature
        } else {
            $mail.Body = $Body + "`r`n`r`n" + $mail.Body
        }
    } elseif ($RemoveAutoSignature -and [int]$mail.BodyFormat -eq 2) {
        $mail.HTMLBody = Remove-AutoSignature $mail.HTMLBody
    }

    if ($To) { foreach ($a in ($To -split ';')) { if ($a.Trim()) { $r = $mail.Recipients.Add($a.Trim()); $r.Type = 1 } } }
    if ($CC) { foreach ($a in ($CC -split ';')) { if ($a.Trim()) { $r = $mail.Recipients.Add($a.Trim()); $r.Type = 2 } } }
    if ($BCC) { foreach ($a in ($BCC -split ';')) { if ($a.Trim()) { $r = $mail.Recipients.Add($a.Trim()); $r.Type = 3 } } }
    foreach ($f in $Attachment) {
        if (-not $f) { continue }
        if (-not (Test-Path -LiteralPath $f)) { throw "Attachment not found: $f" }
        [void]$mail.Attachments.Add((Resolve-Path -LiteralPath $f).Path)
    }
    Assert-OutlookRecipients $mail

    if ($wantSend) {
        if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "send"))) {
            $mail.Save()
            Write-ComposeSummary -Mail $mail -Title "FORWARD DRAFT SAVED (not sent)" -Original $original -Json:$Json -Action "forward"
            if (-not $Json) { Approve-Action -Action "send" -Confirmed $false -What "sending" | Out-Null }
            exit 0
        }
        $summary = $mail
        $mail.Send()
        Write-ComposeSummary -Mail $summary -Title "FORWARD SENT" -Original $original -Json:$Json -Action "forward"
        exit 0
    }
    $mail.Save()
    if ($Display) { $mail.Display() }
    Write-ComposeSummary -Mail $mail -Title "FORWARD DRAFT SAVED" -Original $original -Json:$Json -Action "forward" `
        -Next "Change note: -DraftEntryID <EntryID> -Body ""..."" -Confirm. Send: -DraftEntryID <EntryID> -Send -Confirm"
}
catch {
    Stop-WithError $_
}
