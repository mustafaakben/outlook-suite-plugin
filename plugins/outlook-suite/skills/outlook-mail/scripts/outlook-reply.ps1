param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [string]$Body = "",
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$ReplyAll,
    [switch]$RemoveAutoSignature,
    [string]$DraftEntryID = "",
    [string]$AddTo = "",
    [string]$AddCC = "",
    [string]$AddBCC = "",
    [string[]]$Attachment = @(),
    [switch]$Send,
    [switch]$Confirm,
    [switch]$Display,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Reply: reply (or -ReplyAll) inside the original thread. The quoted thread is kept
# intact; your text goes at the top in Outlook's own HTML style (Aptos 11pt div blocks), or
# pass your own HTML with -HTML / -HtmlFile.
#
# Default: saves ONE reply draft and prints its EntryID (nothing is sent).
#   Update that draft instead of making another: -DraftEntryID <draft EntryID> -Body "..."
#   (replacing the text is an edit: previews unless -Confirm).
#   Send: -Send -Confirm (or -Confirm alone, as before). Outward-facing: previews unless confirmed.
# -RemoveAutoSignature drops the signature Outlook inserts automatically into replies
# (useful when your text already ends with your own signature).
#
# Usage: .\outlook-reply.ps1 -EntryID "00000000..." -Body "Thanks, see you Friday." -ReplyAll
# Usage: .\outlook-reply.ps1 -EntryID "00000000..." -HtmlFile .\reply.html -ReplyAll -RemoveAutoSignature
# Usage: .\outlook-reply.ps1 -DraftEntryID "<draft>" -Body "New wording" -Confirm
# Usage: .\outlook-reply.ps1 -DraftEntryID "<draft>" -Send -Confirm

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
        $target = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
        $original = "$($target.Subject) ($($target.EntryID))"
        $mail = if ($ReplyAll) { $target.ReplyAll() } else { $target.Reply() }
        $isNew = $true
        if (-not $Body -and -not $HtmlFile) { Stop-WithError "Give the reply text with -Body (or -HtmlFile)." }
    }

    if ($Body -or $HtmlFile) {
        if (-not $isNew) {
            $editOk = $Confirm -or -not (Test-ConfirmRequired -Action "edit-draft")
            if (-not $editOk) {
                Write-ComposeSummary -Mail $mail -Title "PREVIEW (draft not changed)" -Json:$Json -Action "edit-draft"
                if (-not $Json) { Approve-Action -Action "edit-draft" -Confirmed $false -What "replacing the reply text" | Out-Null }
                exit 0
            }
        }
        if ([int]$mail.BodyFormat -eq 2) {
            $content = Get-ComposeBody -Body $Body -HTML:$HTML -HtmlFile $HtmlFile -Native:(-not $HTML -and -not $HtmlFile)
            Set-ReplyBlock -Item $mail -Html $content.Html -RemoveAutoSignature:$RemoveAutoSignature
        } else {
            if (-not $isNew) { Stop-WithError "This draft is plain text; rewrite it with outlook-draft-update.ps1 -Body." }
            $mail.Body = $Body + "`r`n`r`n" + $mail.Body
        }
    }

    foreach ($pair in @(@($AddTo, 1), @($AddCC, 2), @($AddBCC, 3))) {
        foreach ($addr in ($pair[0] -split ';')) { if ($addr.Trim()) { $r = $mail.Recipients.Add($addr.Trim()); $r.Type = $pair[1] } }
    }
    foreach ($f in $Attachment) {
        if (-not $f) { continue }
        if (-not (Test-Path -LiteralPath $f)) { throw "Attachment not found: $f" }
        [void]$mail.Attachments.Add((Resolve-Path -LiteralPath $f).Path)
    }
    Assert-OutlookRecipients $mail

    $kind = if ($ReplyAll) { "REPLY ALL" } else { "REPLY" }
    if ($wantSend) {
        if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "send"))) {
            $mail.Save()
            Write-ComposeSummary -Mail $mail -Title "$kind DRAFT SAVED (not sent)" -Original $original -Json:$Json -Action "reply"
            if (-not $Json) { Approve-Action -Action "send" -Confirmed $false -What "sending" | Out-Null }
            exit 0
        }
        $summary = $mail
        $mail.Send()
        Write-ComposeSummary -Mail $summary -Title "$kind SENT" -Original $original -Json:$Json -Action "reply"
        exit 0
    }
    $mail.Save()
    if ($Display) { $mail.Display() }
    Write-ComposeSummary -Mail $mail -Title "$kind DRAFT SAVED" -Original $original -Json:$Json -Action "reply" `
        -Next "Change text: -DraftEntryID <EntryID> -Body ""..."" -Confirm. Send: -DraftEntryID <EntryID> -Send -Confirm"
}
catch {
    Stop-WithError $_
}
