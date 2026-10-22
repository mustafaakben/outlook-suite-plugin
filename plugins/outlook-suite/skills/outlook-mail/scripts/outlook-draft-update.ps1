param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$To = "",
    [string]$CC = "",
    [string]$BCC = "",
    [string]$AddTo = "",
    [string]$AddCC = "",
    [string]$AddBCC = "",
    [string]$Subject = "",
    [string]$Body = "",
    [switch]$AppendBody,
    [switch]$HTML,
    [string]$HtmlFile = "",
    [switch]$Native,
    [string[]]$AddAttachment = @(),
    [string[]]$RemoveAttachment = @(),
    [string]$Account = "",
    [ValidateSet("", "Low", "Normal", "High")]
    [string]$Importance = "",
    [ValidateSet("", "Normal", "Personal", "Private", "Confidential")]
    [string]$Sensitivity = "",
    [int]$Limit = 20,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Draft Update: change an existing draft (new, reply or forward).
#   Replace recipients (-To/-CC/-BCC) or add (-AddTo/-AddCC/-AddBCC); -Subject; -Account;
#   -Importance; -Sensitivity; -AddAttachment.
#   Body: -Body replaces the text you wrote (in a reply/forward only your block above the
#   quoted thread is replaced; the thread is kept). -AppendBody adds to the end instead.
#   Replacing text or removing attachments (-RemoveAttachment name or number) are edits:
#   they preview unless -Confirm. Appending and adding never need confirmation.
# -Index counts Drafts by last change, newest first (volatile; prefer -EntryID).
# Usage: .\outlook-draft-update.ps1 -EntryID "00000000..." -Subject "New subject" -AddAttachment .\a.pdf
# Usage: .\outlook-draft-update.ps1 -EntryID "00000000..." -Body "Rewritten text" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $mail = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days 3650 -Folder "Drafts" -DateProperty "LastModificationTime"
    if ($mail.Sent) { Stop-WithError "This item was already sent; only drafts can be updated." }

    $isEdit = ($Body -or $HtmlFile) -and -not $AppendBody
    if ($RemoveAttachment.Count -gt 0) { $isEdit = $true }
    if ($isEdit -and -not ($Confirm -or -not (Test-ConfirmRequired -Action "edit-draft"))) {
        Write-ComposeSummary -Mail $mail -Title "PREVIEW (draft not changed)" -Json:$Json -Action "edit-draft"
        if (-not $Json) { Approve-Action -Action "edit-draft" -Confirmed $false -What "this edit (replacing text or removing attachments)" | Out-Null }
        exit 0
    }

    $changes = @()
    $fields = @{}
    foreach ($k in "To", "CC", "BCC", "Subject", "Account", "Importance", "Sensitivity") {
        if ($PSBoundParameters.ContainsKey($k) -and $PSBoundParameters[$k]) { $fields[$k] = $PSBoundParameters[$k]; $changes += $k }
    }
    if ($AddAttachment.Count) { $fields.Attachment = $AddAttachment; $changes += "attachments added" }
    Set-MailFields -Mail $mail @fields
    foreach ($pair in @(@($AddTo, 1), @($AddCC, 2), @($AddBCC, 3))) {
        foreach ($addr in ($pair[0] -split ';')) { if ($addr.Trim()) { $r = $mail.Recipients.Add($addr.Trim()); $r.Type = $pair[1]; $changes += "recipient added" } }
    }

    foreach ($ra in $RemoveAttachment) {
        $removed = $false
        for ($i = $mail.Attachments.Count; $i -ge 1; $i--) {
            $a = $mail.Attachments.Item($i)
            if ("$i" -eq "$ra" -or $a.FileName -ieq $ra) { $a.Delete(); $removed = $true; $changes += "removed attachment $($a.FileName)"; break }
        }
        if (-not $removed) { throw "No attachment '$ra' on this draft." }
    }

    if ($Body -or $HtmlFile) {
        $isHtml = [int]$mail.BodyFormat -eq 2
        if ($isHtml) {
            $content = Get-ComposeBody -Body $Body -HTML:$HTML -HtmlFile $HtmlFile -Native:(-not $HTML -and -not $HtmlFile)
            $h = $mail.HTMLBody
            $hasBlock = $h.Contains("<!--suite-reply-start-->")
            if ($AppendBody) {
                if ($hasBlock) { $h = $h.Replace("</div><!--suite-reply-end-->", "$($content.Html)</div><!--suite-reply-end-->") }
                else { $m = [regex]::Match($h, '</body>'); $h = if ($m.Success) { $h.Insert($m.Index, $content.Html) } else { $h + $content.Html } }
                $mail.HTMLBody = $h
                $changes += "body appended"
            } elseif ($hasBlock -or $mail.ConversationIndex.Length -gt 44) {
                # Reply/forward draft: replace only our block above the thread (or insert it).
                Set-ReplyBlock -Item $mail -Html $content.Html
                $changes += "reply text replaced"
            } else {
                $mail.HTMLBody = $content.Html
                $changes += "body replaced"
            }
        } else {
            $mail.Body = if ($AppendBody) { $mail.Body + "`r`n" + $Body } else { $Body }
            $changes += $(if ($AppendBody) { "body appended" } else { "body replaced" })
        }
    }

    Assert-OutlookRecipients $mail
    $mail.Save()
    Write-ComposeSummary -Mail $mail -Title "DRAFT UPDATED ($($changes -join ', '))" -Json:$Json -Action "draft-update" `
        -Next "Send: outlook-send-draft.ps1 -EntryID <EntryID> -Confirm"
}
catch {
    Stop-WithError $_
}
