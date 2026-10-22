param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [Parameter(Mandatory = $true)]
    [ValidateSet("Status", "MarkRead", "MarkUnread", "AlwaysMove", "StopAlwaysMove", "AlwaysDelete", "StopAlwaysDelete", "AlwaysCategorize", "StopAlwaysCategorize")]
    [string]$Action,
    [string]$Folder = "",
    [string]$Category = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Conversation Actions: act on a whole thread (the conversation an email belongs to).
#   Status                 what is set for this thread (always move / delete / categories)
#   MarkRead / MarkUnread  every message in the thread
#   AlwaysMove -Folder X   move the thread now and every future message in it  (StopAlwaysMove)
#   AlwaysDelete           delete the thread now and every future message ("ignore") (StopAlwaysDelete)
#   AlwaysCategorize -Category "A;B"   tag the thread now and future messages (StopAlwaysCategorize)
# The Always* actions are gated: preview unless -Confirm.
# Usage: .\outlook-conversation-actions.ps1 -EntryID "00000000..." -Action Status
# Usage: .\outlook-conversation-actions.ps1 -EntryID "00000000..." -Action AlwaysMove -Folder "Archive/Newsletters" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    $conv = $null; try { $conv = $email.GetConversation() } catch { }
    if (-not $conv) { Stop-WithError "This email has no conversation (conversations are off for this mailbox, or it is a draft)." }
    $store = $email.Parent.Store
    $count = 0; try { $count = $conv.GetTable().GetRowCount() } catch { }

    $status = [ordered]@{ Topic = $email.ConversationTopic; ConversationID = $conv.ConversationID; Messages = $count }
    try { $mv = $conv.GetAlwaysMoveToFolder($store); $status.AlwaysMoveTo = $(if ($mv) { $mv.FolderPath } else { "" }) } catch { }
    try { $status.AlwaysDelete = @{ 0 = "no"; 1 = "yes"; 2 = "not supported" }[[int]$conv.GetAlwaysDelete($store)] } catch { }
    try { $status.AlwaysCategories = [string]$conv.GetAlwaysAssignCategories($store) } catch { }

    if ($Action -eq "Status") {
        if ($Json) { Write-Json $status; exit 0 }
        Write-Host "=== CONVERSATION ===" -ForegroundColor Cyan
        foreach ($k in $status.Keys) { Write-Host "  ${k}: $($status[$k])" }
        exit 0
    }

    if ($Action -in "MarkRead", "MarkUnread") {
        if ($Action -eq "MarkRead") { $conv.MarkAsRead() } else { $conv.MarkAsUnread() }
        $status.Done = $Action
        if ($Json) { Write-Json $status; exit 0 }
        Write-Host "=== THREAD $(if ($Action -eq 'MarkRead') { 'MARKED READ' } else { 'MARKED UNREAD' }) ($count messages) ===" -ForegroundColor Green
        Write-Host "  Topic: $($status.Topic)"
        exit 0
    }

    $target = $null
    if ($Action -eq "AlwaysMove") {
        if (-not $Folder) { Stop-WithError "AlwaysMove needs -Folder." }
        $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    }
    if ($Action -eq "AlwaysCategorize" -and -not $Category) { Stop-WithError "AlwaysCategorize needs -Category." }
    $what = switch ($Action) {
        "AlwaysMove" { "move this thread ($count messages) and all future messages in it to $($target.FolderPath)" }
        "StopAlwaysMove" { "stop moving this thread automatically" }
        "AlwaysDelete" { "delete this thread ($count messages) and all future messages in it" }
        "StopAlwaysDelete" { "stop deleting this thread automatically" }
        "AlwaysCategorize" { "tag this thread and all future messages in it with '$Category'" }
        "StopAlwaysCategorize" { "stop tagging this thread automatically" }
    }
    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "conversation"))) {
        Write-Host "=== CONVERSATION ACTION (PREVIEW) ===" -ForegroundColor Yellow
        foreach ($k in $status.Keys) { Write-Host "  ${k}: $($status[$k])" }
        Write-Host "  Would: $what" -ForegroundColor Cyan
        Approve-Action -Action "conversation" -Confirmed $false -What "this conversation action" | Out-Null
        exit 0
    }
    switch ($Action) {
        "AlwaysMove" { $conv.SetAlwaysMoveToFolder($target, $store) }
        "StopAlwaysMove" { $conv.StopAlwaysMoveToFolder($store) }
        "AlwaysDelete" { $conv.SetAlwaysDelete(1, $store) }
        "StopAlwaysDelete" { $conv.StopAlwaysDelete($store) }
        "AlwaysCategorize" { $conv.SetAlwaysAssignCategories((@($Category -split ';' | ForEach-Object { $_.Trim() }) -join ", "), $store) }
        "StopAlwaysCategorize" { $conv.ClearAlwaysAssignCategories($store) }
    }
    $status.Done = $what
    if ($Json) { Write-Json $status; exit 0 }
    Write-Host "=== CONVERSATION ACTION DONE ===" -ForegroundColor Green
    Write-Host "  $what"
}
catch {
    Stop-WithError $_
}
