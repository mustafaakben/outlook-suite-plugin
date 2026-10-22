param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [switch]$OpenRecallDialog,
    [switch]$DeleteUnread,
    [Alias("ConfirmDeleteFromSent")]
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Recall: help recall a sent email (Exchange / Microsoft 365 only).
# Outlook's object model has no recall method, so there are two routes:
#   -OpenRecallDialog  opens the message and Outlook's own "Recall This Message" dialog on
#                      this PC; the user chooses the options and clicks OK.
#   -DeleteUnread      deletes your copy from Sent Items only (NOT a recipient recall);
#                      irreversible, previews unless -Confirm.
# With neither switch it shows the message and the manual steps.
# -Index counts Sent Items (newest first).
# Usage: .\outlook-recall.ps1 -EntryID "00000000..." -OpenRecallDialog

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days -Folder "Sent" -DateProperty "SentOn"

    if ($DeleteUnread) {
        if (-not (Confirm-ItemAction -Action "delete" -Title "DELETE FROM SENT ITEMS (not a recall)" -Item $email -Confirmed $Confirm -What "the delete from Sent Items")) { exit 0 }
        $rec = ConvertTo-MailRecord $email
        $email.Delete()
        if ($Json) { Write-Json ([ordered]@{ Action = "recall"; Result = "deleted from Sent Items"; EntryID = $rec.EntryID; Subject = $rec.Subject }); exit 0 }
        Write-Host "=== REMOVED FROM SENT ITEMS ===" -ForegroundColor Green
        Write-Host "  Subject: $($rec.Subject)" -ForegroundColor Yellow
        Write-Host "  No recall was sent to recipients." -ForegroundColor Gray
        exit 0
    }

    if ($OpenRecallDialog) {
        $email.Display()
        $insp = $email.GetInspector
        $insp.CommandBars.ExecuteMso("RecallThisMessage")
        Write-ActionResult -Title "RECALL DIALOG OPENED" -Item $email -Changes ([ordered]@{ Next = "Choose the recall options in the Outlook window on this PC and click OK." }) -Json:$Json -Action "recall"
        exit 0
    }

    Write-ActionResult -Title "RECALL (manual)" -Item $email -Changes ([ordered]@{
        Steps = "Open the message in Sent Items > Message > Actions (or ...) > Recall This Message"
        Shortcut = "Run again with -OpenRecallDialog to open that dialog for this message"
        Limits = "Exchange/M365 recipients only; fails if already read or moved; recipients may see the recall notice"
    }) -Json:$Json -Action "recall"
}
catch {
    Stop-WithError $_
}
