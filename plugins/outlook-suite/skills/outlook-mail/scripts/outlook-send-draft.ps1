param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Send Draft: send an existing draft (for example one made by outlook-draft.ps1 or
# outlook-reply.ps1 and reviewed in Outlook). Recipients are resolved first.
# Outward-facing: previews unless -Confirm.
# Usage: .\outlook-send-draft.ps1 -EntryID "00000000..." -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $mail = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    if ($mail.Sent) { Stop-WithError "This item was already sent; it is not a draft." }
    Assert-OutlookRecipients $mail
    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "send"))) {
        Write-ComposeSummary -Mail $mail -Title "PREVIEW (draft not sent)" -Json:$Json -Action "send"
        if (-not $Json) { Approve-Action -Action "send" -Confirmed $false -What "sending" | Out-Null }
        exit 0
    }
    $summary = $mail
    $mail.Send()
    Write-ComposeSummary -Mail $summary -Title "SENT" -Json:$Json -Action "send"
}
catch {
    Stop-WithError $_
}
