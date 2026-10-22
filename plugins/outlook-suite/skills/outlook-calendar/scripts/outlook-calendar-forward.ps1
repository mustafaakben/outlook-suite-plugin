param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [datetime]$OccurrenceStart,

    [Parameter(Mandatory=$true)]
    [string]$To,

    [string]$Body = "",
    [int]$Days = 7,
    [switch]$Send,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Calendar Forward: send an event as a calendar attachment (.vcs) by email.
# Default saves a draft. -Send sends it (irreversible: previews unless -Confirm).
# Usage: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "colleague@example.com"
# With message: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "a@example.com" -Body "FYI, the meeting details"
# Send now: .\outlook-calendar-forward.ps1 -EntryID "00000000..." -To "a@example.com" -Send -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days $Days -OccurrenceStart $occ
    $rec = ConvertTo-EventRecord $item

    if ($Send -and (Confirm-CalendarAction -Action "send" -Title "FORWARD EVENT" -Item $item -Confirmed $Confirm -Details ([ordered]@{ To = $To }) -What "the forward" -Json:$Json) -ne $true) { Write-CalendarPreview; exit 0 }

    $mail = $item.ForwardAsVcal()
    $mail.To = $To
    if ($Body) { $mail.Body = $Body + "`r`n`r`n" + $mail.Body }
    Assert-OutlookRecipients $mail
    if ($Send) { $mail.Send(); $result = "sent" } else { $mail.Save(); $result = "draft saved" }

    if ($Json) {
        $o = ConvertTo-EventJson $rec; $o.Action = "forward"; $o.Result = $result; $o.To = $To
        if (-not $Send) { $o.DraftEntryID = [string]$mail.EntryID }
        Write-Json $o; exit 0
    }
    Write-Host "=== EVENT FORWARDED ($result) ===" -ForegroundColor Green
    Write-EventRecord $rec -Prefix "  "
    Write-Host "  To: $To" -ForegroundColor Cyan
    if (-not $Send) { Write-Host "  Draft EntryID: $($mail.EntryID)" -ForegroundColor DarkGray; Write-Host "  Review it in Drafts, or rerun with -Send -Confirm." -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
