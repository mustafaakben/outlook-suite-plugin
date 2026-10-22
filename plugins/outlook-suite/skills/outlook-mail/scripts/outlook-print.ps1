param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [Alias("ConfirmPrint")]
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Print: send an email to the default printer.
# Outward-facing action: previews unless -Confirm (alias -ConfirmPrint).
# Usage: .\outlook-print.ps1 -EntryID "00000000..." -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    if (-not (Confirm-ItemAction -Action "print" -Title "PRINT EMAIL" -Item $email -Confirmed $Confirm -What "printing")) { exit 0 }
    $email.PrintOut()
    Write-ActionResult -Title "SENT TO DEFAULT PRINTER" -Item $email -Json:$Json -Action "print"
}
catch {
    Stop-WithError $_
}
