param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [string]$Name = "",
    [string]$Search = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Contact Delete: move one contact to Deleted Items.
# Irreversible action: previews unless -Confirm (or confirmation for "delete" is turned off;
# see the mail SKILL.md "Confirmation").
# Usage: .\outlook-contacts-delete.ps1 -EntryID "00000000..." -Confirm
# Usage: .\outlook-contacts-delete.ps1 -Name "John Doe" -Confirm     # exact full name, must be unique

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

try {
    $null = Connect-Outlook
    if ($Name -and -not $EntryID -and $Index -le 0) {
        $all = Get-ContactRecords -Namespace $Namespace
        $hits = @($all | Where-Object { $_.FullName -eq $Name })
        if ($hits.Count -eq 0) { Stop-WithError "No contact named exactly '$Name'." -Hint "Find it with outlook-contacts-search.ps1 -Name ""$Name"" and delete by -EntryID." }
        if ($hits.Count -gt 1) { Stop-WithError "$($hits.Count) contacts are named '$Name'." -Hint "Delete by -EntryID (see outlook-contacts-search.ps1 -Name ""$Name"")." }
        $EntryID = $hits[0].EntryID
    }
    $contact = Get-ContactTarget -Namespace $Namespace -EntryID $EntryID -Index $Index -Search $Search
    $rec = ConvertTo-ContactRecord $contact

    if (-not $Json -and -not $Confirm -and (Test-ConfirmRequired -Action "delete")) {
        Write-Host "=== DELETE CONTACT (PREVIEW) ===" -ForegroundColor Yellow
        Write-ContactRecord $rec
    }
    if ((Approve-ContactAction -Action "delete" -Confirmed $Confirm -What "the contact delete" -Json:$Json) -ne $true) {   # fail closed
        if ($Json) {
            $o = [ordered]@{ Action = "delete"; Result = "preview"; Confirmed = $false }
            foreach ($k in $rec.Keys) { $o[$k] = $rec[$k] }
            Write-Json $o
        }
        exit 0
    }
    $contact.Delete()
    Write-ContactResult -Title "CONTACT DELETED (moved to Deleted Items)" -Rec $rec -Action "delete" -Json:$Json
}
catch {
    Stop-WithError $_
}
