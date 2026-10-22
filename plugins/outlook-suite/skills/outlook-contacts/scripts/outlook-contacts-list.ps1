param(
    [int]$Limit = 20,
    [string]$Search = "",
    [switch]$AllStores,
    [switch]$Json
)

# Outlook Contacts List: contacts sorted by name, read with one fast table query.
# Usage: .\outlook-contacts-list.ps1
# Usage: .\outlook-contacts-list.ps1 -Search "John" -Limit 50
# Usage: .\outlook-contacts-list.ps1 -AllStores -Json      # every mailbox's Contacts folder
# -Search matches name, emails, company, job title and department (case-insensitive).
# The [n] numbers are what -Index means in the other contact scripts (without -AllStores).

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

try {
    $null = Connect-Outlook
    $recs = Get-ContactRecords -Namespace $Namespace -AllStores:$AllStores -Search $Search
    $shown = @($recs | Select-Object -First $Limit)
    if ($Json) {
        Write-Json ([ordered]@{ Total = $recs.Count; Count = $shown.Count; Contacts = $shown })
        exit 0
    }
    $scope = if ($Search) { " matching '$Search'" } else { "" }
    Write-Host "=== CONTACTS ($($shown.Count) of $($recs.Count)$scope) ===" -ForegroundColor Cyan
    if ($recs.Count -eq 0) { Write-Host "No contacts found." -ForegroundColor Gray; exit 0 }
    $i = 0
    foreach ($r in $shown) { $i++; Write-ContactRecord $r -Number $i }
    if ($recs.Count -gt $shown.Count) { Write-Host "($($recs.Count - $shown.Count) more; raise -Limit)" -ForegroundColor DarkGray }
}
catch {
    Stop-WithError $_
}
