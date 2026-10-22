param(
    [string]$Name = "",
    [string]$Email = "",
    [string]$Company = "",
    [int]$Limit = 10,
    [switch]$AllStores,
    [switch]$Json
)

# Outlook Contact Search: your Contacts folder by name, email and/or company (AND, contains).
# Usage: .\outlook-contacts-search.ps1 -Name "John"
# Usage: .\outlook-contacts-search.ps1 -Email "example.com" -Company "Acme" -Json
# To find anyone (your organization's directory and recent mail too), use outlook-people.ps1.

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

if (-not $Name -and -not $Email -and -not $Company) {
    Stop-WithError "Give at least one of -Name, -Email or -Company." -Hint "Example: outlook-contacts-search.ps1 -Name ""John"""
}

try {
    $null = Connect-Outlook
    $n = $Name.ToLowerInvariant(); $e = $Email.ToLowerInvariant(); $c = $Company.ToLowerInvariant()
    $all = Get-ContactRecords -Namespace $Namespace -AllStores:$AllStores
    $hits = @($all | Where-Object {
        (-not $n -or ([string]$_.FullName).ToLowerInvariant().Contains($n)) -and
        (-not $e -or ((@($_.Email, $_.Email2, $_.Email3) -join "`n").ToLowerInvariant().Contains($e))) -and
        (-not $c -or ([string]$_.Company).ToLowerInvariant().Contains($c))
    })
    $shown = @($hits | Select-Object -First $Limit)
    if ($Json) {
        Write-Json ([ordered]@{ Total = $hits.Count; Count = $shown.Count; Contacts = $shown })
        exit 0
    }
    $crit = @(); if ($Name) { $crit += "name '$Name'" }; if ($Email) { $crit += "email '$Email'" }; if ($Company) { $crit += "company '$Company'" }
    Write-Host "=== CONTACT SEARCH: $($crit -join ' and ') ($($hits.Count) found) ===" -ForegroundColor Cyan
    if ($hits.Count -eq 0) { Write-Host "No contacts match. Try outlook-people.ps1 to search the directory and recent mail too." -ForegroundColor Gray; exit 0 }
    $i = 0
    foreach ($r in $shown) { $i++; Write-ContactRecord $r -Number $i }
    if ($hits.Count -gt $shown.Count) { Write-Host "($($hits.Count - $shown.Count) more; raise -Limit)" -ForegroundColor DarkGray }
}
catch {
    Stop-WithError $_
}
