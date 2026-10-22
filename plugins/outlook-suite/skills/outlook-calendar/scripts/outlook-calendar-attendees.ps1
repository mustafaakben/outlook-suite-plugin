param(
    [string]$EntryID = "",
    [int]$Index = 0,
    [datetime]$OccurrenceStart,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Calendar Attendees: who is invited to a meeting and how each one responded.
# Usage: .\outlook-calendar-attendees.ps1 -EntryID "00000000..."
# Usage: .\outlook-calendar-attendees.ps1 -EntryID "00000000..." -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

$typeNames = @{ 0 = "Organizer"; 1 = "Required"; 2 = "Optional"; 3 = "Resource" }
$colors = @{ "Organizer" = "Cyan"; "Tentative" = "Yellow"; "Accepted" = "Green"; "Declined" = "Red"; "None" = "Gray"; "Not Responded" = "DarkGray" }

try {
    $null = Connect-Outlook
    $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
    $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -Index $Index -Days $Days -OccurrenceStart $occ
    $rec = ConvertTo-EventRecord $item

    $people = New-Object System.Collections.Generic.List[object]
    foreach ($r in $item.Recipients) {
        $smtp = ""
        try { $u = $r.AddressEntry.GetExchangeUser(); if ($u) { $smtp = $u.PrimarySmtpAddress } } catch { }
        if (-not $smtp) { try { $smtp = $r.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS) } catch { } }
        if (-not $smtp) { $smtp = [string]$r.Address }
        $people.Add([ordered]@{
            Name     = [string]$r.Name
            Email    = $smtp
            Type     = $typeNames[[int]$r.Type]
            Response = $script:ResponseNames[[int]$r.MeetingResponseStatus]
        })
    }
    $summary = [ordered]@{
        Accepted  = @($people | Where-Object { $_.Response -eq "Accepted" }).Count
        Tentative = @($people | Where-Object { $_.Response -eq "Tentative" }).Count
        Declined  = @($people | Where-Object { $_.Response -eq "Declined" }).Count
        NoResponse = @($people | Where-Object { $_.Response -in @("None", "Not Responded") }).Count
    }

    if ($Json) {
        $o = ConvertTo-EventJson $rec
        $o.People = $people.ToArray()
        $o.Summary = $summary
        Write-Json $o; exit 0
    }

    Write-Host "`n=== MEETING ATTENDEES ===" -ForegroundColor Cyan
    Write-EventRecord $rec -Prefix "  "
    if ($people.Count -eq 0) { Write-Host "`nThis is an appointment (no attendees)." -ForegroundColor Gray; exit 0 }
    foreach ($t in "Required", "Optional", "Resource") {
        $group = @($people | Where-Object { $_.Type -eq $t })
        if ($group.Count -eq 0) { continue }
        Write-Host "`n${t}:" -ForegroundColor White
        foreach ($p in $group) {
            $c = $colors[$p.Response]; if (-not $c) { $c = "Gray" }
            Write-Host "  $($p.Name) <$($p.Email)>" -NoNewline
            Write-Host " - $($p.Response)" -ForegroundColor $c
        }
    }
    Write-Host "`n--- Responses: $($summary.Accepted) accepted, $($summary.Tentative) tentative, $($summary.Declined) declined, $($summary.NoResponse) no response ---" -ForegroundColor Cyan
}
catch {
    Stop-WithError $_
}
