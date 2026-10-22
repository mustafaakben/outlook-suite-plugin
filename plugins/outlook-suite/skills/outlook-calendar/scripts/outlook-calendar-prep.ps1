param(
    [string]$EntryID = "",
    [datetime]$OccurrenceStart,
    [switch]$Next,
    [int]$Days = 90,
    [int]$MaxAttendees = 25,
    [int]$Related = 10,
    [string]$TimeZone = "",
    [switch]$Json
)

# Outlook Calendar Prep: a read-only brief for one meeting.
#   - the invite: when, where, organizer, join link, agenda (the invite's own text)
#   - the people: title, department and office from the directory, their response, and how
#     much mail the user exchanged with each in the last -Days (count and last date)
#   - related mail: messages whose subjects share keywords with the meeting subject
#     (subjects, dates and senders only, never bodies)
# Nothing is changed or sent.
# Usage: .\outlook-calendar-prep.ps1 -Next                       # the user's next real meeting
# Usage: .\outlook-calendar-prep.ps1 -EntryID "00000000..." [-OccurrenceStart "2026-11-03 09:00"]
# Usage: .\outlook-calendar-prep.ps1 -Next -Days 180 -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_calendar.ps1")

$stopWords = @("meeting", "meet", "sync", "call", "chat", "zoom", "teams", "with", "and", "the", "for", "about", "re", "fw", "fwd",
    "discussion", "discuss", "check", "checkin", "check-in", "catch", "catchup", "update", "updates", "weekly", "monthly", "daily",
    "session", "office", "hours", "invitation", "invite", "accepted", "tentative", "canceled", "cancelled", "from", "this", "that", "your", "our")

try {
    $null = Connect-Outlook
    Set-DisplayTimeZone $TimeZone
    $sw = [Diagnostics.Stopwatch]::StartNew()

    # 1. The meeting
    if ($Next) {
        $now = Get-Date
        $pick = @()
        foreach ($span in 2, 14) {   # the next meeting is nearly always within 2 days; read 14 only if not
            $upcoming = Get-CalendarEvents -Namespace $Namespace -From $now -To $now.AddDays($span)
            $pick = @($upcoming | Where-Object { $_.Kind -eq "meeting" -and -not $_.Cancelled -and $_.Response -ne "Declined" -and $_.Busy -ne "Free" -and $_.Start -ge $now.AddMinutes(-5) })
            if ($pick.Count) { break }
        }
        if ($pick.Count -eq 0) { Stop-WithError "No upcoming meeting in the next 14 days." }
        $item = Get-CalendarItem -Namespace $Namespace -EntryID $pick[0].EntryID -OccurrenceStart $(if ($pick[0].Recurring) { $pick[0].Start } else { $null })
    } else {
        $occ = $null; if ($PSBoundParameters.ContainsKey("OccurrenceStart")) { $occ = $OccurrenceStart }
        $item = Get-CalendarItem -Namespace $Namespace -EntryID $EntryID -OccurrenceStart $occ
    }
    $rec = ConvertTo-EventRecord $item

    # 2. Invite details: join link and agenda from the invite text
    $body = ""; try { $body = [string]$item.Body } catch { }
    $joinPattern = '(https://teams\.microsoft\.com/l/meetup-join/[^\s>"]+|https://[\w.-]*zoom\.us/(j|my|w)/[^\s>"]+|https://meet\.google\.com/[^\s>"]+|https://[\w.-]*webex\.com/[^\s>"]+)'
    $join = ""
    foreach ($src in @([string]$item.Location, $body)) { $m = [regex]::Match($src, $joinPattern); if ($m.Success) { $join = $m.Value.TrimEnd(">", ")", ".", ","); break } }
    # Agenda: the invite text before the Teams/Zoom boilerplate, links shortened, blank runs collapsed.
    $agenda = ($body -split "(?m)^_{10,}|(?im)^\s*Microsoft Teams( meeting| Need help)|(?im)^\s*Join Zoom Meeting")[0]
    $agenda = [regex]::Replace($agenda, 'https?://\S+', '<link>')
    $agenda = [regex]::Replace($agenda.Trim(), '(\r?\n\s*){3,}', "`r`n`r`n")
    if ($agenda.Length -gt 1500) { $agenda = $agenda.Substring(0, 1500) + " ..." }

    # 3. People: everyone invited except the user, with directory details
    $self = ""; try { $self = [string]$Namespace.CurrentUser.Name } catch { }
    $typeNames = @{ 0 = "Organizer"; 1 = "Required"; 2 = "Optional"; 3 = "Resource" }
    $people = New-Object System.Collections.Generic.List[object]
    $total = 0
    foreach ($r in $item.Recipients) {
        if ([string]$r.Name -eq $self) { continue }
        $total++
        if ($people.Count -ge $MaxAttendees) { continue }
        $p = [ordered]@{ Name = [string]$r.Name; Email = ""; Type = $typeNames[[int]$r.Type]; Response = $script:ResponseNames[[int]$r.MeetingResponseStatus]
            Title = ""; Department = ""; Office = ""; Received = 0; Sent = 0; LastContact = $null }
        $entry = $null; try { $entry = $r.AddressEntry } catch { }
        if ($entry -and (Get-Command ConvertTo-DirectoryPerson -ErrorAction SilentlyContinue)) {
            $d = ConvertTo-DirectoryPerson -Entry $entry
            $p.Email = $d.Email; $p.Title = $d.Title; $p.Department = $d.Department; $p.Office = $d.Office
        } elseif ($entry) {
            try { $u = $entry.GetExchangeUser(); if ($u) { $p.Email = [string]$u.PrimarySmtpAddress; $p.Title = [string]$u.JobTitle; $p.Department = [string]$u.Department; $p.Office = [string]$u.OfficeLocation } } catch { }
        }
        if (-not $p.Email) { $p.Email = [string]$r.Address }
        $people.Add($p)
    }
    # The organizer, when it is someone else and not already a recipient
    if ($rec.Organizer -and $rec.Organizer -ne $self -and @($people | Where-Object { $_.Name -eq $rec.Organizer }).Count -eq 0) {
        $o = [ordered]@{ Name = $rec.Organizer; Email = ""; Type = "Organizer"; Response = "Organizer"; Title = ""; Department = ""; Office = ""; Received = 0; Sent = 0; LastContact = $null }
        try { $e = $item.GetOrganizer(); if ($e) { $d = ConvertTo-DirectoryPerson -Entry $e; $o.Email = $d.Email; $o.Title = $d.Title; $o.Department = $d.Department; $o.Office = $d.Office } } catch { }
        $people.Insert(0, $o); $total++
    }

    # 4. Mail exchanged with each person: received (every received folder) and sent (Sent Items).
    $since = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-dd HH:mm")
    $byEmail = @{}; foreach ($p in $people) { if ($p.Email -match "@") { $byEmail[$p.Email.ToLowerInvariant()] = $p } }
    $haveReceived = [bool](Get-Command Find-ReceivedRows -ErrorAction SilentlyContinue)
    if ($byEmail.Count) {
        $cond = (($byEmail.Keys | ForEach-Object { """$($script:Prop.SenderSmtp)"" = '$(ConvertTo-DaslText $_)'" }) -join " OR ")
        $rows = @()
        if ($haveReceived) { $rows = Find-ReceivedRows -Namespace $Namespace -Condition $cond -Days $Days }
        else {
            $t = $Namespace.GetDefaultFolder(6).GetTable("@SQL=($cond) AND ""urn:schemas:httpmail:datereceived"" >= '$since'")
            $t.Columns.RemoveAll(); foreach ($c in @("SenderName", $script:Prop.SenderSmtp, "ReceivedTime", "Subject")) { [void]$t.Columns.Add($c) }
            while (-not $t.EndOfTable) { $v = $t.GetNextRow().GetValues(); $rows += , @([string]$v[0], [string]$v[1], $v[2], [string]$v[3]) }
        }
        foreach ($row in $rows) {
            $p = $byEmail[([string]$row[1]).ToLowerInvariant()]
            if (-not $p) { continue }
            $p.Received++
            if (-not $p.LastContact -or $row[2] -gt $p.LastContact) { $p.LastContact = $row[2] }
        }
    }
    $named = @($people | Where-Object { $_.Name })
    if ($named.Count) {
        $cond = (($named | ForEach-Object { $n = ConvertTo-DaslText $_.Name; """urn:schemas:httpmail:displayto"" LIKE '%$n%' OR ""urn:schemas:httpmail:displaycc"" LIKE '%$n%'" }) -join " OR ")
        $t = $Namespace.GetDefaultFolder(5).GetTable("@SQL=($cond) AND ""urn:schemas:httpmail:date"" >= '$since'")
        $t.Columns.RemoveAll(); foreach ($c in @("To", "CC", "SentOn")) { [void]$t.Columns.Add($c) }
        while (-not $t.EndOfTable) {
            $v = $t.GetNextRow().GetValues()
            $line = ([string]$v[0] + ";" + [string]$v[1]).ToLowerInvariant()
            foreach ($p in $named) {
                if ($line.Contains($p.Name.ToLowerInvariant())) {
                    $p.Sent++
                    if (-not $p.LastContact -or $v[2] -gt $p.LastContact) { $p.LastContact = $v[2] }
                }
            }
        }
    }

    # 5. Related mail: subjects sharing keywords with the meeting subject (no bodies)
    $words = @([regex]::Matches($rec.Subject, "[\p{L}\p{N}][\p{L}\p{N}'&-]*") | ForEach-Object { $_.Value })
    $keywords = @($words | Where-Object { ($_.Length -ge 4 -or ($_.Length -ge 2 -and $_ -cmatch '^[A-Z0-9]+$')) -and $stopWords -notcontains $_.ToLowerInvariant() } | Select-Object -Unique -First 5)
    $relatedMail = New-Object System.Collections.Generic.List[object]
    if ($keywords.Count -gt 0 -and $Related -gt 0) {
        $need = [Math]::Min(2, $keywords.Count)
        $cond = (($keywords | ForEach-Object { """urn:schemas:httpmail:subject"" LIKE '%$(ConvertTo-DaslText $_)%'" }) -join " OR ")
        $cands = New-Object System.Collections.Generic.List[object]
        $recv = @()
        if ($haveReceived) { $recv = Find-ReceivedRows -Namespace $Namespace -Condition $cond -Days $Days }
        foreach ($row in $recv) { $cands.Add([ordered]@{ Subject = [string]$row[3]; Date = $row[2]; Who = "from $($row[0])"; Folder = "received" }) }
        $t = $Namespace.GetDefaultFolder(5).GetTable("@SQL=($cond) AND ""urn:schemas:httpmail:date"" >= '$since'")
        $t.Columns.RemoveAll(); foreach ($c in @("Subject", "SentOn", "To")) { [void]$t.Columns.Add($c) }
        while (-not $t.EndOfTable) { $v = $t.GetNextRow().GetValues(); $cands.Add([ordered]@{ Subject = [string]$v[0]; Date = $v[1]; Who = "to $([string]$v[2])"; Folder = "sent" }) }
        foreach ($c in $cands) {
            $hits = @($keywords | Where-Object { $c.Subject -match [regex]::Escape($_) }).Count
            if ($hits -ge $need) { $c.Matches = $hits; $relatedMail.Add($c) }
        }
    }
    $relatedTop = @($relatedMail | Sort-Object @{ Expression = { $_.Matches }; Descending = $true }, @{ Expression = { $_.Date }; Descending = $true } | Select-Object -First $Related)
    $elapsed = $sw.ElapsedMilliseconds

    if ($Json) {
        $o = ConvertTo-EventJson $rec
        $o.JoinLink = $join; $o.Agenda = $agenda
        $o.AttendeesTotal = $total; $o.MailDays = $Days
        $o.People = @($people | ForEach-Object { $x = ConvertTo-EventJson $_; $x.LastContact = Format-MailDateIso $_.LastContact; $x })
        $o.Keywords = $keywords
        $o.RelatedMail = @($relatedTop | ForEach-Object { $x = ConvertTo-EventJson $_; $x.Date = Format-MailDateIso $_.Date; $x })
        $o.ElapsedMs = $elapsed
        Write-Json $o
        exit 0
    }

    Write-Host "`n=== MEETING PREP ===" -ForegroundColor Cyan
    Write-EventRecord $rec -Prefix "  "
    if ($join) { Write-Host "      Join: $join" -ForegroundColor Cyan }
    Write-Host "`nAgenda (from the invite):" -ForegroundColor Cyan
    if ($agenda) { foreach ($l in ($agenda -split "\r?\n")) { Write-Host "  $l" } } else { Write-Host "  (the invite has no text)" -ForegroundColor DarkGray }

    Write-Host "`nPeople ($total besides you$(if ($total -gt $people.Count) { ", first $($people.Count) shown" })), mail in the last $Days days:" -ForegroundColor Cyan
    foreach ($p in $people) {
        $role = @($p.Title, $p.Department) | Where-Object { $_ }
        $mail = if ($p.Received + $p.Sent) { "$($p.Received) from, $($p.Sent) to them, last $(([datetime]$p.LastContact).ToString('MMM d'))" } else { "no mail" }
        Write-Host ("  {0}  [{1}, {2}]" -f $p.Name, $p.Type, $p.Response) -ForegroundColor White
        if ($role) { Write-Host "      $($role -join ', ')$(if ($p.Office) { "  ($($p.Office))" })" -ForegroundColor Gray }
        Write-Host "      $($p.Email)  |  $mail" -ForegroundColor DarkGray
    }

    Write-Host "`nRelated mail (subject keywords: $($keywords -join ', ')):" -ForegroundColor Cyan
    if ($relatedTop.Count -eq 0) { Write-Host "  none found" -ForegroundColor DarkGray }
    foreach ($m in $relatedTop) { Write-Host ("  {0}  {1}  ({2})" -f ([datetime]$m.Date).ToString("MMM d"), $m.Subject, $m.Who) }
    Write-Host "`n($elapsed ms; read-only, nothing was changed or sent)" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
