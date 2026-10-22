param(
    [Parameter(Position = 0)]
    [string]$Query = "",
    [int]$Limit = 10,
    [ValidateSet("Contacts", "Directory", "Recent")]
    [string[]]$Sources = @("Contacts", "Directory", "Recent"),
    [int]$Days = 180,
    [switch]$Detail,
    [switch]$Profile,
    [int]$ProfileDays = 365,
    [int]$UpcomingDays = 60,
    [switch]$RefreshDirectory,
    [string]$Members = "",
    [string]$Org = "",
    [int]$Levels = 4,
    [string]$Department = "",
    [string]$Title = "",
    [switch]$Json
)

# Outlook People: find anyone by name, partial name or email, with their contact points.
# Searches your Contacts (every mailbox), your organization's directory (Global Address List) and the
# people in your Inbox and Sent Items from the last -Days. Best match first.
#
# Usage: .\outlook-people.ps1 "jane"                      # first name, last name, partial or email
# Usage: .\outlook-people.ps1 "jane doe" -Detail -Json     # -Detail adds each person's manager
# Usage: .\outlook-people.ps1 "jane doe" -Profile          # "who is this": one person, with mail and meetings together
# Usage: .\outlook-people.ps1 -Members "group name or address"   # members of a distribution list or contact group
# Usage: .\outlook-people.ps1 -Org "jane doe"               # manager chain, direct reports, peers
# Usage: .\outlook-people.ps1 -Department "information tech" -Title "director"   # directory by attribute
# Usage: .\outlook-people.ps1 -RefreshDirectory            # rebuild the local directory name index now
#
# Recent mail means senders in every received folder (Inbox and rule-filed folders) and
# recipients in Sent Items.
# -Profile shows contact points plus counts and dates of mail exchanged (received folders and
# Sent Items, last -ProfileDays) and meetings together (past -ProfileDays, next -UpcomingDays). Subjects
# only, never bodies.

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

if (-not $Query -and -not $RefreshDirectory -and -not $Members -and -not $Org -and -not $Department -and -not $Title) {
    Stop-WithError "Give a name, partial name or email." -Hint "Example: outlook-people.ps1 ""jane doe"""
}

# Mail exchanged with one person: received mail from them (every received folder), Sent Items to them.
function Get-MailWith {
    param($Namespace, $Person, [int]$Days)
    $since = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-dd HH:mm")
    $email = ConvertTo-DaslText $Person.Email
    $name = ConvertTo-DaslText $Person.Name
    $recent = New-Object System.Collections.Generic.List[object]
    $result = [ordered]@{ Days = $Days; Received = 0; Sent = 0; LastReceived = $null; LastSent = $null; Recent = $recent }

    $cond = """urn:schemas:httpmail:fromemail"" = '$email' OR ""$($script:Prop.SenderSmtp)"" = '$email'"
    $found = Find-ReceivedRows -Namespace $Namespace -Condition $cond -Days $Days   # assign first: the list comes back comma-wrapped
    $rows = @($found | Sort-Object { $_[2] } -Descending)
    $result.Received = $rows.Count
    if ($rows.Count -gt 0) { $result.LastReceived = $rows[0][2] }
    foreach ($row in ($rows | Select-Object -First 5)) { $recent.Add([ordered]@{ Direction = "from"; Date = $row[2]; Subject = $row[3] }) }
    if ($name) {
        $sent = $Namespace.GetDefaultFolder(5)
        $f = "@SQL=(""urn:schemas:httpmail:displayto"" LIKE '%$name%' OR ""urn:schemas:httpmail:displaycc"" LIKE '%$name%') AND ""urn:schemas:httpmail:date"" >= '$since'"
        $t = $sent.GetTable($f)
        $t.Columns.RemoveAll(); [void]$t.Columns.Add("Subject"); [void]$t.Columns.Add("SentOn")
        $t.Sort("SentOn", $true)
        $result.Sent = $t.GetRowCount()
        $n = [Math]::Min($result.Sent, 5)
        if ($n -gt 0) {
            $rows = $t.GetArray($n)
            $result.LastSent = $rows[0, 1]
            for ($i = 0; $i -lt $n; $i++) {
                $r = [ordered]@{ Direction = "to"; Date = $rows[$i, 1]; Subject = [string]$rows[$i, 0] }
                $recent.Add($r)
            }
        }
    }
    $sorted = @($recent | Sort-Object { $_.Date } -Descending | Select-Object -First 5)
    $result.Recent = $sorted
    return $result
}

# Meetings with one person (organizer or attendee), past and upcoming, canceled ones skipped.
# One indexed table query finds the meetings that name them; recurring series are expanded with
# one date + subject restrict, and each occurrence is checked for the person, so the calendar is
# never walked item by item. Approximate: matching is by display name, so an attendee shown only
# as a raw address or as "Last, First" is missed.
function Get-MeetingsWith {
    param($Namespace, $Person, [int]$PastDays, [int]$FutureDays)
    $cal = $Namespace.GetDefaultFolder(9)
    $from = (Get-Date).Date.AddDays(-$PastDays); $to = (Get-Date).Date.AddDays($FutureDays + 1)
    $now = Get-Date
    $res = [ordered]@{ PastDays = $PastDays; UpcomingDays = $FutureDays; Past = 0; Upcoming = 0; Last = $null; Next = $null }
    $name = ConvertTo-DaslText ([string]$Person.Name)
    if (-not $name) { return $res }
    $needle = ([string]$Person.Name).ToLowerInvariant()
    $f = "@SQL=(""urn:schemas:httpmail:displayto"" LIKE '%$name%' OR ""urn:schemas:httpmail:displaycc"" LIKE '%$name%' OR ""urn:schemas:calendar:organizer"" LIKE '%$name%')"
    $t = $cal.GetTable($f)
    $t.Columns.RemoveAll()
    foreach ($c in @("Subject", "Start", "IsRecurring", "MeetingStatus")) { [void]$t.Columns.Add($c) }
    $starts = New-Object System.Collections.Generic.List[object]
    $series = @{}
    while (-not $t.EndOfTable) {
        $row = $t.GetNextRow()
        if ([int]$row.Item(4) -in 0, 5, 7) { continue }   # not a meeting, or canceled (olMeetingCanceled, olMeetingReceivedAndCanceled)
        if ($row.Item(3)) { $series[[string]$row.Item(1)] = 1; continue }
        $s = [datetime]$row.Item(2)
        if ($s -ge $from -and $s -lt $to) { $starts.Add([ordered]@{ Start = $s; Subject = [string]$row.Item(1) }) }
    }
    if ($series.Count -gt 0) {
        $items = $cal.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")
        $window = "[Start] >= '$(Format-JetDate $from)' AND [Start] < '$(Format-JetDate $to)'"
        # One restrict for all series: each Restrict call expands the recurrences once.
        $subjects = @($series.Keys | ForEach-Object { "[Subject] = '$($_.Replace("'", "''"))'" }) -join " OR "
        $occ = $items.Restrict("$window AND ($subjects)")
        # A subject can be shared by another series without this person: keep only occurrences
        # that name them (organizer or attendee) and are not canceled.
        foreach ($o in $occ) {
            if ([int]$o.MeetingStatus -in 0, 5, 7) { continue }
            $who = (@([string]$o.Organizer, [string]$o.RequiredAttendees, [string]$o.OptionalAttendees) -join ";").ToLowerInvariant()
            if (-not $who.Contains($needle)) { continue }
            $starts.Add([ordered]@{ Start = $o.Start; Subject = [string]$o.Subject })
        }
    }
    foreach ($m in ($starts | Sort-Object { $_.Start })) {
        if ($m.Start -lt $now) { $res.Past++; $res.Last = $m }
        else { $res.Upcoming++; if (-not $res.Next) { $res.Next = $m } }
    }
    return $res
}
function ConvertTo-PersonJson($p) {
    $o = [ordered]@{}
    foreach ($k in $p.Keys) { $o[$k] = $p[$k] }
    $o.LastContact = Format-MailDateIso $p.LastContact
    $o.Sources = @($p.Sources)
    $o.OtherEmails = @($p.OtherEmails)
    return $o
}

function Write-PersonText($p, [int]$Number = 0) {
    $head = if ($Number -gt 0) { "[$Number] " } else { "" }
    $kind = if ($p.Kind -and $p.Kind -ne "Person") { " [$($p.Kind)]" } else { "" }
    # One Write-Host per line: the warm host renders each call as its own line.
    Write-Host "$head$($p.Name)$kind  ($(@($p.Sources) -join ', '); score $($p.Score))" -ForegroundColor Yellow
    $role = @($p.Title, $p.Department, $p.Company) | Where-Object { $_ }
    if ($role) { Write-Host "    $($role -join ' | ')" -ForegroundColor Gray }
    if ($p.Email) { Write-Host "    Email: $($p.Email)" }
    foreach ($e in @($p.OtherEmails)) { if ($e) { Write-Host "    Email: $e" } }
    if ($p.BusinessPhone) { Write-Host "    Work: $($p.BusinessPhone)" }
    if ($p.MobilePhone) { Write-Host "    Mobile: $($p.MobilePhone)" }
    if ($p.HomePhone) { Write-Host "    Home: $($p.HomePhone)" }
    if ($p.Office) { Write-Host "    Office: $($p.Office)" }
    if ($p.Manager) { Write-Host "    Manager: $($p.Manager)" }
    if ($p.Received -or $p.Sent) {
        Write-Host "    Recent mail: $($p.Received) received, $($p.Sent) sent; last $(Format-MailDate $p.LastContact)" -ForegroundColor DarkGray
    }
    if ($p.ContactEntryID) { Write-Host "    Contact EntryID: $($p.ContactEntryID)" -ForegroundColor DarkGray }
}

try {
    $null = Connect-Outlook

    if ($RefreshDirectory) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $names = Get-DirectoryNames -Namespace $Namespace -Refresh
        $ms = $sw.ElapsedMilliseconds
        if (-not $Query) {
            if ($Json) { Write-Json ([ordered]@{ Action = "refresh-directory"; Entries = $names.Count; Ms = $ms }) }
            else { Write-Host "Directory name index rebuilt: $($names.Count) entries in $([Math]::Round($ms / 1000, 1)) s." -ForegroundColor Green }
            exit 0
        }
    }

    if ($Org) {
        $peerLimit = if ($PSBoundParameters.ContainsKey("Limit")) { $Limit } else { 10 }
        if (-not $PSBoundParameters.ContainsKey("Limit")) { $Limit = 25 }
        $found = Resolve-Person -Namespace $Namespace -Query $Org -Limit 5 -Sources Contacts, Directory
        $cands = @($found | Where-Object { $_.Email -and $_.Kind -eq "Person" -and $_.Sources -contains "Directory" })
        if ($cands.Count -eq 0) { Stop-WithError "No one in the directory matches '$Org'." -Hint "Org charts exist only for directory people. Try the full name or the email." }
        $user = $null
        $rc = $Namespace.CreateRecipient($cands[0].Email)
        if ($rc.Resolve() -and (Test-ExchangeEntry $rc.AddressEntry)) { $user = $rc.AddressEntry.GetExchangeUser() }
        if (-not $user) { Stop-WithError "Could not open the directory entry for '$Org'." }
        $chart = Get-OrgChart -User $user -Levels $Levels -Limit $Limit -PeerLimit $peerLimit
        $others = @($cands | Select-Object -Skip 1 | Where-Object { $_.Match -ge $cands[0].Match })
        if ($Json) {
            $o = [ordered]@{ Query = $Org }
            foreach ($k in $chart.Keys) { $o[$k] = $chart[$k] }
            $o.OtherMatches = @(foreach ($x in $others) { [ordered]@{ Name = $x.Name; Email = $x.Email } })
            Write-Json $o
            exit 0
        }
        $line = { param($p) "$($p.Name)$(if ($p.Title) { " | $($p.Title)" })$(if ($p.Department) { " | $($p.Department)" })$(if ($p.Email) { " <$($p.Email)>" })" }
        Write-Host "=== ORG: $($chart.Person.Name) ===" -ForegroundColor Cyan
        Write-Host (& $line $chart.Person) -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Manager chain ($($chart.Managers.Count) levels, closest first):" -ForegroundColor Cyan
        if ($chart.Managers.Count -eq 0) { Write-Host "  (no manager listed)" -ForegroundColor Gray }
        $i = 0; foreach ($m in $chart.Managers) { $i++; Write-Host "  $i. $(& $line $m)" }
        foreach ($sec in @(@("Direct reports", $chart.DirectReports), @("Peers (same manager)", $chart.Peers))) {
            Write-Host ""
            Write-Host "$($sec[0]): $($sec[1].Total)" -ForegroundColor Cyan
            foreach ($p in $sec[1].People) { Write-Host "  - $(& $line $p)" }
            if ($sec[1].Total -gt $sec[1].People.Count) { Write-Host "  ($($sec[1].Total - $sec[1].People.Count) more; raise -Limit)" -ForegroundColor DarkGray }
        }
        if ($others.Count -gt 0) {
            Write-Host ""
            Write-Host "Other people matching '$Org' (use the email to pick one):" -ForegroundColor DarkGray
            foreach ($x in $others) { Write-Host "  $($x.Name) <$($x.Email)>" -ForegroundColor DarkGray }
        }
        exit 0
    }

    if ($Department -or $Title) {
        if (-not $PSBoundParameters.ContainsKey("Limit")) { $Limit = 25 }
        $res = Find-DirectoryByAttribute -Namespace $Namespace -Department $Department -Title $Title -Query $Query -Limit $Limit
        # Email and phone for the people shown (one directory read each). A display name shared by
        # several entries can only open the first, so those get no address.
        $names = Get-DirectoryNames -Namespace $Namespace
        $gal = Get-DirectoryList -Namespace $Namespace
        foreach ($r in $res.Records) {
            $r.Email = ""; $r.BusinessPhone = ""; $r.Office = ""
            if (@($names -eq $r.Name).Count -gt 1) { $r.Email = "(several directory entries share this name)"; continue }
            try {
                $e = $gal.AddressEntries.Item($r.Name)
                if ($e -and $e.Name -eq $r.Name) {
                    $p = ConvertTo-DirectoryPerson -Entry $e
                    $r.Email = $p.Email; $r.BusinessPhone = $p.BusinessPhone; $r.Office = $p.Office
                }
            } catch { }
        }
        $crit = @(); if ($Department) { $crit += "department '$Department'" }; if ($Title) { $crit += "title '$Title'" }; if ($Query) { $crit += "name '$Query'" }
        if ($Json) {
            Write-Json ([ordered]@{ Department = $Department; Title = $Title; Query = $Query; Total = $res.Total; Count = @($res.Records).Count; People = @($res.Records) })
            exit 0
        }
        Write-Host "=== DIRECTORY: $($crit -join ' and ') ($($res.Total) found) ===" -ForegroundColor Cyan
        $i = 0
        foreach ($r in $res.Records) {
            $i++
            Write-Host "[$i] $($r.Name)" -ForegroundColor Yellow
            $role = @($r.Title, $r.Department) | Where-Object { $_ }
            if ($role) { Write-Host "    $($role -join ' | ')" -ForegroundColor Gray }
            if ($r.Email) { Write-Host "    Email: $($r.Email)" }
            if ($r.BusinessPhone) { Write-Host "    Work: $($r.BusinessPhone)" }
            if ($r.Office) { Write-Host "    Office: $($r.Office)" }
        }
        if ($res.Total -gt @($res.Records).Count) { Write-Host "($($res.Total - @($res.Records).Count) more; raise -Limit or narrow the search)" -ForegroundColor DarkGray }
        exit 0
    }

    if ($Members) {
        if (-not $PSBoundParameters.ContainsKey("Limit")) { $Limit = 100 }   # each member is a server lookup (~25 ms)
        $grp = Get-GroupMembers -Namespace $Namespace -Group $Members -Limit $Limit
        if ($Json) {
            $list = @(foreach ($m in $grp.Members) { $m })
            Write-Json ([ordered]@{ Group = $grp.Name; Email = $grp.Email; Source = $grp.Source; Total = $grp.Total; Count = $list.Count; Members = $list })
            exit 0
        }
        $addr = if ($grp.Email) { " <$($grp.Email)>" } else { "" }
        Write-Host "=== MEMBERS: $($grp.Name)$addr ($($grp.Source), $($grp.Total) members) ===" -ForegroundColor Cyan
        $i = 0
        foreach ($m in $grp.Members) {
            $i++
            $kind = if ($m.Kind -ne "Person") { " [$($m.Kind)]" } else { "" }
            $mail = if ($m.Email) { " <$($m.Email)>" } else { "" }
            Write-Host "[$i] $($m.Name)$mail$kind"
        }
        if ($grp.Total -gt $grp.Members.Count) { Write-Host "($($grp.Total - $grp.Members.Count) more; raise -Limit)" -ForegroundColor DarkGray }
        exit 0
    }

    if (-not $Profile) {
        $people = Resolve-Person -Namespace $Namespace -Query $Query -Limit $Limit -Sources $Sources -Days $Days -Detail:$Detail
        if ($Json) {
            $list = @(foreach ($p in $people) { ConvertTo-PersonJson $p })
            Write-Json ([ordered]@{ Query = $Query; Count = $list.Count; People = $list })
            exit 0
        }
        Write-Host "=== PEOPLE: '$Query' ($($people.Count) found) ===" -ForegroundColor Cyan
        if ($people.Count -eq 0) {
            Write-Host "Nobody matches in $($Sources -join ', ') (recent mail: last $Days days). Try part of the name or the email." -ForegroundColor Gray
            exit 0
        }
        $i = 0
        foreach ($p in $people) { $i++; Write-PersonText $p $i }
        exit 0
    }

    # -Profile: one person
    # Contacts and the directory first (fast); recent mail only when they find nobody.
    $quick = @($Sources | Where-Object { $_ -ne "Recent" })
    $people = @()
    if ($quick.Count -gt 0) {
        $found = Resolve-Person -Namespace $Namespace -Query $Query -Limit 5 -Sources $quick -Days $Days
        $people = @($found | Where-Object { $_.Email -and $_.Kind -eq "Person" })
    }
    if ($people.Count -eq 0 -and $Sources -contains "Recent") {
        $found = Resolve-Person -Namespace $Namespace -Query $Query -Limit 5 -Sources Recent -Days $Days
        $people = @($found | Where-Object { $_.Email -and $_.Kind -eq "Person" })
    }
    if ($people.Count -eq 0) { Stop-WithError "No person found for '$Query'." -Hint "Try part of the name, or the email address." }
    $p = $people[0]
    if (-not $p.Manager) {
        # Manager for the chosen person only (one directory lookup, not one per candidate).
        try {
            $rc = $Namespace.CreateRecipient($p.Email)
            if ($rc.Resolve() -and (Test-ExchangeEntry $rc.AddressEntry)) {
                $m = $rc.AddressEntry.GetExchangeUser().GetExchangeUserManager()
                if ($m) { $p.Manager = [string]$m.Name }
            }
        } catch { }
    }
    $others = @($people | Select-Object -Skip 1 | Where-Object { $_.Score -ge $p.Score - 10 })
    $mail = Get-MailWith -Namespace $Namespace -Person $p -Days $ProfileDays
    $meet = Get-MeetingsWith -Namespace $Namespace -Person $p -PastDays $ProfileDays -FutureDays $UpcomingDays

    if ($Json) {
        $m = [ordered]@{ Days = $mail.Days; Received = $mail.Received; Sent = $mail.Sent
            LastReceived = (Format-MailDateIso $mail.LastReceived); LastSent = (Format-MailDateIso $mail.LastSent)
            Recent = @(foreach ($r in $mail.Recent) { [ordered]@{ Direction = $r.Direction; Date = (Format-MailDateIso $r.Date); Subject = $r.Subject } }) }
        $mt = [ordered]@{ PastDays = $meet.PastDays; UpcomingDays = $meet.UpcomingDays; Past = $meet.Past; Upcoming = $meet.Upcoming
            Last = $(if ($meet.Last) { [ordered]@{ Start = (Format-MailDateIso $meet.Last.Start); Subject = $meet.Last.Subject } } else { $null })
            Next = $(if ($meet.Next) { [ordered]@{ Start = (Format-MailDateIso $meet.Next.Start); Subject = $meet.Next.Subject } } else { $null }) }
        $alt = @(foreach ($o in $others) { [ordered]@{ Name = $o.Name; Email = $o.Email; Score = $o.Score } })
        Write-Json ([ordered]@{ Query = $Query; Person = (ConvertTo-PersonJson $p); Mail = $m; Meetings = $mt; OtherMatches = $alt })
        exit 0
    }

    Write-Host "=== WHO IS: $($p.Name) ===" -ForegroundColor Cyan
    Write-PersonText $p
    Write-Host ""
    Write-Host "Mail (last $ProfileDays days, Inbox and Sent Items): $($mail.Received) from them, $($mail.Sent) to them" -ForegroundColor Cyan
    if ($mail.LastReceived) { Write-Host "  Last from them: $(Format-MailDate $mail.LastReceived)" }
    if ($mail.LastSent) { Write-Host "  Last to them:   $(Format-MailDate $mail.LastSent)" }
    foreach ($r in $mail.Recent) { Write-Host "  $(Format-MailDate $r.Date)  $($r.Direction.PadRight(4)) $($r.Subject)" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "Meetings together: $($meet.Past) in the last $ProfileDays days, $($meet.Upcoming) in the next $UpcomingDays days" -ForegroundColor Cyan
    if ($meet.Last) { Write-Host "  Last: $(Format-MailDate $meet.Last.Start)  $($meet.Last.Subject)" }
    if ($meet.Next) { Write-Host "  Next: $(Format-MailDate $meet.Next.Start)  $($meet.Next.Subject)" }
    if ($others.Count -gt 0) {
        Write-Host ""
        Write-Host "Other close matches (use the email to pick one):" -ForegroundColor DarkGray
        foreach ($o in $others) { Write-Host "  $($o.Name) <$($o.Email)>" -ForegroundColor DarkGray }
    }
}
catch {
    Stop-WithError $_
}
