# Calendar helpers shared by the outlook-calendar scripts.
# Dot-source after _common.ps1:
#   . (Join-Path $PSScriptRoot "_common.ps1")
#   . (Join-Path $PSScriptRoot "_calendar.ps1")

$script:BusyNames = @{ 0 = "Free"; 1 = "Tentative"; 2 = "Busy"; 3 = "Out of Office"; 4 = "Working Elsewhere" }
$script:BusyValues = @{ "Free" = 0; "Tentative" = 1; "Busy" = 2; "OutOfOffice" = 3; "WorkingElsewhere" = 4 }
$script:ResponseNames = @{ 0 = "None"; 1 = "Organizer"; 2 = "Tentative"; 3 = "Accepted"; 4 = "Declined"; 5 = "Not Responded" }

# Items.Restrict on calendars wants this exact invariant format.
function ConvertTo-RestrictDate([datetime]$Date) {
    return $Date.ToString("MM/dd/yyyy HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
}

# Calendar events overlapping [From, To), recurrences expanded, sorted by start.
#   Pass 1: Items.Restrict with IncludeRecurrences and a few SetColumns columns (occurrences).
#   Pass 2: two Folder.GetTable reads (one-off events in the range, all series masters) for
#           EntryID, organizer, response and attendee counts, matched by start+subject
#           (one-off) or subject+time of day (series). Unmatched rows fall back to a slow pass.
#
# Kind:  allday  = all-day event
#        hold    = no one else invited (a plain appointment, or a meeting with only the user):
#                  the user's placeholder, may be booked over
#        firm    = no one else invited, but not movable: classes (course codes), office hours,
#                  exams, travel, flights, Out of Office time, private events (Get-FirmHoldPatterns)
#        meeting = other people are invited
# Blocking = takes time (Tentative, Busy, Out of Office) and is not cancelled or declined.
# Occurrences of a series carry the series' EntryID; use -OccurrenceStart to act on one.
function Get-CalendarEvents {
    param($Namespace, [datetime]$From, [datetime]$To, $Folder = $null, [string[]]$FirmPatterns = $null)
    if (-not $Folder) { $Folder = $Namespace.GetDefaultFolder(9) }
    if ($null -eq $FirmPatterns) { $FirmPatterns = Get-FirmHoldPatterns }
    $items = $Folder.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    $items.SetColumns("Subject,Start,End,MeetingStatus,BusyStatus,AllDayEvent,Location,IsRecurring")
    $filter = "[Start] < '$(ConvertTo-RestrictDate $To)' AND [End] > '$(ConvertTo-RestrictDate $From)'"
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in $items.Restrict($filter)) {
        $list.Add([ordered]@{
            EntryID       = $null
            Subject       = [string]$e.Subject
            Start         = [datetime]$e.Start
            End           = [datetime]$e.End
            Kind          = ""
            AllDay        = [bool]$e.AllDayEvent
            Busy          = $script:BusyNames[[int]$e.BusyStatus]
            Response      = ""
            Organizer     = ""
            Attendees     = 0
            Location      = [string]$e.Location
            Categories    = ""
            Private       = $false
            Recurring     = [bool]$e.IsRecurring
            Cancelled     = ([int]$e.MeetingStatus -eq 5 -or [int]$e.MeetingStatus -eq 7)
            Blocking      = $false
            ConflictsWith = @()
            MeetingStatus = [int]$e.MeetingStatus
        })
    }
    if ($list.Count -eq 0) { return , $list }

    $single = @{}; $series = @{}; $seriesBySubject = @{}
    $cols = @($script:Prop.LongTermEntryID, "Subject", "Start", "Organizer", "ResponseStatus", "RequiredAttendees", "OptionalAttendees", "Categories", "Sensitivity")
    $jet = "[Start] < '$(Format-JetDate $To)' AND [End] > '$(Format-JetDate $From)' AND [IsRecurring] = False"
    foreach ($pass in @(@{ Filter = $jet; Series = $false }, @{ Filter = "[IsRecurring] = True"; Series = $true })) {
        $tb = $Folder.GetTable($pass.Filter)
        $tb.Columns.RemoveAll()
        foreach ($c in $cols) { [void]$tb.Columns.Add($c) }
        while (-not $tb.EndOfTable) {
            $v = $tb.GetNextRow().GetValues()
            $meta = @{
                EntryID = (ConvertTo-HexId $v[0]); Organizer = [string]$v[3]; Response = [int]$v[4]
                Attendees = Get-OtherAttendeeCount -Names (([string]$v[5] + ";" + [string]$v[6]).Split(";")) -Organizer ([string]$v[3])
                Categories = [string]$v[7]; Private = ([int]$v[8] -eq 2)
            }
            $start = [datetime]$v[2]
            if ($pass.Series) {
                # Two series with the same subject and time of day: mark the key ambiguous ($null);
                # those occurrences then fall back to the exact slow match.
                $sk = "$($v[1])|$($start.ToString('HHmm'))"
                if ($series.ContainsKey($sk)) { $series[$sk] = $null } else { $series[$sk] = $meta }
                if ($seriesBySubject.ContainsKey([string]$v[1])) { $seriesBySubject[[string]$v[1]] = $null } else { $seriesBySubject[[string]$v[1]] = $meta }
            } else {
                $single["$($start.Ticks)|$($v[1])"] = $meta
            }
        }
    }

    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($r in $list) {
        if ($r.Recurring) {
            $sk = "$($r.Subject)|$($r.Start.ToString('HHmm'))"
            if ($series.ContainsKey($sk)) { $meta = $series[$sk] }       # $null when ambiguous
            else { $meta = $seriesBySubject[$r.Subject] }               # moved occurrence: unique subject only
        } else {
            $meta = $single["$($r.Start.Ticks)|$($r.Subject)"]
        }
        if ($meta) {
            $r.EntryID = $meta.EntryID; $r.Organizer = $meta.Organizer; $r.Attendees = $meta.Attendees
            $r.Response = $script:ResponseNames[$meta.Response]; $r.Categories = $meta.Categories; $r.Private = $meta.Private
        } else { $unmatched.Add($r) }
    }
    if ($unmatched.Count -gt 0) { Add-CalendarEntryIds -Namespace $Namespace -Records $unmatched -Folder $Folder -WithMeta }

    foreach ($r in $list) {
        $r.Kind = Get-EventKind -AllDay $r.AllDay -MeetingStatus $r.MeetingStatus -Attendees $r.Attendees -Subject $r.Subject -Busy $r.Busy -Private $r.Private -FirmPatterns $FirmPatterns -Location $r.Location
        $r.Blocking = ($r.Busy -in @("Tentative", "Busy", "Out of Office") -and -not $r.Cancelled -and $r.Response -ne "Declined")
        $r.Remove("MeetingStatus")
    }
    return , $list
}

# Slow fallback: EntryID (and with -WithMeta organizer, response, attendees) by walking the
# expanded range. Matched by start time and subject, so a record never gets another event's id.
function Add-CalendarEntryIds {
    param($Namespace, $Records, $Folder = $null, [switch]$WithMeta)
    $want = @($Records | Where-Object { -not $_.EntryID })
    if ($want.Count -eq 0) { return }
    if (-not $Folder) { $Folder = $Namespace.GetDefaultFolder(9) }
    $from = [datetime]::MaxValue; $to = [datetime]::MinValue
    $keys = @{}
    foreach ($r in $want) {
        $keys["$($r.Start.Ticks)|$($r.Subject)"] = $true
        if ($r.Start -lt $from) { $from = $r.Start }
        if ($r.Start -gt $to) { $to = $r.Start }
    }
    $found = @{}
    $items = $Folder.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    foreach ($e in $items.Restrict("[Start] >= '$(ConvertTo-RestrictDate $from)' AND [Start] < '$(ConvertTo-RestrictDate $to.AddMinutes(1))'")) {
        $k = "$(([datetime]$e.Start).Ticks)|$([string]$e.Subject)"
        if ($keys.ContainsKey($k) -and -not $found.ContainsKey($k)) {
            $m = @{ EntryID = [string]$e.EntryID }
            if ($WithMeta) {
                $m.Organizer = [string]$e.Organizer
                $m.Response = [int]$e.ResponseStatus
                $m.Attendees = 0; try { $m.Attendees = Get-OtherAttendeeCount -Names @($e.Recipients | ForEach-Object { [string]$_.Name }) -Organizer ([string]$e.Organizer) } catch { }
                $m.Categories = [string]$e.Categories
                $m.Private = ([int]$e.Sensitivity -eq 2)
            }
            $found[$k] = $m
            if ($found.Count -ge $keys.Count) { break }
        }
    }
    foreach ($r in $want) {
        $m = $found["$($r.Start.Ticks)|$($r.Subject)"]
        if (-not $m) { continue }
        $r.EntryID = $m.EntryID
        if ($WithMeta) { $r.Organizer = $m.Organizer; $r.Response = $script:ResponseNames[$m.Response]; $r.Attendees = $m.Attendees; $r.Categories = $m.Categories; $r.Private = $m.Private }
    }
}

# Flags overlapping meetings and firm events (ConflictsWith = subjects). Holds may overlap
# anything: they are the user's own placeholders and give way, so they are never a conflict.
function Set-CalendarConflicts {
    param($Records)
    $timed = @($Records | Where-Object { $_.Blocking -and $_.Kind -in @("meeting", "firm") })
    for ($i = 0; $i -lt $timed.Count; $i++) {
        for ($j = $i + 1; $j -lt $timed.Count; $j++) {
            if ($timed[$j].Start -ge $timed[$i].End) { break }
            $timed[$i].ConflictsWith += $timed[$j].Subject
            $timed[$j].ConflictsWith += $timed[$i].Subject
        }
    }
}

# Per-day summary used by the week overview and the free-time finder.
#   Tier 1 open        no meetings, no holds
#   Tier 2 holds-only  no meetings, only holds and firm events (classes, office hours, ...)
#   Tier 3 one-meeting one meeting booked (a new one makes 2, the daily maximum)
#   Tier 4 busy        2 or more meetings already (-MaxPerDay sets the cap)
# Away = an all-day Out of Office or Busy event (travel, conference): no slots that day.
function Get-CalendarDay {
    param($Records, [datetime]$Date, [int]$MaxPerDay = 2)
    $d0 = $Date.Date; $d1 = $d0.AddDays(1)
    $day = @($Records | Where-Object { $_.Start -lt $d1 -and $_.End -gt $d0 -and -not $_.Cancelled -and $_.Response -ne "Declined" })
    $meetings = @($day | Where-Object { $_.Kind -eq "meeting" -and $_.Busy -ne "Free" })
    $holds = @($day | Where-Object { $_.Kind -eq "hold" })
    $firm = @($day | Where-Object { $_.Kind -eq "firm" -and $_.Busy -ne "Free" })
    $allday = @($day | Where-Object { $_.Kind -eq "allday" })
    $away = @($allday | Where-Object { $_.Busy -in @("Out of Office", "Busy") }).Count -gt 0
    $conflicts = @($day | Where-Object { $_.ConflictsWith.Count -gt 0 }).Count
    $tier = 4
    if ($meetings.Count -eq 0 -and $holds.Count -eq 0 -and $firm.Count -eq 0) { $tier = 1 }
    elseif ($meetings.Count -eq 0) { $tier = 2 }
    elseif ($meetings.Count -lt $MaxPerDay) { $tier = 3 }
    return [ordered]@{
        Date      = $d0
        Tier      = $tier
        TierName  = $script:TierNames[$tier]
        Meetings  = $meetings.Count
        Holds     = $holds.Count
        Firm      = $firm.Count
        AllDay    = @($allday | ForEach-Object { $_.Subject })
        Away      = $away
        OverCap   = ($meetings.Count -gt $MaxPerDay)
        Conflicts = $conflicts
        Events    = $day
    }
}
$script:TierNames = @{ 1 = "Open day"; 2 = "Holds-only day"; 3 = "One-meeting day"; 4 = "Busy day" }
# Subjects of no-attendee events that are commitments, not movable holds. Regex, matched
# case-sensitively (prefix (?i) for case-insensitive). Override with
# "calendar": { "firmHolds": [ ... ] } in %APPDATA%\outlook-suite\config.json, or -FirmHolds
# on list, week and free.
$script:DefaultFirmHolds = @(
    '^[A-Z]{2,4}[- ]?\d{3,4}\b',  # course code first: XYZ 101, ABC1234
    '\b[A-Z]{2,4}-\d{3,4}\b',     # course code anywhere: ABC-1234-A
    '(?i)\bclass(es)?\b',
    '(?i)\boffice hours?\b',
    '(?i)\bexams?\b',
    '(?i)\btravel',
    '(?i)\bflights?\b',
    '(?i)\b(appointments?|appt|doctor|dentist|dental|medical|haircut)\b',
    '^(?i:dr)\.? '                # "Dr Smith", "Dr. appointment"
)

# A physical place: digits plus a street word ("123 Main St", "2 Elm Avenue") or a
# "City, ST 12345" ending. Online meeting links never count.
$script:StreetAddress = '(?i)(\b\d+[A-Za-z]?\s+(\w+\s){0,4}(st|street|ave|avenue|rd|road|blvd|boulevard|dr|drive|ln|lane|way|ct|court|pkwy|parkway|hwy|highway|pl|place|cir|circle|ter|terrace)\b)|(,\s*[A-Z]{2}\s+\d{5}\b)'
$script:OnlineLocation = '(?i)teams|zoom|webex|meet\.google|https?://'
function Get-FirmHoldPatterns {
    param([string[]]$Override = $null)
    if ($Override) { return , $Override }
    $cfgPath = Join-Path $env:APPDATA "outlook-suite\config.json"
    if (Test-Path -LiteralPath $cfgPath) {
        try {
            $cfg = Get-Content -LiteralPath $cfgPath -Raw | ConvertFrom-Json
            $list = $null
            if ($cfg.calendar -and $null -ne $cfg.calendar.firmHolds) { $list = $cfg.calendar.firmHolds }
            elseif ($null -ne $cfg.'calendar.firmHolds') { $list = $cfg.'calendar.firmHolds' }
            if ($null -ne $list) { return , @($list | ForEach-Object { [string]$_ }) }
        } catch { }
    }
    return , $script:DefaultFirmHolds
}

# People invited other than the organizer and the user, from display names (the table path has
# only the "Required/OptionalAttendees" name strings, which leave the organizer out; the item
# path has Recipients, which can include the user). One function so both paths always agree.
function Get-OtherAttendeeCount {
    param([string[]]$Names, [string]$Organizer = "")
    if ($null -eq $script:SelfName) {
        $script:SelfName = ""
        try { $script:SelfName = [string]$Namespace.CurrentUser.Name } catch { }
    }
    $skip = @($Organizer, $script:SelfName | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLowerInvariant() })
    $seen = @{}
    foreach ($n in $Names) {
        $k = ([string]$n).Trim().ToLowerInvariant()
        if ($k -and $skip -notcontains $k) { $seen[$k] = 1 }
    }
    return $seen.Count
}
$script:SelfName = $null

# meeting | hold | firm | allday (see Get-CalendarEvents). Stable API, also used by other skills:
#   Get-EventKind -AllDay -MeetingStatus -Attendees -Subject -Busy -Private -FirmPatterns [-Location]
# -Attendees is the number of OTHER people invited (organizer and the user excluded): use
# Get-OtherAttendeeCount. A meeting the user organizes with zero others is a hold.
# A no-attendee event is firm when it is Out of Office, private, matches a firm pattern, or
# takes place at a street address (an in-person appointment).
function Get-EventKind {
    param([bool]$AllDay, [int]$MeetingStatus, [int]$Attendees, [string]$Subject, [string]$Busy, [bool]$Private, [string[]]$FirmPatterns, [string]$Location = "")
    if ($AllDay) { return "allday" }
    if (-not ($MeetingStatus -eq 0 -or ($MeetingStatus -eq 1 -and $Attendees -eq 0))) { return "meeting" }
    if ($Busy -eq "Out of Office" -or $Private) { return "firm" }
    foreach ($p in $FirmPatterns) { if ($p -and $Subject -cmatch $p) { return "firm" } }
    if ($Location -and $Location -notmatch $script:OnlineLocation -and $Location -match $script:StreetAddress) { return "firm" }
    return "hold"
}

# Calendar item by EntryID (any mailbox) or by -Index in the next -Days (volatile).
# -OccurrenceStart picks one occurrence of a series (EntryID of a series opens the whole series).
function Get-CalendarItem {
    param($Namespace, [string]$EntryID = "", [int]$Index = 0, [int]$Days = 7, $OccurrenceStart = $null)
    $item = $null
    if ($EntryID) {
        $item = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID
    } elseif ($Index -gt 0) {
        Write-Host "Note: -Index is volatile (positions shift as events change). Prefer -EntryID from outlook-calendar-list.ps1." -ForegroundColor DarkGray
        $from = Get-Date
        $items = $Namespace.GetDefaultFolder(9).Items
        $items.Sort("[Start]")
        $items.IncludeRecurrences = $true
        $n = 0
        foreach ($e in $items.Restrict("[Start] >= '$(ConvertTo-RestrictDate $from)' AND [Start] <= '$(ConvertTo-RestrictDate $from.AddDays($Days))'")) {
            $n++
            if ($n -eq $Index) { return $e }
        }
        throw "No event at -Index $Index in the next $Days days ($n events). Use -EntryID from outlook-calendar-list.ps1."
    } else {
        throw "Provide -EntryID (preferred) or -Index to choose an event."
    }
    if ($OccurrenceStart) {
        if (-not $item.IsRecurring) { throw "-OccurrenceStart was given, but this event is not part of a series." }
        try { return $item.GetRecurrencePattern().GetOccurrence([datetime]$OccurrenceStart) }
        catch { throw "The series has no occurrence starting at $(([datetime]$OccurrenceStart).ToString('yyyy-MM-dd HH:mm')). Use the exact Start shown by outlook-calendar-list.ps1." }
    }
    return $item
}

# Record for one opened appointment item (used after create/update and in previews).
function ConvertTo-EventRecord {
    param($Item)
    $ms = [int]$Item.MeetingStatus
    $attendees = 0
    try { $attendees = Get-OtherAttendeeCount -Names @($Item.Recipients | ForEach-Object { [string]$_.Name }) -Organizer ([string]$Item.Organizer) } catch { }
    $kind = Get-EventKind -AllDay ([bool]$Item.AllDayEvent) -MeetingStatus $ms -Attendees $attendees -Subject ([string]$Item.Subject) -Busy $script:BusyNames[[int]$Item.BusyStatus] -Private ([int]$Item.Sensitivity -eq 2) -FirmPatterns (Get-FirmHoldPatterns) -Location ([string]$Item.Location)
    $isOccurrence = $false
    try { $isOccurrence = ([int]$Item.RecurrenceState -eq 2 -or [int]$Item.RecurrenceState -eq 3) } catch { }
    return [ordered]@{
        EntryID    = [string]$Item.EntryID
        Subject    = [string]$Item.Subject
        Start      = [datetime]$Item.Start
        End        = [datetime]$Item.End
        Kind       = $kind
        AllDay     = [bool]$Item.AllDayEvent
        Busy       = $script:BusyNames[[int]$Item.BusyStatus]
        Response   = $script:ResponseNames[[int]$Item.ResponseStatus]
        Organizer  = [string]$Item.Organizer
        Attendees  = $attendees
        Location   = [string]$Item.Location
        Recurring  = [bool]$Item.IsRecurring
        Occurrence = $isOccurrence
        Categories = [string]$Item.Categories
    }
}

# ---------------------------------------------------------------------------
# Display time zone (-TimeZone): changes how times are SHOWN, never how they are computed.
# ---------------------------------------------------------------------------

$script:DisplayTz = $null
$script:DisplayTzLabel = ""
$script:TimeZoneAliases = @{
    "eastern" = "Eastern Standard Time"; "et" = "Eastern Standard Time"; "est" = "Eastern Standard Time"; "edt" = "Eastern Standard Time"
    "central" = "Central Standard Time"; "ct" = "Central Standard Time"; "cst" = "Central Standard Time"; "cdt" = "Central Standard Time"
    "mountain" = "Mountain Standard Time"; "mt" = "Mountain Standard Time"; "arizona" = "US Mountain Standard Time"
    "pacific" = "Pacific Standard Time"; "pt" = "Pacific Standard Time"; "pst" = "Pacific Standard Time"; "pdt" = "Pacific Standard Time"
    "alaska" = "Alaskan Standard Time"; "hawaii" = "Hawaiian Standard Time"; "utc" = "UTC"; "gmt" = "UTC"
    # IANA names, for Windows PowerShell 5.1 (PowerShell 7 also resolves any IANA name itself)
    "america/new_york" = "Eastern Standard Time"; "america/chicago" = "Central Standard Time"; "america/denver" = "Mountain Standard Time"
    "america/phoenix" = "US Mountain Standard Time"; "america/los_angeles" = "Pacific Standard Time"; "america/anchorage" = "Alaskan Standard Time"
    "pacific/honolulu" = "Hawaiian Standard Time"; "europe/london" = "GMT Standard Time"; "europe/paris" = "Romance Standard Time"
    "europe/berlin" = "W. Europe Standard Time"; "europe/istanbul" = "Turkey Standard Time"; "asia/tokyo" = "Tokyo Standard Time"
    "asia/shanghai" = "China Standard Time"; "asia/kolkata" = "India Standard Time"; "australia/sydney" = "AUS Eastern Standard Time"
}

# Sets the display zone from "Eastern", "Pacific", "UTC", a Windows id ("Pacific Standard Time")
# or an IANA name ("America/Los_Angeles"). Empty = Outlook's local time.
function Set-DisplayTimeZone {
    param([string]$Name)
    if (-not $Name) { return }
    $id = $script:TimeZoneAliases[$Name.Trim().ToLowerInvariant()]
    if (-not $id) { $id = $Name.Trim() }
    try { $script:DisplayTz = [TimeZoneInfo]::FindSystemTimeZoneById($id) }
    catch { throw "Unknown time zone '$Name'. Use Eastern, Central, Mountain, Pacific, UTC, a Windows id like 'Pacific Standard Time', or an IANA name like America/Los_Angeles." }
    $script:DisplayTzLabel = if ($script:TimeZoneAliases.ContainsKey($Name.Trim().ToLowerInvariant()) -and $Name -notmatch "/") { "$((Get-Culture).TextInfo.ToTitleCase($Name.Trim().ToLowerInvariant())) time" } else { $script:DisplayTz.Id }
    if ($Name.Trim().Length -le 3) { $script:DisplayTzLabel = $Name.Trim().ToUpperInvariant() }
}

function ConvertTo-DisplayTime($Date) {
    if (-not $script:DisplayTz) { return [datetime]$Date }
    return [TimeZoneInfo]::ConvertTime([datetime]::SpecifyKind([datetime]$Date, "Unspecified"), [TimeZoneInfo]::Local, $script:DisplayTz)
}

# ISO 8601 in the display zone (with that zone's offset).
function Format-DisplayIso($Date) {
    $d = ConvertTo-DisplayTime $Date
    if (-not $script:DisplayTz) { return Format-MailDateIso $d }
    return ([datetimeoffset]::new($d, $script:DisplayTz.GetUtcOffset($d))).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

# Email-ready slot in a plain email style: "Mon, Oct 26, 10:30 to 11:30 a.m." (never a dash).
function Format-SlotLine {
    param($Start, $End)
    $s = ConvertTo-DisplayTime $Start; $e = ConvertTo-DisplayTime $End
    $mer = { param($d) if ($d.Hour -lt 12) { "a.m." } else { "p.m." } }
    $from = $s.ToString("h:mm", [Globalization.CultureInfo]::InvariantCulture)
    $to = $e.ToString("h:mm", [Globalization.CultureInfo]::InvariantCulture)
    $range = if ((& $mer $s) -eq (& $mer $e)) { "$from to $to $(& $mer $e)" } else { "$from $(& $mer $s) to $to $(& $mer $e)" }
    $line = "$($s.ToString('ddd, MMM d', [Globalization.CultureInfo]::InvariantCulture)), $range"
    if ($script:DisplayTzLabel) { $line += " $script:DisplayTzLabel" }
    return $line
}

# Text line for times: "Thu, Oct 22  9:00 AM - 10:00 AM" (all-day: "Thu, Oct 22  All day").
function Format-EventWhen {
    param($Rec)
    $s = [datetime]$Rec.Start; $e = [datetime]$Rec.End
    if ($Rec.AllDay) {
        $days = [int]($e.Date - $s.Date).TotalDays
        if ($days -gt 1) { return "$($s.ToString('ddd, MMM d')) - $($e.AddDays(-1).ToString('ddd, MMM d'))  All day" }
        return "$($s.ToString('ddd, MMM d'))  All day"
    }
    $s = ConvertTo-DisplayTime $s; $e = ConvertTo-DisplayTime $e
    $zone = if ($script:DisplayTzLabel) { "  ($script:DisplayTzLabel)" } else { "" }
    if ($s.Date -eq $e.Date) { return "$($s.ToString('ddd, MMM d'))  $($s.ToString('h:mm tt')) - $($e.ToString('h:mm tt'))$zone" }
    return "$($s.ToString('ddd, MMM d h:mm tt')) - $($e.ToString('ddd, MMM d h:mm tt'))$zone"
}

# JSON copy of an event record: dates as ISO 8601 with offset, internal fields dropped.
# With a display zone, timed records also get StartShown / EndShown in that zone.
function ConvertTo-EventJson {
    param($Rec)
    $o = [ordered]@{}
    foreach ($k in $Rec.Keys) {
        if ($k -eq "Events") { continue }
        $o[$k] = $Rec[$k]
        if ($Rec[$k] -is [datetime]) { $o[$k] = Format-MailDateIso $Rec[$k] }
    }
    if ($script:DisplayTz -and $Rec.Contains("Start") -and $Rec.Contains("End") -and -not $Rec.AllDay) {
        $o.StartShown = Format-DisplayIso $Rec.Start
        $o.EndShown = Format-DisplayIso $Rec.End
        $o.TimeZoneShown = $script:DisplayTz.Id
    }
    return $o
}

# Prints one event (list, previews and action results share this layout).
function Write-EventRecord {
    param($Rec, [string]$Prefix = "", [switch]$Detail)
    $tag = "[$($Rec.Kind)]"
    Write-Host "$Prefix$($Rec.Subject)  $tag" -ForegroundColor Yellow
    Write-Host "      When: $(Format-EventWhen $Rec)"
    if ($Rec.EntryID) { Write-Host "      EntryID: $($Rec.EntryID)" -ForegroundColor DarkGray }
    if ($Rec.Recurring) { Write-Host "      Series: yes. To act on this occurrence only, add -OccurrenceStart '$(([datetime]$Rec.Start).ToString('yyyy-MM-dd HH:mm'))'" -ForegroundColor DarkGray }
    if ($Rec.Location) { Write-Host "      Location: $($Rec.Location)" -ForegroundColor Gray }
    if ($Rec.Kind -eq "meeting") { Write-Host "      Organizer: $($Rec.Organizer)  ($($Rec.Attendees) others invited, you: $($Rec.Response))" -ForegroundColor Gray }
    if ($Rec.Busy -ne "Busy") { Write-Host "      Shows as: $($Rec.Busy)" -ForegroundColor Gray }
    if ($Rec.Cancelled) { Write-Host "      CANCELLED" -ForegroundColor Red }
    if ($Rec.ConflictsWith -and $Rec.ConflictsWith.Count -gt 0) { Write-Host "      CONFLICT with: $($Rec.ConflictsWith -join '; ')" -ForegroundColor Red }
    if ($Detail -and $Rec.Categories) { Write-Host "      Categories: $($Rec.Categories)" -ForegroundColor Gray }
}

# Preview of an irreversible calendar action, then the shared confirm gate.
function Confirm-CalendarAction {
    param([string]$Action, [string]$Title, $Item, [bool]$Confirmed, $Details = [ordered]@{}, [string]$What = "", [switch]$Json)
    if ($Confirmed -or -not (Test-ConfirmRequired -Action $Action)) { return $true }
    if (-not $What) { $What = $Title.ToLower() }
    if ($Json) { Set-CalendarPreviewJson -Action $Action -What $What -Rec (ConvertTo-EventRecord $Item) -Details $Details; return $false }
    Write-Host "=== $Title (PREVIEW) ===" -ForegroundColor Yellow
    Write-EventRecord (ConvertTo-EventRecord $Item) -Prefix "  "
    foreach ($k in $Details.Keys) { Write-Host "  ${k}: $($Details[$k])" -ForegroundColor Cyan }
    return (Approve-Action -Action $Action -Confirmed $false -What $What)
}

# -Json form of a preview: one object, Result = "preview", nothing done. It is STORED in
# $script:CalendarPreviewJson, never written here: Confirm-CalendarAction must return exactly
# one [bool] (pipeline output would make its result a truthy array and the action would
# run). The caller prints it with Write-CalendarPreview right before exiting.
function Set-CalendarPreviewJson {
    param([string]$Action, [string]$What, $Rec, $Details = [ordered]@{})
    $o = [ordered]@{ Action = $Action; Result = "preview"; Done = $false; What = $What; Hint = "Run the same command with -Confirm to do it." }
    $o.Event = ConvertTo-EventJson $Rec
    foreach ($k in $Details.Keys) { $o[$k] = $Details[$k] }
    $script:CalendarPreviewJson = $o
}
$script:CalendarPreviewJson = $null

# Prints the stored -Json preview (call at script level, after a gate returned not-$true).
function Write-CalendarPreview { if ($script:CalendarPreviewJson) { Write-Json $script:CalendarPreviewJson } }

# Attendee text ("Jane Doe", "jdoe", "jane@x.edu") -> resolved recipient info.
# Uses Resolve-Person from the contacts skill (_people.ps1) when it is installed,
# otherwise Outlook's own resolver (GAL, Contacts, recent addresses).
function Resolve-CalendarAttendee {
    param($Namespace, [string]$Text)
    $query = $Text
    $note = ""
    if ($Text -notmatch "@" -and (Get-Command Resolve-PersonSmtp -ErrorAction SilentlyContinue)) {
        try {
            $hits = @(Resolve-PersonSmtp -Namespace $Namespace -Query $Text)
            if ($hits.Count -eq 1) { $query = [string]$hits[0] }
            else { $note = "'$Text' matches more than one person ($($hits -join '; ')). Use the email address." }
        }
        catch { $note = $_.Exception.Message }   # ambiguous or not found: say so, never guess
    }
    $r = $Namespace.CreateRecipient($query)
    $ok = $false
    if (-not $note) { try { $ok = $r.Resolve() } catch { } }
    $smtp = ""
    if ($ok) {
        try { $u = $r.AddressEntry.GetExchangeUser(); if ($u) { $smtp = $u.PrimarySmtpAddress } } catch { }
        if (-not $smtp) { try { $smtp = $r.AddressEntry.PropertyAccessor.GetProperty($script:PR_SMTP_ADDRESS) } catch { } }
        if (-not $smtp) { $smtp = [string]$r.Address }
    }
    if (-not $ok -and -not $note) { $note = "Outlook could not resolve '$Text'. Use an email address." }
    $name = ""; if ($ok) { $name = [string]$r.Name }
    return [ordered]@{ Input = $Text; Resolved = [bool]$ok; Name = $name; Email = $smtp; Note = $note; Recipient = $r }
}

# "9", "09:00", "9:00 AM", "5 pm", "17:30" -> TimeSpan. Throws on anything else.
function ConvertTo-TimeOfDay {
    param([string]$Text, [string]$Name = "time")
    $t = $Text.Trim()
    if ($t -match '^\d{1,2}$') { $t = "$t`:00" }
    $formats = @("H:mm", "HH:mm", "h:mm tt", "h tt", "h:mmtt", "htt", "h:mm t", "h t")
    $d = [datetime]::MinValue
    foreach ($f in $formats) {
        if ([datetime]::TryParseExact($t.ToUpperInvariant().Replace(".", ""), $f, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.TimeOfDay }
    }
    throw "Cannot read $Name '$Text'. Use a time like 9, 09:00, 9:00 AM or 17:30."
}

# The contacts skill's people search, loaded once here at script scope when it is installed.
$script:PeopleHelpers = Join-Path $PSScriptRoot "..\..\outlook-contacts\scripts\_people.ps1"
if (-not (Get-Command Resolve-PersonSmtp -ErrorAction SilentlyContinue) -and (Test-Path -LiteralPath $script:PeopleHelpers)) { . $script:PeopleHelpers }

# Splits "a; b, c" into trimmed names.
function Split-AttendeeList([string]$Text) {
    return @($Text -split "[;,]" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}
