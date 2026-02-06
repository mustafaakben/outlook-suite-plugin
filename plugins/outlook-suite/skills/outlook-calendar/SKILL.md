---
name: outlook-calendar
description: "Manage Outlook calendar via PowerShell: list events, create/update/delete appointments, send meeting invites, respond to meetings, recurring events, check availability, and find rooms. Scripts auto-start Outlook when needed."
---

# Outlook Calendar

Manage Microsoft Outlook calendar events and meetings using PowerShell COM objects.

## Prerequisites

- Microsoft Outlook installed (scripts auto-start Outlook if it's not already running)
- PowerShell 5.1+ (included in Windows 10/11)
- Some features (rooms, free/busy) require Exchange/Office 365

## Events and Appointments

List, create, update, and delete calendar events.

### outlook-calendar-list.ps1

List upcoming calendar events with EntryIDs for use with action scripts.

**Required Parameters:** None (uses defaults)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | Number of days to look ahead (or ahead+behind with -Past) |
| `-Past` | switch | false | Include past events (shows -Days before and after today) |
| `-Limit` | int | 20 | Maximum number of events to display |

```powershell
# List upcoming events (next 7 days)
& "./scripts/outlook-calendar-list.ps1" -Days 7

# List events including past
& "./scripts/outlook-calendar-list.ps1" -Days 7 -Past

# List next 30 days, max 50 events
& "./scripts/outlook-calendar-list.ps1" -Days 30 -Limit 50
```

### outlook-calendar-create.ps1

Create a new appointment or meeting with attendees.

**Required Parameters:** `-Subject`, `-Start`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | — | Event subject/title (required) |
| `-Start` | datetime | — | Start date and time (required) |
| `-End` | datetime | Start+1hr | End date and time |
| `-Location` | string | "" | Event location |
| `-Body` | string | "" | Event description/body text |
| `-AllDay` | switch | false | Create as all-day event |
| `-Reminder` | int | 15 | Reminder minutes before event |
| `-Attendees` | string | "" | Semicolon-separated email addresses (sends invites) |

```powershell
# Create a simple appointment
& "./scripts/outlook-calendar-create.ps1" -Subject "Team Sync" -Start "2026-02-10 14:00"

# Create meeting with attendees (sends invites)
& "./scripts/outlook-calendar-create.ps1" -Subject "Project Review" -Start "2026-02-10 10:00" -End "2026-02-10 11:30" -Location "Conference Room A" -Attendees "john@example.com; jane@example.com"

# Create all-day event
& "./scripts/outlook-calendar-create.ps1" -Subject "Company Holiday" -Start "2026-02-15" -AllDay
```

### outlook-calendar-update.ps1

Update an existing calendar event's details.

**Required Parameters:** `-EntryID` or `-Index` (at least one)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | "" | Outlook EntryID for direct lookup (preferred) |
| `-Index` | int | 0 | Event position in upcoming list (fallback) |
| `-Days` | int | 7 | Days ahead to search when using -Index |
| `-Subject` | string | "" | New subject/title |
| `-Start` | datetime | — | New start time (adjusts end to maintain duration) |
| `-End` | datetime | — | New end time |
| `-Location` | string | "" | New location |
| `-Body` | string | "" | New body text |

```powershell
# Update event by EntryID (preferred)
& "./scripts/outlook-calendar-update.ps1" -EntryID "00000000..." -Subject "Updated Meeting" -Location "Room 202"

# Update event by index (fallback)
& "./scripts/outlook-calendar-update.ps1" -Index 1 -Subject "New Title"

# Reschedule event
& "./scripts/outlook-calendar-update.ps1" -EntryID "00000000..." -Start "2026-02-10 14:00"
```

### outlook-calendar-delete.ps1

Delete a calendar event or meeting.

**Required Parameters:** `-EntryID` or `-Index` (at least one)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | "" | Outlook EntryID for direct lookup (preferred) |
| `-Index` | int | 0 | Event position in upcoming list (fallback) |
| `-Days` | int | 7 | Days ahead to search when using -Index |

Supports PowerShell safety switches: `-WhatIf` and `-Confirm`.

```powershell
# Delete by EntryID (preferred)
& "./scripts/outlook-calendar-delete.ps1" -EntryID "00000000..."

# Delete by index (fallback)
& "./scripts/outlook-calendar-delete.ps1" -Index 1

# Delete event in wider range
& "./scripts/outlook-calendar-delete.ps1" -Index 3 -Days 14
```

## Meetings and Responses

Respond to meeting invitations, view attendee responses, and forward events.

### outlook-calendar-respond.ps1

Accept, tentatively accept, or decline a meeting invitation.

**Required Parameters:** `-Response`, plus `-EntryID` or `-Index`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | "" | Outlook EntryID for direct lookup (preferred) |
| `-Index` | int | 0 | Event position in upcoming list (fallback) |
| `-Response` | string | — | Response type: Accept, Tentative, or Decline (required) |
| `-Days` | int | 7 | Days ahead to search when using -Index |
| `-NoResponse` | switch | false | Save response locally without notifying organizer |

```powershell
# Accept a meeting (preferred)
& "./scripts/outlook-calendar-respond.ps1" -EntryID "00000000..." -Response Accept

# Tentatively accept by index
& "./scripts/outlook-calendar-respond.ps1" -Index 1 -Response Tentative

# Decline without sending response to organizer
& "./scripts/outlook-calendar-respond.ps1" -EntryID "00000000..." -Response Decline -NoResponse
```

### outlook-calendar-attendees.ps1

View meeting attendee list with response status (accepted/declined/tentative).

**Required Parameters:** `-EntryID` or `-Index` (at least one)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | "" | Outlook EntryID for direct lookup (preferred) |
| `-Index` | int | 0 | Event position in upcoming list (fallback) |
| `-Days` | int | 7 | Days ahead to search when using -Index |

```powershell
# View attendees by EntryID (preferred)
& "./scripts/outlook-calendar-attendees.ps1" -EntryID "00000000..."

# View attendees by index
& "./scripts/outlook-calendar-attendees.ps1" -Index 1

# Check meeting in wider range
& "./scripts/outlook-calendar-attendees.ps1" -Index 1 -Days 14
```

### outlook-calendar-forward.ps1

Forward a calendar event using Outlook `ForwardAsVcal` (vCal-style calendar attachment).

**Required Parameters:** `-To`, plus `-EntryID` or `-Index`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | "" | Outlook EntryID for direct lookup (preferred) |
| `-Index` | int | 0 | Event position in upcoming list (fallback) |
| `-To` | string | — | Recipient email address (required) |
| `-Body` | string | "" | Custom message to include |
| `-Days` | int | 7 | Days ahead to search when using -Index |
| `-Send` | switch | false | Send immediately (default: save as draft) |

```powershell
# Forward event as draft (preferred)
& "./scripts/outlook-calendar-forward.ps1" -EntryID "00000000..." -To "colleague@example.com" -Body "FYI"

# Forward and send immediately
& "./scripts/outlook-calendar-forward.ps1" -EntryID "00000000..." -To "colleague@example.com" -Send

# Forward by index
& "./scripts/outlook-calendar-forward.ps1" -Index 1 -To "colleague@example.com" -Body "FYI" -Send
```

## Recurring Events

Create events that repeat on a schedule.

### outlook-calendar-recurring.ps1

Create a recurring calendar event with various recurrence patterns.

**Required Parameters:** `-Subject`, `-Start`, `-Pattern`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | — | Event subject/title (required) |
| `-Start` | datetime | — | First occurrence date/time (required) |
| `-Pattern` | string | — | Recurrence type: Daily, Weekly, Monthly, Yearly (required) |
| `-Interval` | int | 1 | Repeat every N days/weeks/months/years |
| `-Duration` | int | 60 | Event duration in minutes |
| `-Location` | string | "" | Event location |
| `-Body` | string | "" | Event description |
| `-Attendees` | string | "" | Semicolon-separated emails (creates meeting series) |
| `-Occurrences` | int | 0 | End after N occurrences (0 = no limit) |
| `-EndByDate` | datetime | — | End recurrence on this date |
| `-Monday` ... `-Sunday` | switch | false | Weekly: which days to recur on |
| `-DayOfMonth` | int | 0 | Monthly: specific day number (1-31) |
| `-WeekNumber` | string | "" | Monthly/Yearly: First, Second, Third, Fourth, Last |
| `-DayOfWeek` | string | "" | Monthly/Yearly: Day, Weekday, WeekendDay, or specific day name |

For monthly "Nth weekday" recurrence, provide `-WeekNumber` and `-DayOfWeek` together.

```powershell
# Daily standup
& "./scripts/outlook-calendar-recurring.ps1" -Subject "Daily Standup" -Start "2026-02-10 09:00" -Pattern Daily -Duration 15

# Weekly on specific days
& "./scripts/outlook-calendar-recurring.ps1" -Subject "Team Meeting" -Start "2026-02-10 10:00" -Pattern Weekly -Monday -Wednesday -Friday

# Monthly on the 15th
& "./scripts/outlook-calendar-recurring.ps1" -Subject "Monthly Review" -Start "2026-02-10 14:00" -Pattern Monthly -DayOfMonth 15

# Second Tuesday of each month
& "./scripts/outlook-calendar-recurring.ps1" -Subject "Board Meeting" -Start "2026-02-10 09:00" -Pattern Monthly -WeekNumber Second -DayOfWeek Tuesday

# Limited occurrences
& "./scripts/outlook-calendar-recurring.ps1" -Subject "Sprint Planning" -Start "2026-02-10 09:00" -Pattern Weekly -Monday -Occurrences 10
```

## Availability and Rooms

Check attendee availability and find meeting rooms.

### outlook-freebusy.ps1

Check free/busy availability for one or more attendees. Requires Exchange/Office 365.

**Required Parameters:** `-Attendees`, `-Start`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Attendees` | string | — | Semicolon-separated email addresses (required) |
| `-Start` | datetime | — | Start of time range to check (required) |
| `-End` | datetime | Start+8hr | End of time range to check (must be later than `-Start`) |
| `-IntervalMinutes` | int | 30 | Time slot granularity in minutes (minimum 1) |

```powershell
# Check one person's availability
& "./scripts/outlook-freebusy.ps1" -Attendees "john@example.com" -Start "2026-02-10 09:00"

# Check multiple people for a full day
& "./scripts/outlook-freebusy.ps1" -Attendees "john@example.com; jane@example.com" -Start "2026-02-10 09:00" -End "2026-02-10 17:00"

# Finer 15-minute intervals
& "./scripts/outlook-freebusy.ps1" -Attendees "john@example.com" -Start "2026-02-10 09:00" -IntervalMinutes 15
```

### outlook-calendar-rooms.ps1

Find available conference rooms for a time slot. Requires Exchange/Office 365 with room mailboxes.

**Required Parameters:** `-Start`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Start` | datetime | — | Meeting start time (required) |
| `-End` | datetime | Start+1hr | Meeting end time (must be later than `-Start`) |
| `-Building` | string | "" | Filter rooms by building name |
| `-Capacity` | int | 0 | Currently unsupported in Outlook COM room-list data; accepted but ignored with warning |

```powershell
# Find available rooms
& "./scripts/outlook-calendar-rooms.ps1" -Start "2026-02-10 10:00"

# Rooms in specific building
& "./scripts/outlook-calendar-rooms.ps1" -Start "2026-02-10 10:00" -End "2026-02-10 11:00" -Building "HQ"
```

## Reference: Recurrence Patterns

| Pattern | Description | Key Options |
|---------|-------------|-------------|
| Daily | Every N days | `-Interval` |
| Weekly | Specific days of week, every N weeks | `-Monday` through `-Sunday`, `-Interval` |
| Monthly | Specific day or "Nth weekday" of month | `-DayOfMonth` or `-WeekNumber` + `-DayOfWeek` |
| Yearly | Same date each year or "Nth weekday" of month | `-WeekNumber` + `-DayOfWeek` |

## Reference: Busy Status Values

| Value | Meaning |
|-------|---------|
| Free | Time slot is available |
| Tentative | Tentatively booked |
| Busy | Confirmed busy |
| Out of Office | OOF/vacation |
| Working Elsewhere | Working from another location |
