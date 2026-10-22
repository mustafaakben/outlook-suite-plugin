---
name: outlook-calendar
description: "Manage the classic Outlook calendar on Windows through PowerShell COM: find free meeting times (alone or with other people, ranked open day > holds-only day > one-meeting day > busy day, 2-meeting daily cap, proposed as email-ready lines or HTML), week overview, list events tagged meeting / hold / firm (classes, office hours, travel: never booked over) / all-day with conflicts, create holds and meetings, turn a hold into an invite, update, delete (one occurrence or a series), respond to invitations, cancel with notice, read-only meeting prep (people, mail history, related mail, agenda, join link), attendees and responses, recurring series, free/busy grid, rooms. Fast (one range read per call), agent-friendly (EntryIDs, -Json, exit codes), with a confirm switch on everything that sends or deletes."
---

# Outlook Calendar

15 PowerShell scripts in `scripts/` that drive classic Outlook (running on this Windows PC) through COM. They share the mail skill's helpers (`scripts/_common.ps1` loads them) plus `scripts/_calendar.ps1`. Every script works standalone (`pwsh -File scripts/<name>.ps1 ...`) and under Windows PowerShell 5.1; through the warm host (`outlook-mail/scripts/outlook-call.ps1 calendar-list -Days 7`) repeated calls are faster.

## How to work

1. **Look before booking.** `outlook-calendar-week.ps1` shows what the week holds; `outlook-calendar-free.ps1` proposes slots, already ranked. Never answer "there is no time": the finder fills from the best tier down and, if nothing fits, says what to widen.
2. **Find first, then act by EntryID.** `outlook-calendar-list.ps1` (or `-Detail` on week) prints an `EntryID` for every event. `-Index` still works but positions shift; do not use it for actions.
3. **Series:** every occurrence of a recurring event carries the series' EntryID. To act on one occurrence, add `-OccurrenceStart "yyyy-MM-dd HH:mm"` (the Start the list printed). Without it, update and delete act on the whole series; the delete preview says "ENTIRE SERIES".
4. **Output contract:** success exits 0; failure prints `Error: <reason>` and exits 1. A gated action without confirmation prints `PREVIEW ONLY: ... was NOT done.`, changes nothing, and exits 0.
5. **Ask before anything that reaches other people** (invites, updates to attendees, responses, forwards) or deletes. When the user already said to go ahead, pass `-Confirm`. Creating or changing the user's own holds needs no confirmation.

## Meetings, holds, all-day

Every event has `Kind`:

| Kind | Meaning |
|---|---|
| `meeting` | other people are invited (a 1:1 the user organizes with one invitee is a meeting) |
| `hold` | no one else is invited: the user's own placeholder (Focus block, Planning, Prep block, Lunch, TODOs). Holds may be booked over when needed. |
| `firm` | no one else is invited, but it is a commitment and is never booked over: classes (course codes like `ABC-1234-A`, "Class"), office hours, exams, travel, flights, time shown as Out of Office, private events |
| `allday` | all-day event; one shown as Out of Office or Busy makes the day "Away" |

`Blocking` = takes time (Tentative, Busy, Out of Office) and is not cancelled or declined. Conflicts are flagged between overlapping meetings and firm events; holds never conflict.

Firm subjects are regex patterns (case-sensitive; prefix `(?i)` to ignore case). Replace the defaults with `-FirmHolds "pattern1","pattern2"` on list, week and free, or for good in `%APPDATA%\outlook-suite\config.json`: `{ "calendar": { "firmHolds": ["^[A-Z]{2,4}[- ]?\\d{3,4}\\b", "(?i)office hours"] } }`.

## Which script when

| Need | Script | Key parameters |
|---|---|---|
| Find a meeting time | `outlook-calendar-free.ps1` | `-Attendees "Jane Doe; john.doe@example.org" -Duration 30 -From -To \| -Days 7 -WorkStart 9 -WorkEnd "5 PM" -Step 30 -Buffer 0 -MaxPerDay 2 -PerDay 2 -Top 12 -IncludeWeekends -FirmHolds -TimeZone -Json` |
| Email-ready time options | `outlook-calendar-free.ps1 -Propose` | `-Count 3 [-Html] [-TimeZone Pacific] [-Json]` plus the finder's parameters |
| What does the week look like | `outlook-calendar-week.ps1` | `-From <date> \| -Next -Weeks 1 -MaxPerDay 2 -IncludeWeekends -Detail -FirmHolds -Json` |
| List events | `outlook-calendar-list.ps1` | `-Days 7 \| -Past \| -From -To, -Kind meeting\|hold\|firm\|allday -ConflictsOnly -FirmHolds -TimeZone -Limit 20 -Json` |
| New hold or meeting | `outlook-calendar-create.ps1` | `-Subject -Start (-End \| -Duration) -AllDay -Location -Body -BusyStatus Free\|Tentative\|Busy\|OutOfOffice\|WorkingElsewhere -Categories -Reminder 15 \| -NoReminder`; meeting: `-Attendees -Optional -Confirm` |
| Prepare for a meeting (read-only) | `outlook-calendar-prep.ps1` | `-Next \| -EntryID [-OccurrenceStart]`, `-Days 90 -MaxAttendees 25 -Related 10 -TimeZone -Json` |
| Cancel a meeting the user organizes, notify everyone | `outlook-calendar-cancel.ps1` | `-EntryID [-OccurrenceStart] [-Message "..."] -Confirm` |
| Turn a hold into an invite | `outlook-calendar-hold-to-invite.ps1` | `-EntryID -Attendees [-Optional -Subject -Location -Body] -Confirm` |
| Change an event | `outlook-calendar-update.ps1` | `-EntryID [-OccurrenceStart] -Subject -Start -End -Location -Body -BusyStatus -Categories -Reminder \| -NoReminder [-SendUpdate -Confirm]` |
| Delete | `outlook-calendar-delete.ps1` | `-EntryID [-OccurrenceStart] -Confirm` |
| Accept / tentative / decline | `outlook-calendar-respond.ps1` | `-EntryID [-OccurrenceStart] -Response Accept\|Tentative\|Decline [-NoResponse] -Confirm` |
| Who is invited, who answered | `outlook-calendar-attendees.ps1` | `-EntryID [-OccurrenceStart] -Json` |
| Forward an event | `outlook-calendar-forward.ps1` | `-EntryID -To [-Body]` (draft) or `-Send -Confirm` |
| Recurring hold or meeting series | `outlook-calendar-recurring.ps1` | `-Subject -Start -Pattern Daily\|Weekly\|Monthly\|Yearly -Interval -Duration -Monday..-Sunday -DayOfMonth -WeekNumber -DayOfWeek -Occurrences \| -EndByDate -BusyStatus -Categories -NoReminder`; meeting: `-Attendees -Confirm` |
| Free/busy grid for people | `outlook-freebusy.ps1` | `-Attendees -Start [-End] -IntervalMinutes 30 -Json` |
| Free rooms | `outlook-calendar-rooms.ps1` | `-Start [-End \| -Duration 60] -Building -Limit -Json` |

## Finding a time (`outlook-calendar-free.ps1`)

The scheduling rule, built in:

1. It summarizes the whole week first (Monday to Friday around the range): meetings, holds and firm events per day, all-day events, away days, days over the cap, conflicts.
2. Each day gets a tier, and slots are offered best tier first:
   - **Open day**: no meetings and no holds.
   - **Holds-only day**: no meetings, only holds and firm events; the slot may cover a hold (named in the output), never a firm event.
   - **One-meeting day**: one meeting booked, so the new one is the second, the daily maximum.
   - **Busy day**: the cap (`-MaxPerDay`, default 2) is already reached. Offered only when no better day has room, and labeled.
3. Within a day, slots covering the fewest holds and no tentative attendees come first; picks are spread at least 2 hours apart.
4. Other people come from Exchange free/busy (names resolve through the contacts skill's people search when it is installed, otherwise Outlook's resolver). Busy or Out of Office blocks; tentative is allowed and noted. An ambiguous or unknown name is reported, never guessed. Anyone whose free/busy could not be read (unresolved, not shared, or no data that far ahead) is listed as **NOT CHECKED** in the header and on each slot: tell the user those times were not checked against that person.
5. Meetings and firm events block; holds do not. Away days (all-day Out of Office or Busy) get no slots. Today starts at the next step after now, never in the past; a `-From` in the past starts today.
6. If nothing fits, the search widens by a week (twice at most) and then lists each skipped day with its reason and the options to try. It never just says there is no time.

**`-Propose`** prints the best `-Count` slots (default 3, best tier first, different days first) as email-ready lines in a plain email style, e.g. `Mon, Oct 26, 10:30 to 11:30 a.m.` (always "to", never a dash). `-Html` prints the same as one Outlook-native block (`<div>` in Aptos 11pt black, `class="elementToProof"`, a `<ul>` of slots) to paste into a draft's HTML. `-Json` returns the slots with `Text`, plus `Html`.

**`-TimeZone`** (list, free, propose) changes only how times are SHOWN: `Eastern`, `Central`, `Mountain`, `Pacific`, `UTC`, a Windows id (`Pacific Standard Time`) or an IANA name (`America/Los_Angeles`). Text lines carry the zone name; JSON adds `StartShown`/`EndShown` with that zone's offset next to the original `Start`/`End`.

Present the result to the user grouped under the tier names, then book with `outlook-calendar-create.ps1` (a hold) or, with attendees, after they pick a slot, with `-Confirm`.

```powershell
& "./scripts/outlook-calendar-free.ps1" -Attendees "Jane Doe; john@example.com" -Duration 60
& "./scripts/outlook-calendar-free.ps1" -From "2026-11-02" -To "2026-11-06" -Duration 45 -Buffer 15
& "./scripts/outlook-calendar-week.ps1" -Next -Detail
```

## Meeting prep (`outlook-calendar-prep.ps1`)

Read-only brief for one meeting (`-Next` picks the user's next real meeting: other people invited, not cancelled, not declined), in about 2 to 3 s:

- **Invite**: when, where, organizer, the join link (Teams, Zoom, Meet, Webex) and the agenda, which is the invite's own text with the Teams/Zoom boilerplate and links removed.
- **People**: everyone invited except the user (the organizer too, when it is someone else), with title, department and office from the directory, their response, and the mail exchanged in the last `-Days`: received from them (every received folder, matched by address), sent to them (Sent Items), and the last date. Sent counts match the person's display name in the To/CC line, so two people with the same name can overcount.
- **Related mail**: messages whose subjects share at least two keywords with the meeting subject (one if the subject has only one), with subject, date and sender or recipients only. Bodies are never read.

## Cancelling (`outlook-calendar-cancel.ps1`)

Only for meetings the user organizes: it marks the meeting cancelled and sends the cancellation to every attendee (Outlook's "Cancel meeting"), with `-Message` at the top of the notice, then removes the user's copy. A series EntryID cancels the whole series (the preview says "ENTIRE SERIES"); `-OccurrenceStart` cancels one occurrence. Gated: without `-Confirm` it previews who would be notified and changes nothing. For a meeting someone else organized, decline with `outlook-calendar-respond.ps1`; for a hold, use `outlook-calendar-delete.ps1`.

## Examples

```powershell
# Events of the next two weeks, meetings only, as JSON
& "./scripts/outlook-calendar-list.ps1" -Days 14 -Kind meeting -Json

# A tentative hold without reminder
& "./scripts/outlook-calendar-create.ps1" -Subject "Backup slot" -Start "2026-10-28 17:00" -Duration 30 -BusyStatus Tentative -NoReminder

# Turn that hold into a Zoom meeting (preview first, then -Confirm)
& "./scripts/outlook-calendar-hold-to-invite.ps1" -EntryID "00000000..." -Attendees "jane@example.com" -Location "Zoom"

# Move one occurrence of a series by an hour
& "./scripts/outlook-calendar-update.ps1" -EntryID "00000000..." -OccurrenceStart "2026-11-03 09:00" -Start "2026-11-03 10:00"

# Decline one occurrence and tell the organizer
& "./scripts/outlook-calendar-respond.ps1" -EntryID "00000000..." -OccurrenceStart "2026-11-03 09:00" -Response Decline -Confirm
```

## Confirmation

Sending invitations (create or recurring with attendees, hold-to-invite), `-SendUpdate`, cancelling, forwarding with `-Send`, responding and deleting are gated: without `-Confirm` they print a preview and do nothing. The gate is shared with the mail skill (actions `send`, `respond`, `delete`); see the mail skill's Confirmation section for turning it off per shell (`OUTLOOK_SUITE_REQUIRE_CONFIRM=0`) or per user (`%APPDATA%\outlook-suite\config.json`).

With `-Json`, an unconfirmed action prints one JSON object with `"Result": "preview"` and `"Done": false`.

**Changed in 3.0 (was immediate before):** create and recurring with `-Attendees`, `outlook-calendar-respond.ps1`, and `outlook-calendar-forward.ps1 -Send` now preview unless `-Confirm`.

Notes: deleting a meeting does not notify anyone (decline with `outlook-calendar-respond.ps1` instead). Updating a meeting you organize saves it locally unless you add `-SendUpdate -Confirm`. `respond -NoResponse` changes only the user's copy and never calls Outlook's `Respond()` (Microsoft's documentation is contradictory about whether that call sends by itself): Accept turns a Free or Tentative show-as into Busy, Tentative turns Free or Busy into Tentative (Out of Office and Working Elsewhere are never touched; the preview shows old and new), and Decline removes it from the user's calendar without telling the organizer. Acting on a series EntryID without `-OccurrenceStart` responds for the ENTIRE SERIES, which the preview says.

## Reference

| Recurrence | Options |
|---|---|
| Daily | `-Interval` |
| Weekly | `-Monday` .. `-Sunday`, `-Interval` |
| Monthly | `-DayOfMonth`, or `-WeekNumber First\|Second\|Third\|Fourth\|Last` + `-DayOfWeek` |
| Yearly | same date, or `-WeekNumber` + `-DayOfWeek` |

Free/busy characters (JSON `Status`): 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere. Strings start at midnight of the start date, one character per interval.

Rooms come from the organization's "All Rooms" list; Outlook does not expose room capacity.
