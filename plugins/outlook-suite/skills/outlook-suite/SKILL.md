---
name: outlook-suite
description: "Complete Microsoft Outlook automation via PowerShell COM objects on Windows. Covers email (triage, Copilot priority and suggested replies, read, search, draft, send, reply, forward, move, delete, flag, categorize, rules, attachments, out-of-office), calendar (week overview, find meeting times ranked by open, holds-only, one-meeting and busy days, email-ready time proposals, meetings vs holds vs firm events, hold-to-invite, read-only meeting prep, cancel with notice, create/update/delete, respond, recurring, free/busy, rooms), people and contacts (find anyone across Contacts, the directory and recent mail, who-is profiles, org chart, department and title search, distribution list members, dedupe, vCard/CSV import, contact CRUD), tasks (daily or weekly agenda, overdue triage in bulk, tasks from email, recurring tasks, due views with flagged email), and notes (full-text search, color and category filters). Use when the user mentions Outlook, email, calendar, meetings, scheduling, finding a time, people, contacts, tasks, to-dos, agenda, notes, or any mail/scheduling operation on Windows."
---

# Outlook Suite

Automates Microsoft Outlook across email, calendar, people and contacts, tasks and notes using PowerShell COM objects: 94 scripts in 5 modules (mail 56, calendar 15, contacts 10, tasks 9, notes 4). Version 3.0 brings every module to the same standard: one shared connection and output layer, fast table queries, `-Json` on every script, `Error: ...` with exit code 1 on failure, Windows PowerShell 5.1 and PowerShell 7, and the same confirmation gate for every irreversible action.

## Prerequisites

- Classic Microsoft Outlook for Windows (scripts start it with its window if it is not running).
- Windows PowerShell 5.1 or PowerShell 7. The warm host (`outlook-call.ps1`) needs PowerShell 7 (`pwsh`); under 5.1 it runs scripts directly.
- Exchange Server or Microsoft 365 for some features (free/busy of others, rooms, the directory, task assignment, distribution lists).

## Safety Rules

1. **Confirmation gates.** Send, delete, invite, respond, assign, bulk changes and every other irreversible action in every module only preview unless `-Confirm` is passed (the user can switch this off per shell or per action; see "Confirmation" in [outlook-mail/SKILL.md](../outlook-mail/SKILL.md)). With `-Json`, a preview is a JSON object with `Result` "preview". Ask the user before passing `-Confirm` unless they already said to go ahead.
2. **Draft-first for compose.** Create as a draft, show the user, then send only after approval.
3. **Verify targets.** When acting on a specific item, confirm the subject or person with the user before modifying.
4. **Prefer EntryID.** Use EntryID over `-Index` when targeting items (index numbers shift as items change).

## Fast path: the warm host for every skill

`outlook-mail/scripts/outlook-call.ps1` runs any script of any module through a warm Outlook session (PowerShell 7), so repeated calls skip the start-up cost. Name the script without the `outlook-` prefix; it is found in whichever module holds it. Arguments, output and exit codes are exactly those of a direct run.

```powershell
& "../outlook-mail/scripts/outlook-call.ps1" triage
& "../outlook-mail/scripts/outlook-call.ps1" calendar-free -Attendees "jane.doe@example.org" -Duration 30 -Propose
& "../outlook-mail/scripts/outlook-call.ps1" people "jane doe" -Profile
& "../outlook-mail/scripts/outlook-call.ps1" agenda -Week
& "../outlook-mail/scripts/outlook-call.ps1" notes-list -Search "budget" -Json
```

The host starts on first use and exits after 30 idle minutes. The first call of each script in a new host also loads that script (a few seconds once); later calls take well under a second.

## Item Targeting Pattern

All modules use **EntryID** to target items. An EntryID identifies one item, but it can change when the item moves to another folder or mailbox; scripts that move or copy print the new one.

1. Run a listing or search script to discover items (for mail, prefer `outlook-find.ps1`; for people, `outlook-people.ps1`).
2. Use the EntryID from the results in the action scripts.

```powershell
powershell -File outlook-find.ps1 -From "jane.doe@example.org" -Days 3
powershell -File outlook-read-body.ps1 -EntryID "00000000..."
```

## Modules

Each module has its own SKILL.md with full parameter tables and examples. **Load the relevant module documentation based on what the user needs.**

### Email — [outlook-mail/SKILL.md](../outlook-mail/SKILL.md)

56 scripts:

- **Triage**: `outlook-triage.ps1` (unread untagged mail, mail by tag, mail not replied to) with Copilot's priority; `outlook-priority.ps1` ranks all received mail by Copilot priority; `outlook-smart-reply.ps1` turns Copilot's suggested replies into drafts.
- **Find and read**: `outlook-find.ps1` for EntryIDs with rich filters, whole-mailbox search, full bodies with numbered link citations (`[URL-001]`), threads, headers, attachments. `outlook-everything.ps1` searches mail, calendar, contacts, tasks, notes and local files at once through the Windows Search index; `outlook-attachment-read.ps1` returns the text of Word, Excel, PowerPoint, PDF and `.msg` attachments (or of a local file).
- **Compose and send**: draft, update, send a reviewed draft, send-as, HTML reply and forward in the thread, voting, receipts, scheduled send.
- **Organize**: move, copy, delete, bulk actions, read state, flags, importance, sensitivity, categories, folders and rules.
- **Account**: accounts and stores, out-of-office, junk lists, signatures, reminders, save-as, print, `.msg` import.

### Calendar — [outlook-calendar/SKILL.md](../outlook-calendar/SKILL.md)

15 scripts. New in 3.0:

- **Event kinds** in every list, week and finder result: `meeting` (others invited, including a 1:1 the user organizes), `hold` (no one else invited: the user's own placeholder, which may be booked over), `firm` (no one else invited but a commitment that is never booked over: classes, office hours, exams, travel, Out of Office, private events) and `allday`.
- **Week overview**: `outlook-calendar-week.ps1` shows per day the meetings, holds, firm events, all-day events, conflicts and the day's tier.
- **Find a time**: `outlook-calendar-free.ps1` finds slots for the user alone or with attendees, ranked by tier (open day, holds-only day, one-meeting day, busy day) with a daily cap of 2 meetings, and never answers "no time". `-Propose` prints the best options as email-ready lines (or Outlook-native HTML); `-TimeZone` shows times in another zone. `-Tier`, `-IncludeBusy`, `-Spread`, `-Prefer Morning|Afternoon` and `-SkipDays` steer which days and times are offered.
- **Hold to invite**: `outlook-calendar-hold-to-invite.ps1` turns a hold into a meeting with attendees (previews until `-Confirm`).
- **Meeting prep**: `outlook-calendar-prep.ps1 -Next` (or `-EntryID`) is a read-only brief: join link and agenda from the invite, each attendee's title and department, how much mail the user exchanged with each, and related mail by subject (subjects and dates only, never bodies).
- **Cancel with notice**: `outlook-calendar-cancel.ps1` cancels a meeting the user organizes and notifies every attendee, with an optional message; it previews until `-Confirm`.
- **Respond without notifying**: `outlook-calendar-respond.ps1 -NoResponse` changes only the user's copy and never calls Outlook's `Respond()`. Accept turns a Free or Tentative show-as into Busy, Tentative turns Free or Busy into Tentative, Out of Office is left alone, and Decline removes the event. Responding to a series without `-OccurrenceStart` is labeled ENTIRE SERIES.
- All 10 earlier scripts (list, create, update, delete, respond, attendees, forward, recurring, free/busy, rooms) migrated to the suite standard.

### People and Contacts — [outlook-contacts/SKILL.md](../outlook-contacts/SKILL.md)

10 scripts. New in 3.0:

- **Find anyone**: `outlook-people.ps1 "<name, partial or email>"` searches Contacts, the directory (Global Address List) and recent mail, and returns contact points: emails, phones, title, department, office, manager.
- **Who is this**: `-Profile` adds mail exchanged and meetings together.
- **Org chart**: `-Org` shows the manager chain, direct reports and peers.
- **Directory search**: `-Department` and `-Title` ("who in IT", "deans"), from a local index built once.
- **Groups**: `-Members` lists a distribution list or contact group.
- **Hygiene**: `outlook-contacts-dedupe.ps1` reports duplicates and merges same-email groups; `outlook-contacts-import.ps1` imports vCard or CSV with a preview first.
- The 7 earlier contact scripts migrated to the suite standard.

### Tasks — [outlook-tasks/SKILL.md](../outlook-tasks/SKILL.md)

9 scripts. New in 3.0:

- **Agenda**: `outlook-agenda.ps1` is a read-only brief for a day or a week: meetings, firm events and holds (the calendar skill's classification), tasks and flagged email due and overdue, and unread mail Copilot rates High, each with a count.
- **Overdue triage**: `outlook-tasks-triage.ps1` groups overdue tasks by age and category and reschedules, completes or clears due dates in bulk, behind the bulk confirmation.
- **Tasks from email**: `outlook-tasks-from-mail.ps1` creates a task linked back to an email (the email is never changed).
- **Recurring tasks**: `outlook-tasks-create.ps1 -Recurrence Daily|Weekly|Monthly|Yearly -Interval -Until`, with reminders; the list shows how each task repeats.
- **Due views**: `outlook-tasks-list.ps1 -Due Today|Overdue|Week|NoDate`, `-Flagged` adds flagged email as to-dos, plus `-Search` and `-Category`.
- The 6 earlier task scripts migrated to the suite standard.

### Notes — [outlook-notes/SKILL.md](../outlook-notes/SKILL.md)

4 scripts. New in 3.0: `-Search` matches the whole note text, `-Category` and `-Color` filter, `-Json` everywhere, and `outlook-notes-read.ps1` lists links as `[URL-001]`. All 4 scripts migrated to the suite standard.

## Script Quick Reference

> **IMPORTANT FOR AI AGENTS:** This index gives you the correct script name for each operation. Before running any script, you **MUST** read the relevant module SKILL.md (linked in each heading) for full parameter documentation, required vs optional parameters, and usage examples. Do NOT guess parameters from the script name alone.

### Email — [outlook-mail/SKILL.md](../outlook-mail/SKILL.md)

All 56 mail scripts, with their key parameters, are in the **"Which script when"** table in [outlook-mail/SKILL.md](../outlook-mail/SKILL.md). Start with `outlook-triage.ps1` (what needs attention), `outlook-priority.ps1` (mail ranked by Copilot priority), `outlook-smart-reply.ps1` (Copilot's suggested replies), `outlook-find.ps1` (EntryIDs), `outlook-read-body.ps1` / `outlook-conversation.ps1` (read), `outlook-reply.ps1` / `outlook-draft.ps1` (compose as draft), `outlook-send-draft.ps1 -Confirm` (send after approval).

### Calendar — [outlook-calendar/SKILL.md](../outlook-calendar/SKILL.md)

- `outlook-calendar-week.ps1`: the week at a glance, per day meetings, holds, firm and all-day events, conflicts and tier
- `outlook-calendar-free.ps1`: find meeting times, ranked by tier; `-Propose` for email-ready options
- `outlook-calendar-list.ps1`: events with EntryIDs and their kind
- `outlook-calendar-create.ps1` / `outlook-calendar-update.ps1` / `outlook-calendar-delete.ps1`: appointments and meetings
- `outlook-calendar-hold-to-invite.ps1`: turn a hold into a meeting
- `outlook-calendar-prep.ps1`: read-only brief for the next (or a given) meeting: people, mail history, related mail, agenda, join link
- `outlook-calendar-cancel.ps1`: cancel a meeting you organize and notify attendees (`-Confirm`)
- `outlook-calendar-respond.ps1` / `outlook-calendar-attendees.ps1` / `outlook-calendar-forward.ps1`: meeting responses (`-NoResponse` for your copy only), attendees, forward as .ics
- `outlook-calendar-recurring.ps1`: recurring events
- `outlook-freebusy.ps1` / `outlook-calendar-rooms.ps1`: free/busy of others, available rooms

### People and Contacts — [outlook-contacts/SKILL.md](../outlook-contacts/SKILL.md)

- `outlook-people.ps1`: find anyone and their contact points; `-Profile`, `-Org`, `-Department` / `-Title`, `-Members`
- `outlook-contacts-list.ps1` / `outlook-contacts-search.ps1`: browse and search your Contacts
- `outlook-contacts-create.ps1` / `outlook-contacts-update.ps1` / `outlook-contacts-delete.ps1`: manage contacts
- `outlook-contacts-dedupe.ps1` / `outlook-contacts-import.ps1`: duplicates, vCard/CSV import
- `outlook-contacts-photo.ps1` / `outlook-contacts-export.ps1`: photos, vCard export

### Tasks — [outlook-tasks/SKILL.md](../outlook-tasks/SKILL.md)

- `outlook-agenda.ps1`: today's (or the week's) brief across calendar, tasks, flagged email and priority mail
- `outlook-tasks-list.ps1`: tasks by due view, status, category; `-Flagged` adds flagged email
- `outlook-tasks-triage.ps1`: overdue tasks by age and category, bulk reschedule, complete or clear
- `outlook-tasks-create.ps1` / `outlook-tasks-from-mail.ps1`: new task, recurring task, task from an email
- `outlook-tasks-update.ps1` / `outlook-tasks-complete.ps1` / `outlook-tasks-delete.ps1`: manage tasks
- `outlook-tasks-assign.ps1`: assign a task to someone (sends a task request)

### Notes — [outlook-notes/SKILL.md](../outlook-notes/SKILL.md)

- `outlook-notes-list.ps1`: notes with full-text search, color and category filters
- `outlook-notes-read.ps1`: one note in full
- `outlook-notes-create.ps1` / `outlook-notes-delete.ps1`: create and delete notes

> **AI AGENTS:** Read the relevant module SKILL.md linked in each section heading above for full parameter documentation and examples. Do NOT guess parameter names or assume defaults; always check the module docs first.

## How to Use

1. **Identify the domain**: which module handles the user's request.
2. **Load the module**: read the module SKILL.md linked above.
3. **Find the target**: for actions on existing items, run a listing or search script first to get the EntryID.
4. **Execute**: run the script with the parameters documented in the module SKILL.md, through `outlook-call.ps1` when you will make several calls.
5. **Confirm irreversible actions**: show the preview to the user, then add `-Confirm` only with their approval.

## Script Execution

All scripts are PowerShell files in each module's `scripts/` directory:

```powershell
powershell -File <module>/scripts/<script-name>.ps1 -Param1 value1 -Param2 value2
pwsh -File outlook-mail/scripts/outlook-call.ps1 <script-name-without-outlook-> -Param1 value1
```

Scripts handle the Outlook connection themselves and reuse the warm host's session when they run inside it.
