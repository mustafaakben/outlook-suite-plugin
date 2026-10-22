# Changelog

All notable changes to the Outlook Suite plugin. Versions follow [Semantic Versioning](https://semver.org/). The format is based on [Keep a Changelog](https://keepachangelog.com/).

## [3.0.0] - 2026-10-22

Calendar, people and contacts, tasks and notes are brought to the mail standard and extended for daily productivity. The suite now has 92 scripts: mail 54, calendar 15, contacts 10, tasks 9, notes 4.

### Added
- Calendar: every event is classified as `meeting` (others invited), `hold` (the user's own placeholder, may be booked over), `firm` (no one else invited but never booked over: classes, office hours, exams, travel, Out of Office, private events) or `allday`.
- Calendar: `outlook-calendar-week.ps1`, a week overview with meetings, holds, firm and all-day events, conflicts and each day's tier.
- Calendar: `outlook-calendar-free.ps1` finds meeting times for the user alone or with attendees, ranked by tier (open day, holds-only day, one-meeting day, busy day) with a daily cap of 2 meetings; `-Propose` prints email-ready options (text or Outlook-native HTML) and `-TimeZone` shows times in another zone.
- Calendar: `outlook-calendar-hold-to-invite.ps1` turns a hold into a meeting with attendees.
- Calendar: `outlook-calendar-prep.ps1`, a read-only brief for the next (`-Next`) or a given meeting: join link and agenda from the invite, each attendee's title and department, the mail exchanged with each (count and last date), and related mail by subject keywords (subjects and dates only, never bodies).
- Calendar: `outlook-calendar-cancel.ps1` cancels a meeting the user organizes and notifies every attendee, with an optional message; one occurrence or the whole series; previews until `-Confirm`.
- People: `outlook-people.ps1` finds anyone by name, partial name or email across Contacts, the directory and recent mail, with contact points (emails, phones, title, department, office, manager); `-Profile` (mail exchanged and meetings together), `-Org` (manager chain, direct reports, peers), `-Department` / `-Title` directory search, `-Members` (distribution lists and contact groups).
- Contacts: `outlook-contacts-dedupe.ps1` (duplicate report, merge of same-email groups) and `outlook-contacts-import.ps1` (vCard and CSV, preview first, `-Update`).
- Tasks: `outlook-agenda.ps1`, a read-only day or week brief: meetings, firm events and holds, tasks and flagged email due and overdue, and unread mail rated High by Copilot, with counts per section.
- Tasks: `outlook-tasks-triage.ps1` groups overdue tasks by age and category and reschedules, completes or clears due dates in bulk.
- Tasks: `outlook-tasks-from-mail.ps1` creates a task linked back to an email (sender, date, an `outlook:` link and the email's Message-ID), without changing the email.
- Tasks: recurring tasks (`-Recurrence Daily|Weekly|Monthly|Yearly`, `-Interval`, `-Until`) and reminders; the list shows how a task repeats.
- Tasks: list views `-Due Today|Overdue|Week|NoDate`, `-Flagged` (flagged email as to-dos), `-Search`, `-Category`.
- Notes: full-text `-Search`, `-Category` filter, `-Json` on every script, links as `[URL-001]` in `outlook-notes-read.ps1`.
- `-Json` and `Error: ...` with exit code 1 on every script of every module.
- Tests: `Test-CalendarCore.ps1`, `Test-CalendarFree.ps1`, `Test-CalendarPrep.ps1`, `Test-PeopleContacts.ps1`, `Test-Tasks.ps1`, `Test-TasksTriage.ps1`, `Test-Agenda.ps1`, `Test-Notes.ps1`, `Test-SuiteHost.ps1` (every module through the warm host) and `Test-SuiteGates.ps1` (static confirmation-gate guard).

### Changed
- All calendar, contacts, tasks and notes scripts share the mail skill's helpers (`_common.ps1`, loaded through a one-line bridge in each module) and read lists with fast table queries.
- `outlook-call.ps1` routes every module through the warm host: name any script without the `outlook-` prefix.
- Delete, send, respond, assign, import, merge and bulk actions in every module use the same confirmation gate as mail; with `-Json` a preview is a JSON object.
- The agenda uses the calendar skill's classification, so both always agree.
- Calendar: `outlook-calendar-respond.ps1 -NoResponse` no longer calls Outlook's `Respond()` (Microsoft's documentation is contradictory about whether it sends by itself). It changes only the user's copy: Accept turns a Free or Tentative show-as into Busy, Tentative turns Free or Busy into Tentative, Out of Office and Working Elsewhere are left alone, and Decline removes the event. The preview shows the old and new show-as, and says ENTIRE SERIES when no `-OccurrenceStart` is given.
- Calendar: create and recurring with attendees, respond, and forward with `-Send` now preview unless `-Confirm` (they acted immediately before).

### Fixed
- A new sticky note's color was never saved.
- `-StripLinks $false` could not be passed to the notes scripts through `pwsh -File`.
- Task subject search treated `[` and `]` as wildcards.
- Due dates saved by Outlook on the web, mobile and To Do (00:00 UTC) showed one day early, so a task due today appeared overdue.
- Calendar: a one-to-one meeting the user organizes was classified as a hold.
- Calendar: attendee free/busy was misaligned by one hour after a daylight-saving change (and the old free/busy grid started at the wrong slot). Exchange's free/busy string runs on elapsed time from midnight of its start date; it is now indexed by UTC difference, and people with no data that far ahead are marked NOT CHECKED.
- Calendar: `outlook-calendar-update.ps1` validated after editing, which could leave unsaved edits behind; it now validates first, and says "saved, update NOT sent" if the send fails after saving.
- Scripts started outside the warm host and inside it now show the same output.

### Security
- A preview path that writes JSON could make a confirmation helper return more than one value, which PowerShell treats as true, so an action could run without `-Confirm`. Found during testing of this release and fixed before release in every module. Confirmation helpers now return a single boolean and the caller prints the preview.
- Every call site is fail-closed: an action runs only when the gate returns exactly `$true`.
- `tests/Test-SuiteGates.ps1` checks every script statically for this (helpers return only a boolean, every call site is fail-closed) under PowerShell 7 and Windows PowerShell 5.1.
- Live tests force the confirmation gate on and use only undeliverable addresses for outward-facing previews. `tests/Run-AllTests.ps1` checks Sent Items and the Outbox around every test and stops the run if either grows.

## [2.2.1] - 2026-10-22

### Added
- `outlook-priority.ps1 -AllFolders` ranks all received mail: the Inbox plus every folder that rules move mail to. It skips Sent Items, Drafts, Deleted Items, Junk Email, Outbox, Sync Issues and system folders, and each row shows its folder.

### Changed
- `outlook-priority.ps1` sorts newest first inside each priority level across folders.

### Notes
- Copilot rarely scores mail that rules file before it reaches the Inbox. That mail is listed last as "not scored".

## [2.2.0] - 2026-10-22

Copilot insights. Copilot stores its analysis on each email in Exchange, and the mail skill now reads it through COM. No AI is called. The fields are undocumented Microsoft properties in `PSETID_Common`, kept in one place (`$script:CopilotProp` in `_common.ps1`).

### Added
- `outlook-priority.ps1`: mail ranked by Copilot's priority (High > Normal > Low > not scored), with the same filters as triage. `-Why` adds Copilot's reason and `-Summary` prints counts only.
- `outlook-smart-reply.ps1`: Copilot's up to 3 suggested replies, for one email or as a list. `-Use N [-ReplyAll]` saves a reply draft in the thread and never sends.
- `outlook-reply.ps1 -SmartReply N`: a reply draft from suggestion N.
- `outlook-triage.ps1` shows Copilot's one-line headline on each row. `-Copilot` adds the reason and suggested replies.
- `outlook-read-body.ps1` prints a Copilot block (priority, headline, reason, suggestions) in text, `-Json` and HTML output.
- `outlook-find.ps1` and `outlook-read.ps1` take `-Priority High,Normal,Low` and add `CopilotPriority` to each row.
- `Get-CopilotInsight` in `_common.ps1`.
- `tests/Test-MailSmartReply.ps1`.

### Fixed
- Under Windows PowerShell 5.1, `-Json` wrote empty arrays as `{}` instead of `[]`.

## [2.1.0] - 2026-10-22

### Added
- `outlook-triage.ps1` with three views: `untagged` (the default: unread and untagged mail from the last 3 days), `tagged` (grouped by tag set) and `reply` (not replied, oldest first, with days waiting). Every filter works with every view.
- Copilot's inbox priority (the High / Low arrow in Outlook), read from `CopilotInboxScore`.
- Replied, replied all and forwarded state, and whether the user is in To or only CC.
- `tests/Test-MailTriage.ps1`.

### Fixed
- After Outlook closed or crashed, the warm host failed or reported "No accounts found". It now detects a dead session and reconnects.
- When Outlook is not running, scripts start it with its window and wait up to 90 s. A hidden Outlook started after a crash froze on its recovery prompt.

## [2.0.0] - 2026-10-22

The mail skill rebuilt to be fast and agent-friendly. It now has 51 scripts, 16 of them new.

### Added
- A warm host (`outlook-host.ps1` and `outlook-call.ps1`) keeps one Outlook session behind a current-user named pipe and exits after 30 idle minutes.
- `-Json` on every script.
- Links shown as numbered citations (`[URL-001]`) with a reference table. Safe Links are decoded. `outlook-links.ps1 -Ref` returns the full address.
- Full threads (`outlook-conversation.ps1`).
- HTML reply and forward inside the thread in Outlook's own format, with an updatable draft block.
- Rules: create them with every condition, exception and action, and enable, disable, rename, reorder, run or delete them.
- Folder management, bulk actions, multi-folder and whole-mailbox search, selection, internet headers (SPF, DKIM, DMARC), and conversation-wide actions.
- Junk lists, signatures, reminders, stores and sizes, copy, `.msg` import, and sending a reviewed draft.
- One confirmation switch for every irreversible action: per call (`-Confirm`), per shell (`OUTLOOK_SUITE_REQUIRE_CONFIRM`) or per user (`%APPDATA%\outlook-suite\config.json`).
- Live tests in `tests/`. They print counts and PASS/FAIL only.

### Changed
- Searches read one table per folder instead of opening every item, so they take about half the time (for example, find over 30 days went from 2.1 s to 0.86 s).
- EntryID first everywhere. Move and copy print the new EntryID. An unknown folder is an error, never a silent fallback to the Inbox.
- Every failure prints `Error:` and exits 1.
- Every script runs under Windows PowerShell 5.1 and PowerShell 7 with identical results. Old parameter names still work.

## [1.0.0] - 2026-02-06

### Added
- The first plugin marketplace release, with mail, calendar, contacts, tasks and notes skills driving classic Outlook through PowerShell COM.
- `-StripLinks` on read-body, to keep long URLs out of agent context (2026-02-18).

[3.0.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/6
[2.2.1]: https://github.com/mustafaakben/outlook-suite-plugin/pull/4
[2.2.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/3
[2.1.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/2
[2.0.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/1
[1.0.0]: https://github.com/mustafaakben/outlook-suite-plugin/commit/39ffaef
