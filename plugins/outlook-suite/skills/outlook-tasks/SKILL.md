---
name: outlook-tasks
description: "Manage Outlook tasks via PowerShell, triage overdue tasks in bulk (group by age and category; reschedule, complete or clear due dates with a preview first), and get a daily agenda: a read-only brief of meetings vs holds, tasks due and overdue, flagged email and priority unread mail for a day or week; list tasks with due views (today, overdue, this week), include flagged email as to-dos, create (including recurring tasks and reminders), update, complete, assign and delete tasks, and turn an email into a linked task. Fast table queries, -Json output, confirmation for delete, assign and bulk triage. Scripts auto-start Outlook when needed."
---

# Outlook Tasks

Manage Microsoft Outlook tasks (and flagged email as to-dos) with PowerShell COM scripts. All scripts share the outlook-suite helpers in `../outlook-mail/scripts/_common.ps1` (loaded through `scripts/_common.ps1`), so they behave like the mail scripts: same connection handling, `-Json`, errors as `Error: ...` with exit code 1, and the same confirmation gate.

## Prerequisites

- Classic Microsoft Outlook for Windows (scripts start it with its window if it is not running).
- Windows PowerShell 5.1 or PowerShell 7. Task assignment needs Exchange or Microsoft 365.

## Fast path (warm host)

Run any script through the mail skill's warm host to skip Outlook start-up cost on every call (PowerShell 7):

```powershell
& "../outlook-mail/scripts/outlook-call.ps1" tasks-list -Due Today
& "../outlook-mail/scripts/outlook-call.ps1" tasks-complete -EntryID "00000000..."
```

Arguments and output are exactly the same as running the script directly.

## Choosing a task

Action scripts take `-EntryID` (preferred, from `outlook-tasks-list.ps1`). `-Index N` is the N-th task in the default `outlook-tasks-list.ps1` order (open tasks, by due date) and shifts as tasks change. `-Subject` (complete and delete only) must match exactly one task.

## Confirmation

`outlook-tasks-delete.ps1`, `outlook-tasks-assign.ps1` and the actions of `outlook-tasks-triage.ps1` only preview what they would do unless you pass `-Confirm` (exit code 0, "PREVIEW ONLY"). Assigning sends a task request email, so it follows the `send` setting; delete follows `delete`; triage follows `bulk` and also needs the `-Plan` code from its preview. See "Confirmation" in the outlook-mail SKILL.md for turning this off per shell or per user. Create, update and complete change only your own items and run directly.

## outlook-agenda.ps1

A read-only "what's on my plate" brief for one day, or 7 days with `-Week`. Nothing is changed. Each section has a count, and lists show at most `-Limit` items (counts are always complete).

- **Calendar**, per day, classified by the calendar skill (the same rules as its list, week and free-time scripts): meetings (someone else is invited), firm events (classes, office hours, exams, travel, Out of Office, private events), holds (your own movable blocks) and all-day events, with the day's tier (Open, Holds-only, One-meeting, Busy), conflicts, the 2-meeting cap and away days. Cancelled and declined events are listed apart and not counted.
- **Tasks**: due in the window, and overdue.
- **Flagged email**: follow-ups due in the window, and overdue.
- **Priority unread mail**: unread Inbox mail from the last `-MailDays` days (default 3) that Copilot rates High or that is marked High importance, with Copilot's headline.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Day` | string | today | `today`, `tomorrow`, a weekday name (the next one, today included) or a date |
| `-Week` | switch | false | 7 days starting at `-Day` |
| `-MailDays` | int | 3 | How far back to look for priority unread mail |
| `-Limit` | int | 10 | Items listed per section |
| `-Json` | switch | false | `{ from, to, counts, days[{ Date, Meetings, Holds, AllDay, Canceled, Flags }], tasks{due, overdue}, flaggedMail{due, overdue}, priorityMail[] }` |

```powershell
& "./scripts/outlook-agenda.ps1"                    # today
& "./scripts/outlook-agenda.ps1" -Day tomorrow
& "./scripts/outlook-agenda.ps1" -Day monday -Week -Json
```

## outlook-tasks-triage.ps1

Work through overdue tasks in bulk. Without an action it reports the overdue open tasks grouped by age (under 1 week, 1 to 4 weeks, over a month) and by category, with counts and a few examples per group. Waiting and Deferred tasks are left out unless `-IncludeParked`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Older` | int | Only tasks overdue by more than N days |
| `-Category` | string | Only tasks with this category (one name among several) |
| `-Search` | string | Only tasks whose subject contains this text |
| `-IncludeParked` | switch | Include Waiting and Deferred tasks |
| `-Reschedule` | datetime | Action: new due date (not in the past; a later start date is moved back) |
| `-Complete` | switch | Action: mark complete |
| `-ClearDue` | switch | Action: remove the due and start dates |
| `-All` | switch | Allow an action on every overdue task (otherwise an action needs a filter) |
| `-Plan` | string | The plan code from the preview; required to apply |
| `-MaxTasks` | int | Most tasks one run may change (default 100); raise it explicitly for more |
| `-Confirm` | switch | Apply the action (with `-Plan`) |
| `-Limit` | int | Examples shown per group (default 5) |
| `-Json` | switch | Report: `{ overdue, matched, parkedExcluded, byAge, byCategory, examples }`; action: `{ action, result, wouldChange or changed, failed, skippedRecurring, skippedAssigned, plan, undo[], failures[] }` |

An action takes two steps, so what is approved is exactly what runs:

1. Run it without `-Confirm`: a preview with counts, examples and a plan code (a hash of the action, its filters and the exact tasks).
2. Run it again with `-Confirm -Plan <code>`. If the matching tasks changed in between, the code no longer matches and nothing is done. `-Confirm` without `-Plan`, or more tasks than `-MaxTasks`, is refused too.

Never touched, and counted in the preview: recurring tasks (changing or completing one moves the series or creates the next occurrence) and assigned tasks (a Delegator, or completion reports going to someone, since completing them can email that person). The result lists every changed task with its EntryID, old due and start dates and old status (the undo record; restore with `outlook-tasks-update.ps1`), and every failure with its EntryID and error. The exit code is 1 when anything failed, in text and in JSON. Bulk actions also follow the `bulk` confirmation setting.

```powershell
& "./scripts/outlook-tasks-triage.ps1"                                        # the overview
& "./scripts/outlook-tasks-triage.ps1" -Older 30 -ClearDue                    # preview: counts and the plan code
& "./scripts/outlook-tasks-triage.ps1" -Older 30 -ClearDue -Confirm -Plan 3f9a1c2b7d10
& "./scripts/outlook-tasks-triage.ps1" -Category "Teaching" -Reschedule 2026-11-06 -Json
```
## outlook-tasks-list.ps1

Open tasks sorted by due date (no due date last), from one fast table query. Each task shows its EntryID, subject, status and progress, priority, due date (with OVERDUE / TODAY / Tomorrow), how it repeats ("Repeats: Weekly on Fri until 2026-12-18"), its reminder, categories, and "From email" for tasks made from an email. JSON items carry `Recurring`, `Recurrence` (when recurring) and `Reminder`.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Due` | string | Any | `Today`, `Overdue` (before today), `Week` (today and the next 6 days), `NoDate`, or `Any` |
| `-Flagged` | switch | false | Add flagged email (Type `Mail`) from the To-Do List to the tasks |
| `-Status` | string | All | All, NotStarted, InProgress, Completed, Waiting, Deferred |
| `-Search` | string | — | Subject contains this text (plain substring, any case) |
| `-Category` | string | — | Has this category |
| `-IncludeCompleted` | switch | false | Include completed tasks (implied by `-Status Completed`) |
| `-Limit` | int | 20 | Maximum tasks shown (the total matching is always reported) |
| `-Json` | switch | false | `{ folder, total, count, query, items[] }`; task dates as `yyyy-MM-dd` |

```powershell
& "./scripts/outlook-tasks-list.ps1" -Due Today
& "./scripts/outlook-tasks-list.ps1" -Due Overdue -Limit 50
& "./scripts/outlook-tasks-list.ps1" -Flagged -Due Week        # everything due this week, tasks and flagged email
& "./scripts/outlook-tasks-list.ps1" -Status InProgress -Json
```

## outlook-tasks-create.ps1

Create a task (a local item; nothing is sent).

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | **Required** | Task title |
| `-DueDate` / `-StartDate` | datetime | — | e.g. `"2026-02-15"` |
| `-Priority` | string | Normal | Low, Normal, High |
| `-Body` | string | — | Description |
| `-Reminder` | datetime | — | e.g. `"2026-02-10 09:00"` |
| `-Categories` | string | — | e.g. `"Teaching, Urgent"` |
| `-Recurrence` | string | None | Daily, Weekly, Monthly, Yearly. Repeats from the due date (today when no `-DueDate`): weekly on that weekday, monthly on that day of the month |
| `-Interval` | int | 1 | Every N days / weeks / months / years |
| `-Until` | datetime | — | Last date of the series (no `-Until` = no end) |
| `-Json` | switch | false | The new task as JSON (with its EntryID) |

```powershell
& "./scripts/outlook-tasks-create.ps1" -Subject "Review proposal" -DueDate "2026-02-15" -Priority High
& "./scripts/outlook-tasks-create.ps1" -Subject "Call client" -Reminder "2026-02-10 09:00" -Categories "Clients"
& "./scripts/outlook-tasks-create.ps1" -Subject "Weekly report" -Recurrence Weekly -DueDate "2026-11-06" -Until "2026-12-18" -Reminder "2026-11-06 09:00"
& "./scripts/outlook-tasks-create.ps1" -Subject "Pay rent" -Recurrence Monthly -DueDate "2026-11-01"
```

## outlook-tasks-from-mail.ps1

Turn an email into a task that links back to it. The email is only read, never changed (to flag the email itself, use the mail skill's `outlook-flag.ps1`). The task body names the sender and date, holds an `outlook:` link that opens the email in classic Outlook and a short preview, and the email's EntryID is stored on the task (shown by `outlook-tasks-list.ps1` as "From email"). A second task for the same email is refused unless `-Force`.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | **Required** | The email's EntryID (from outlook-find.ps1 and friends) |
| `-Subject` | string | email subject | Task title |
| `-When` | string | NoDate | Today, Tomorrow, ThisWeek (Friday), NextWeek (next Friday), NoDate |
| `-DueDate` | datetime | — | Exact due date (overrides `-When`) |
| `-Priority` | string | email importance | Low, Normal, High |
| `-Categories` / `-Reminder` | | — | As in create |
| `-AttachEmail` | switch | false | Also embed a copy of the email in the task |
| `-Force` | switch | false | Create even if a task for this email exists |
| `-Json` | switch | false | The new task as JSON |

```powershell
& "./scripts/outlook-tasks-from-mail.ps1" -EntryID "00000000..." -When Tomorrow
& "./scripts/outlook-tasks-from-mail.ps1" -EntryID "00000000..." -Subject "Reply to Dan about the budget" -DueDate "2026-11-03" -Priority High
```

## outlook-tasks-update.ps1

Change any of a task's fields; only the fields you pass change. Target with `-EntryID` or `-Index`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Subject`, `-Body` | string | New title / description |
| `-DueDate`, `-StartDate` | datetime | New dates; `-ClearDueDate` removes the due date |
| `-Status` | string | NotStarted, InProgress, Completed, Waiting, Deferred |
| `-Priority` | string | Low, Normal, High |
| `-PercentComplete` | int | 0 to 100; sets InProgress / Completed automatically when `-Status` is not given, and must agree with `-Status` when it is |
| `-Reminder` / `-ClearReminder` | datetime / switch | Set or remove the reminder |
| `-Categories` | string | Replace categories (`""` clears) |
| `-IncludeCompleted` | switch | Let `-Index` count completed tasks too |
| `-Json` | switch | The task as JSON with a `Changes` object |

```powershell
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -Status InProgress -PercentComplete 50
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -DueDate "2026-02-20" -Priority High
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -ClearDueDate -ClearReminder
```

## outlook-tasks-complete.ps1

Mark a task complete (`-EntryID`, `-Index` or exact `-Subject`). Given the EntryID of a flagged email from `outlook-tasks-list.ps1 -Flagged`, it marks that email's flag complete instead. Completing an already completed task is reported, not an error. `-Json` available.

```powershell
& "./scripts/outlook-tasks-complete.ps1" -EntryID "00000000..."
& "./scripts/outlook-tasks-complete.ps1" -Subject "Review proposal"
```

## outlook-tasks-assign.ps1

Send a task request to another person (sends an email; needs `-Confirm`). The person is resolved first, so a typo fails before anything happens. `-KeepCopy` keeps your copy of the task; `-Json` available.

```powershell
& "./scripts/outlook-tasks-assign.ps1" -EntryID "00000000..." -AssignTo "colleague@example.com"            # preview
& "./scripts/outlook-tasks-assign.ps1" -EntryID "00000000..." -AssignTo "colleague@example.com" -Confirm   # sends
```

## outlook-tasks-delete.ps1

Move a task to Deleted Items (`-EntryID`, `-Index` or exact `-Subject`; needs `-Confirm`). `-Json` available.

```powershell
& "./scripts/outlook-tasks-delete.ps1" -EntryID "00000000..."            # preview
& "./scripts/outlook-tasks-delete.ps1" -EntryID "00000000..." -Confirm
```

## Reference

| Status (`-Status`) | Value |
|--------|-------|
| NotStarted | 0 |
| InProgress | 1 |
| Completed | 2 |
| Waiting | 3 |
| Deferred | 4 |

Priority: Low 0, Normal 1, High 2. Tests: `tests/Test-Tasks.ps1`, `tests/Test-TasksTriage.ps1` and `tests/Test-Agenda.ps1` (live, on `[suite-test]` items; print PASS/FAIL, counts and timings only).
