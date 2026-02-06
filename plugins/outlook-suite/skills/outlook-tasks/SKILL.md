---
name: outlook-tasks
description: "Manage Outlook tasks via PowerShell: list, create, update, complete, assign, and delete tasks with priorities, due dates, reminders, and status tracking. Scripts auto-start Outlook when needed."
---

# Outlook Tasks

Manage Microsoft Outlook tasks using PowerShell COM objects.

## Prerequisites

- Microsoft Outlook installed (scripts auto-start Outlook if it is not running)
- PowerShell 5.1+ (included in Windows 10/11)
- Task assignment requires Exchange/Office 365

## List Tasks

### outlook-tasks-list.ps1

Browse tasks with optional status filtering. Returns EntryIDs for use with action scripts.

**Required Parameters:** None (lists active tasks by default)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Limit` | int | 20 | Maximum tasks to display |
| `-IncludeCompleted` | switch | false | Include completed tasks in the list |
| `-Status` | string | All | Filter by status: All, NotStarted, InProgress, Completed, Waiting, Deferred |

```powershell
# List active tasks
& "./scripts/outlook-tasks-list.ps1" -Limit 20

# Include completed tasks
& "./scripts/outlook-tasks-list.ps1" -IncludeCompleted

# Filter by status
& "./scripts/outlook-tasks-list.ps1" -Status InProgress
```

## Create Task

### outlook-tasks-create.ps1

Create a new task with optional due date, priority, and reminder.

**Required Parameters:** `-Subject`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | **Required** | Task subject/title |
| `-DueDate` | datetime | — | Due date (e.g., "2026-02-15") |
| `-StartDate` | datetime | — | Start date |
| `-Priority` | string | Normal | Priority: Low, Normal, High |
| `-Body` | string | — | Task description/notes |
| `-Reminder` | datetime | — | Reminder date and time (e.g., "2026-02-10 09:00") |

```powershell
# Create a task with due date
& "./scripts/outlook-tasks-create.ps1" -Subject "Review proposal" -DueDate "2026-02-15" -Priority High

# Create task with reminder
& "./scripts/outlook-tasks-create.ps1" -Subject "Call client" -Reminder "2026-02-10 09:00"

# Create task with body text
& "./scripts/outlook-tasks-create.ps1" -Subject "Prepare slides" -Body "Include Q4 data and projections"
```

## Update Task

### outlook-tasks-update.ps1

Edit an existing task's fields. Uses EntryID (preferred) or Index as targeting method.

**Required Parameters:** `-EntryID` or `-Index`, plus at least one field to update

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Task's unique EntryID (preferred, from list) |
| `-Index` | int | 0 | Task position number (fallback) |
| `-Subject` | string | — | New subject/title |
| `-DueDate` | datetime | — | New due date |
| `-StartDate` | datetime | — | New start date |
| `-Status` | string | — | New status: NotStarted, InProgress, Completed, Waiting, Deferred |
| `-Priority` | string | — | New priority: Low, Normal, High |
| `-Body` | string | — | New description/notes |
| `-PercentComplete` | int | -1 | Progress percentage (0-100, auto-updates status) |
| `-Reminder` | datetime | — | Set reminder date and time |
| `-ClearReminder` | switch | false | Remove existing reminder |
| `-IncludeCompleted` | switch | false | Allow targeting completed tasks by Index |

```powershell
# Update status and progress by EntryID
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -Status InProgress -PercentComplete 50

# Change due date and priority by index
& "./scripts/outlook-tasks-update.ps1" -Index 1 -DueDate "2026-02-20" -Priority High

# Set a reminder
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -Reminder "2026-02-10 09:00"

# Clear a reminder
& "./scripts/outlook-tasks-update.ps1" -EntryID "00000000..." -ClearReminder
```

## Complete Task

### outlook-tasks-complete.ps1

Mark a task as completed. Supports targeting by EntryID (preferred), Index, or Subject.

**Required Parameters:** One of `-EntryID`, `-Index`, or `-Subject`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Task's unique EntryID (preferred) |
| `-Index` | int | 0 | Task position number (fallback, non-completed tasks only) |
| `-Subject` | string | — | Exact subject match (fallback) |

```powershell
# Complete by EntryID (preferred)
& "./scripts/outlook-tasks-complete.ps1" -EntryID "00000000..."

# Complete by index
& "./scripts/outlook-tasks-complete.ps1" -Index 1

# Complete by subject
& "./scripts/outlook-tasks-complete.ps1" -Subject "Review proposal"
```

## Assign Task

### outlook-tasks-assign.ps1

Delegate a task to another person via Exchange/Microsoft 365. Sends a task request email.

**Required Parameters:** (`-EntryID` or `-Index`) and `-AssignTo`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Task's unique EntryID (preferred) |
| `-Index` | int | 0 | Task position number (fallback) |
| `-AssignTo` | string | **Required** | Recipient email address |
| `-KeepCopy` | switch | false | Keep a copy of the task in your list |
| `-IncludeCompleted` | switch | false | Allow targeting completed tasks by Index |

```powershell
# Assign by EntryID
& "./scripts/outlook-tasks-assign.ps1" -EntryID "00000000..." -AssignTo "colleague@example.com"

# Assign by index and keep a copy
& "./scripts/outlook-tasks-assign.ps1" -Index 1 -AssignTo "colleague@example.com" -KeepCopy
```

## Delete Task

### outlook-tasks-delete.ps1

Delete a task from Outlook. Supports targeting by EntryID (preferred), Index, or Subject.

**Required Parameters:** One of `-EntryID`, `-Index`, or `-Subject`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Task's unique EntryID (preferred) |
| `-Index` | int | 0 | Task position number (fallback, non-completed tasks only) |
| `-Subject` | string | — | Exact subject match (fallback) |

```powershell
# Delete by EntryID (preferred)
& "./scripts/outlook-tasks-delete.ps1" -EntryID "00000000..."

# Delete by index
& "./scripts/outlook-tasks-delete.ps1" -Index 1

# Delete by subject
& "./scripts/outlook-tasks-delete.ps1" -Subject "Old task"
```

## Reference: Task Status Values

| Value | Status |
|-------|--------|
| 0 | Not Started |
| 1 | In Progress |
| 2 | Completed |
| 3 | Waiting |
| 4 | Deferred |

## Reference: Priority Values

| Value | Priority |
|-------|----------|
| 0 | Low |
| 1 | Normal |
| 2 | High |
