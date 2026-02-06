---
name: outlook-suite
description: "Complete Microsoft Outlook automation via PowerShell COM objects on Windows. Covers email (read, search, draft, send, reply, forward, move, delete, flag, categorize, rules, attachments, out-of-office), calendar (list events, create/update/delete appointments, respond to meetings, recurring events, free/busy, rooms), contacts (list, create, search, update, delete, export vCard, photos), tasks (list, create, update, complete, assign, delete with priorities and reminders), and notes (list, create, read, delete sticky notes). Use when the user mentions Outlook, email, calendar, meetings, contacts, tasks, notes, or any mail/scheduling operation on Windows."
---

# Outlook Suite

Automates Microsoft Outlook operations across email, calendar, contacts, tasks, and notes using PowerShell COM objects. 62 scripts organized into 5 focused modules.

## Prerequisites

- Microsoft Outlook installed on Windows (scripts can start Outlook if needed)
- PowerShell 5.1+ (included in Windows 10/11)
- Exchange Server or Microsoft 365 for some features (shared mailboxes, free/busy, rooms, task assignment)

## Safety Rules

1. **Never auto-send** — all send operations require `-Confirm` parameter
2. **Never delete without confirmation** — always confirm with the user before running delete scripts
3. **Draft-first for compose** — create as draft, show the user, then send only after approval
4. **Verify targets** — when acting on a specific item, confirm the subject/recipient with the user before modifying
5. **Prefer EntryID** — always use EntryID over Index when targeting items (Index numbers shift as new items arrive)

## Item Targeting Pattern

All modules use **EntryID** as the primary method to target specific items. EntryID is a permanent, unique identifier that never changes.

**Workflow:**
1. Run a listing or search script to discover items (many return EntryID directly; for mail, prefer `outlook-find.ps1` for stable EntryIDs)
2. Use the EntryID from the results to target a specific item in action scripts

```powershell
# Step 1: Find the item
powershell -File outlook-find.ps1 -From "john@example.com" -Days 3

# Step 2: Act on it using EntryID from results
powershell -File outlook-read-body.ps1 -EntryID "00000000..."
```

Action scripts also accept `-Index` as a fallback, but EntryID is strongly preferred for accuracy.

## Modules

This suite is organized into 5 self-contained modules. Each module has its own detailed documentation with full parameter tables, examples, and reference data. **Load the relevant module documentation based on what the user needs.**

### Email — [outlook-mail/SKILL.md](../outlook-mail/SKILL.md)

The largest module (35 scripts). Handles all email operations:

- **Find and target** — `outlook-find.ps1` for EntryID discovery with rich filtering
- **Read and search** — read inbox, read full body, search, advanced search
- **Compose and send** — draft, update draft, send, send-as, reply, forward, voting buttons
- **Organize** — move, delete, mark read/unread, conversation view, print
- **Flags and metadata** — flag, unflag, importance, sensitivity, save-as
- **Categories** — list, create, assign, remove, delete categories
- **Scheduling** — schedule send, request receipts, recall messages
- **Account management** — folders, accounts, rules, out-of-office, attachment save

### Calendar — [outlook-calendar/SKILL.md](../outlook-calendar/SKILL.md)

10 scripts for calendar and meeting management:

- **Events** — list, create, update, delete appointments
- **Meetings** — respond to invites, manage attendees, forward meetings
- **Recurring events** — create with daily/weekly/monthly/yearly patterns
- **Availability** — check free/busy status, find available rooms

### Contacts — [outlook-contacts/SKILL.md](../outlook-contacts/SKILL.md)

7 scripts for contact management:

- **Browse and search** — list contacts, search by name/email/company
- **Manage** — create, update, delete contacts
- **Extras** — add/remove contact photos, export as vCard

### Tasks — [outlook-tasks/SKILL.md](../outlook-tasks/SKILL.md)

6 scripts for task management:

- **Browse** — list tasks with status/priority filtering
- **Manage** — create, update, complete, delete tasks
- **Delegate** — assign tasks to others (requires Exchange)

### Notes — [outlook-notes/SKILL.md](../outlook-notes/SKILL.md)

4 scripts for Outlook sticky notes:

- **Browse** — list notes with color and date filtering
- **Manage** — create, read full content, delete notes

## Script Quick Reference

> **IMPORTANT FOR AI AGENTS:** This index gives you the correct script name for each operation. Before running any script, you **MUST** read the relevant module SKILL.md (linked in each heading) for full parameter documentation, required vs optional parameters, and usage examples. Do NOT guess parameters from the script name alone.

### Email — [outlook-mail/SKILL.md](../outlook-mail/SKILL.md)

**Find and Target**
- `outlook-find.ps1` — Find emails by criteria, return stable EntryIDs for action scripts

**Read and Search**
- `outlook-read.ps1` — List recent inbox emails with sender, date, unread status, and EntryID
- `outlook-read-body.ps1` — Read full email body by EntryID or index
- `outlook-search.ps1` — Search emails by subject, sender, or body text
- `outlook-search-advanced.ps1` — Multi-filter advanced search (attachments, importance, flags, dates, folders)

**Compose and Send**
- `outlook-draft.ps1` — Create draft email (plain text or HTML, with attachments)
- `outlook-draft-update.ps1` — Modify existing draft (recipients, subject, body, attachments)
- `outlook-send.ps1` — Send email directly (requires -Confirm)
- `outlook-send-as.ps1` — Send from a specific account (requires -Confirm)
- `outlook-reply.ps1` — Reply or Reply All to email (requires -Confirm to send)
- `outlook-forward.ps1` — Forward email to recipients (requires -Confirm to send)
- `outlook-voting.ps1` — Send email with voting buttons/poll (requires -Confirm)

**Organize**
- `outlook-move.ps1` — Move email to a folder by name or path
- `outlook-delete.ps1` — Delete email, soft or permanent (requires -Confirm)
- `outlook-mark-read.ps1` — Mark email as read or unread
- `outlook-conversation.ps1` — View full email conversation thread
- `outlook-print.ps1` — Print email to default printer

**Flags, Importance and Sensitivity**
- `outlook-flag.ps1` — Flag email for follow-up with optional due date and reminder
- `outlook-unflag.ps1` — Remove flag or mark flag as complete
- `outlook-set-importance.ps1` — Set email priority (Low, Normal, High)
- `outlook-set-sensitivity.ps1` — Set email sensitivity (Normal, Personal, Private, Confidential)
- `outlook-save-as.ps1` — Export email to file (msg, txt, html, mht)

**Categories**
- `outlook-list-categories.ps1` — List all categories with colors
- `outlook-create-category.ps1` — Create new category with color
- `outlook-assign-category.ps1` — Assign category to email (appends, does not replace)
- `outlook-remove-category.ps1` — Remove specific or all categories from email
- `outlook-delete-category.ps1` — Delete category from master list

**Scheduling and Receipts**
- `outlook-schedule-send.ps1` — Schedule email to send at a future date/time
- `outlook-request-receipt.ps1` — Send with read and/or delivery receipt requests
- `outlook-recall.ps1` — Attempt to recall a sent email (Exchange only)

**Account and Rule Management**
- `outlook-folders.ps1` — List all mail folders with item and unread counts
- `outlook-accounts-list.ps1` — List all configured email accounts
- `outlook-rules-list.ps1` — List email rules with ON/OFF status
- `outlook-rules-create.ps1` — Create new email rule with conditions and actions
- `outlook-out-of-office.ps1` — Configure Out of Office auto-replies (Exchange)
- `outlook-attachment-save.ps1` — Save email attachments to disk (skips inline images by default)

### Calendar — [outlook-calendar/SKILL.md](../outlook-calendar/SKILL.md)

**Events and Appointments**
- `outlook-calendar-list.ps1` — List upcoming calendar events with EntryIDs
- `outlook-calendar-create.ps1` — Create appointment or meeting with attendees
- `outlook-calendar-update.ps1` — Update an existing event's details (subject, time, location)
- `outlook-calendar-delete.ps1` — Delete a calendar event or meeting

**Meetings and Responses**
- `outlook-calendar-respond.ps1` — Accept, tentatively accept, or decline a meeting invitation
- `outlook-calendar-attendees.ps1` — View attendee list with response status
- `outlook-calendar-forward.ps1` — Forward event as iCalendar (.ics) attachment

**Recurring Events**
- `outlook-calendar-recurring.ps1` — Create recurring event (daily, weekly, monthly, yearly patterns)

**Availability and Rooms**
- `outlook-freebusy.ps1` — Check free/busy availability for attendees (Exchange)
- `outlook-calendar-rooms.ps1` — Find available conference rooms for a time slot (Exchange)

> **AI AGENTS:** Read [outlook-calendar/SKILL.md](../outlook-calendar/SKILL.md) for full parameter tables and examples before running any calendar script.

### Contacts — [outlook-contacts/SKILL.md](../outlook-contacts/SKILL.md)

**Browse and Search**
- `outlook-contacts-list.ps1` — Browse contacts with optional search filtering, returns EntryIDs
- `outlook-contacts-search.ps1` — Search contacts by name, email, and/or company (AND logic)

**Manage**
- `outlook-contacts-create.ps1` — Create new contact (name, email, phone, company, job title)
- `outlook-contacts-update.ps1` — Edit existing contact fields by EntryID or index
- `outlook-contacts-delete.ps1` — Delete contact by EntryID, index, or exact name

**Extras**
- `outlook-contacts-photo.ps1` — Add or remove contact photo (JPG, PNG, GIF, BMP)
- `outlook-contacts-export.ps1` — Export contact as vCard (.vcf) file

### Tasks — [outlook-tasks/SKILL.md](../outlook-tasks/SKILL.md)

**Browse**
- `outlook-tasks-list.ps1` — List tasks with optional status filtering, returns EntryIDs

**Manage**
- `outlook-tasks-create.ps1` — Create new task with due date, priority, and reminder
- `outlook-tasks-update.ps1` — Edit task fields (status, priority, due date, progress, reminder)
- `outlook-tasks-complete.ps1` — Mark task as completed by EntryID, index, or subject
- `outlook-tasks-delete.ps1` — Delete task by EntryID, index, or subject

**Delegate**
- `outlook-tasks-assign.ps1` — Assign task to another person via Exchange

### Notes — [outlook-notes/SKILL.md](../outlook-notes/SKILL.md)

**Browse**
- `outlook-notes-list.ps1` — List notes with optional color and search filtering, returns EntryIDs

**Manage**
- `outlook-notes-create.ps1` — Create sticky note with optional color (Blue, Green, Pink, Yellow, White)
- `outlook-notes-read.ps1` — Read full note content by EntryID or index
- `outlook-notes-delete.ps1` — Delete note by EntryID or index

> **AI AGENTS:** Read the relevant module SKILL.md linked in each section heading above for full parameter documentation and examples. Do NOT guess parameter names or assume defaults — always check the module docs first.

## How to Use

1. **Identify the domain** — determine which module handles the user's request
2. **Load the module** — read the relevant module SKILL.md linked above for full script documentation
3. **Find the target** — for actions on existing items, first run a listing/search script to get EntryID
4. **Execute** — run the appropriate script with parameters documented in the module SKILL.md
5. **Confirm destructive actions** — always verify with the user before send, delete, or modify operations

## Script Execution

All scripts are PowerShell files located in each module's `scripts/` directory. Execute them with:

```powershell
powershell -File <module>/scripts/<script-name>.ps1 -Param1 value1 -Param2 value2
```

Scripts handle COM object lifecycle automatically (connection, cleanup, garbage collection).
