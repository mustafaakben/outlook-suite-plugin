# Outlook Suite (Outlook-Suit)

Local Microsoft Outlook automation suite (Mail, Calendar, Contacts, Tasks, Notes), packaged for Claude Code plugin distribution.

This repository is now a plugin marketplace repo that publishes one plugin: `outlook-suite`.

## Purpose

The `outlook-suite` plugin provides modular Outlook automation skills built on PowerShell + Outlook COM:
- Mail: read, search, triage and manage email and drafts, Copilot priority and suggested replies, rules, categories and folders
- Calendar: find free time ranked by how full each day is, propose times as email-ready text, week overview, meeting prep, holds and invites, respond and cancel
- People and contacts: find anyone across your contacts, your organization's directory and recent mail, profiles, org chart, list members, duplicate merge, vCard/CSV import
- Tasks: a daily or weekly agenda, overdue triage, tasks from email, recurring tasks and due views
- Notes: create, read, search and delete Outlook sticky notes

Every irreversible or outward-facing action previews by default and acts only with `-Confirm`.

## Repository Structure

- `.claude-plugin/marketplace.json` - Marketplace manifest (install source)
- `plugins/outlook-suite/.claude-plugin/plugin.json` - Plugin manifest
- `plugins/outlook-suite/skills/outlook-suite/` - Orchestration skill
- `plugins/outlook-suite/skills/outlook-mail/` - Mail skill + scripts
- `plugins/outlook-suite/skills/outlook-calendar/` - Calendar skill + scripts
- `plugins/outlook-suite/skills/outlook-contacts/` - Contacts skill + scripts
- `plugins/outlook-suite/skills/outlook-tasks/` - Tasks skill + scripts
- `plugins/outlook-suite/skills/outlook-notes/` - Notes skill + scripts

## Install as Plugin

```bash
/plugin marketplace add <github-owner>/<github-repo>
/plugin install outlook-suite@outlook-suite-marketplace
```

## Plugin Lifecycle Commands

```bash
# Refresh marketplace index after pushing updates
/plugin marketplace update outlook-suite-marketplace

# Update installed plugin to latest marketplace version
/plugin update outlook-suite

# Remove plugin
/plugin uninstall outlook-suite

# Validate manifests locally before publishing
claude plugin validate .
claude plugin validate plugins/outlook-suite
```

## Requirements

- Windows 10/11
- Microsoft Outlook desktop installed
- PowerShell 5.1+ (PowerShell 7 `pwsh` for the mail warm host)
- Outlook profile configured with mailbox access
- Claude Code with plugin support (`/plugin ...` commands)

Optional (feature-dependent):
- Exchange Online / Microsoft 365 for features like free/busy, room lookup, out-of-office, recall constraints, and some assignment operations

## Safety Model

The suite follows strict operational safety patterns:
- Prefer `EntryID` targeting over volatile indexes
- Require explicit confirmation for irreversible or outward-facing operations
- Favor draft-first workflows for outgoing email
- Align docs with actual script behavior and safety gates

## Confirmation

Every irreversible or outward-facing mail action is available to agents: send, send-as, send a draft, schedule send, reply/forward send, delete (and permanent delete), recall, rules, folder rename/move/delete/empty, bulk actions, category deletion, out-of-office, junk lists, signatures, reminder dismiss, conversation always-rules, print. Since 3.0 the other modules use the same gate: calendar create and recurring with attendees, hold-to-invite, update with -SendUpdate, cancel, respond, forward -Send and delete; contact delete, dedupe merge and import; task delete, assign and bulk triage; note delete. Each one checks a confirmation flag. Without confirmation the script prints `PREVIEW ONLY: ... was NOT done.`, describes exactly what it would do, and exits 0. Reversible actions (drafts, move, copy, flags, categories, mark read) act immediately.

The flag is the user's choice and can be changed at three levels (first match wins):

1. Per call: add `-Confirm`.
2. Per shell: `$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "0"` turns confirmation off, `"1"` forces it on.
3. Per user: `%APPDATA%\outlook-suite\config.json`, for example `{ "requireConfirm": true, "actions": { "send": true, "bulk": false } }`. The action names are listed in `plugins/outlook-suite/skills/outlook-mail/SKILL.md`.

With nothing configured, confirmation is required. Agents are told to ask the user before passing `-Confirm`. With `-Json`, a preview is a JSON object with `Result` "preview".

All scripts in every module exit 1 with an `Error: ...` line on failure, and accept `-Json` for machine-readable output.

## Local Script Usage (Optional)

You can also run scripts directly from the packaged plugin paths:

```powershell
cd plugins/outlook-suite/skills/outlook-mail

# What needs attention: unread untagged mail (default), mail by tag, or mail not replied to,
# each with Copilot's priority; every filter is a parameter
& ".\scripts\outlook-triage.ps1"
& ".\scripts\outlook-triage.ps1" reply -Category "@Action,@Review" -Priority High -Days 10
& ".\scripts\outlook-priority.ps1" -Days 1 -Unread -Why
& ".\scripts\outlook-smart-reply.ps1" -EntryID "00000000..." -Use 1

# Find candidate emails and capture EntryID
& ".\scripts\outlook-find.ps1" -Days 2 -From "john@example.com"

# Reply by EntryID (saved as a draft in the thread; prints the draft's EntryID)
& ".\scripts\outlook-reply.ps1" -EntryID "00000000..." -Body "Thanks, received."

# Send the reviewed draft
& ".\scripts\outlook-send-draft.ps1" -EntryID "<draft EntryID>" -Confirm

# Same scripts through the warm host (faster for repeated calls)
& ".\scripts\outlook-call.ps1" find -Days 2 -UnreadOnly
```

## What's new in 3.0: examples

Calendar, people and contacts, tasks and notes now work like the mail scripts: fast, `-Json` everywhere, previews before anything irreversible, and all of them run through the warm host. The names below are placeholders.

```powershell
$call = "plugins/outlook-suite/skills/outlook-mail/scripts/outlook-call.ps1"

# Today's brief: meetings vs holds vs firm events, tasks and flagged email due, priority unread mail
pwsh -File $call agenda

# Three meeting times with Jane, best days first, as email-ready lines
pwsh -File $call calendar-free -Attendees "jane.doe@example.org" -Duration 30 -Propose

# The week at a glance: meetings, holds, firm events, conflicts and each day's tier
pwsh -File $call calendar-week

# Prep for the next meeting: agenda and join link, who is coming (title, department,
# mail exchanged with each) and related mail by subject; read-only
pwsh -File $call calendar-prep -Next

# Cancel a meeting you organize and notify everyone: previews who would be told and
# changes nothing until you add -Confirm
pwsh -File $call calendar-cancel -EntryID "00000000..." -Message "I need to move this, new time soon."

# Find someone and their contact points, then a "who is this" profile
pwsh -File $call people "jane doe"
pwsh -File $call people "jane doe" -Profile

# Overdue tasks by age and category; preview a bulk clean-up (add -Confirm to apply it)
pwsh -File $call tasks-triage
pwsh -File $call tasks-triage -Older 30 -ClearDue

# Turn an email into a task due tomorrow (the email itself is not changed)
pwsh -File $call tasks-from-mail -EntryID "00000000..." -When Tomorrow

# Search the full text of your sticky notes
pwsh -File $call notes-list -Search "budget" -Json
```

## Module Documentation

- `plugins/outlook-suite/skills/outlook-suite/SKILL.md`
- `plugins/outlook-suite/skills/outlook-mail/SKILL.md`
- `plugins/outlook-suite/skills/outlook-calendar/SKILL.md`
- `plugins/outlook-suite/skills/outlook-contacts/SKILL.md`
- `plugins/outlook-suite/skills/outlook-tasks/SKILL.md`
- `plugins/outlook-suite/skills/outlook-notes/SKILL.md`
- [`CHANGELOG.md`](CHANGELOG.md): what changed in each version

## Limitations and Known Constraints

- Windows-only (Outlook COM automation)
- Requires Outlook desktop app and configured local profile
- Not intended for Outlook Web App (OWA)-only environments
- Behavior can vary by account type, mailbox policy, and Exchange tenant config
- Message recall is limited by Outlook/Exchange behavior and not guaranteed
- `EntryID` is safest for targeting, but can change across certain mailbox move/import scenarios

## Review Status

A full module-by-module review was completed before packaging:
- Parser/syntax validation across all `.ps1` files
- Safety hardening for destructive operations
- Parameter validation and conflict handling improvements
- Documentation corrections and behavior alignment

## License

No license file is currently included. Add a `LICENSE` before public release if needed.

## Author

Dr. Mustafa Akben
