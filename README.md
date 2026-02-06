# Outlook Suite (Outlook-Suit)

Local Microsoft Outlook automation suite (Mail, Calendar, Contacts, Tasks, Notes), packaged for Claude Code plugin distribution.

This repository is now a plugin marketplace repo that publishes one plugin: `outlook-suite`.

## Purpose

The `outlook-suite` plugin provides modular Outlook automation skills built on PowerShell + Outlook COM:
- Read/search/manage email and drafts
- Create/update/delete calendar events and respond to meetings
- Create/search/update contacts and export vCards
- Create/update/complete/assign tasks
- Create/read/delete Outlook sticky notes

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
- PowerShell 5.1+
- Outlook profile configured with mailbox access
- Claude Code with plugin support (`/plugin ...` commands)

Optional (feature-dependent):
- Exchange Online / Microsoft 365 for features like free/busy, room lookup, out-of-office, recall constraints, and some assignment operations

## Safety Model

The suite follows strict operational safety patterns:
- Prefer stable `EntryID` targeting over volatile indexes
- Require explicit confirmation flags for send/delete/print-sensitive operations
- Favor draft-first workflows for outgoing email
- Align docs with actual script behavior and safety gates

## Local Script Usage (Optional)

You can also run scripts directly from the packaged plugin paths:

```powershell
cd plugins/outlook-suite/skills/outlook-mail

# Find candidate emails and capture EntryID
& ".\scripts\outlook-find.ps1" -Days 2 -From "john@example.com"

# Reply by EntryID
& ".\scripts\outlook-reply.ps1" -EntryID "00000000..." -Body "Thanks, received."
```

## Module Documentation

- `plugins/outlook-suite/skills/outlook-suite/SKILL.md`
- `plugins/outlook-suite/skills/outlook-mail/SKILL.md`
- `plugins/outlook-suite/skills/outlook-calendar/SKILL.md`
- `plugins/outlook-suite/skills/outlook-contacts/SKILL.md`
- `plugins/outlook-suite/skills/outlook-tasks/SKILL.md`
- `plugins/outlook-suite/skills/outlook-notes/SKILL.md`

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
