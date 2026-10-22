# Changelog

All notable changes to the Outlook Suite plugin. Versions follow [Semantic Versioning](https://semver.org/). The format is based on [Keep a Changelog](https://keepachangelog.com/).

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

[2.2.1]: https://github.com/mustafaakben/outlook-suite-plugin/pull/4
[2.2.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/3
[2.1.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/2
[2.0.0]: https://github.com/mustafaakben/outlook-suite-plugin/pull/1
[1.0.0]: https://github.com/mustafaakben/outlook-suite-plugin/commit/39ffaef
