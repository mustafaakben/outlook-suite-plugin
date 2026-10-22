---
name: outlook-mail
description: "Manage classic Outlook email on Windows through PowerShell COM: find, read, full threads, reply/forward in thread with HTML, draft, send, move, bulk actions, folders, rules, categories, flags, attachments, headers, junk, signatures, reminders, out-of-office. Use for any Outlook mail task on the local Windows PC. Fast (one Table query per search), agent-friendly (EntryIDs everywhere, -Json, exit codes), with a confirm switch on every irreversible action."
---

# Outlook Mail

51 PowerShell scripts in `scripts/` that drive classic Outlook (running on this Windows PC) through COM. Every script works standalone (`pwsh -File scripts/<name>.ps1 ...`); a warm host makes repeated calls about twice as fast.

## How to work

1. **Find first, then act by EntryID.** `outlook-find.ps1` (or `outlook-read`, `outlook-search`, `outlook-search-advanced`, `outlook-selection`) prints an `EntryID` for every email. Pass it to any action script with `-EntryID`. `-Index` still exists but positions shift as mail arrives; do not use it for actions.
2. **EntryIDs can change when an email moves.** Scripts that move or copy print the new EntryID (and the old one). Use the new one afterwards.
3. **Folders are names or paths:** `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Outbox`, `Inbox/Projects/2026`, or another mailbox: `someone@x.edu/Inbox`. `outlook-folders.ps1` lists the paths. An unknown folder is an error (exit 1), never a silent fallback to Inbox.
4. **Read the output contract:** success exits 0; any failure prints `Error: <reason>` (often a `Hint:` line) and exits 1. A gated action without confirmation prints `PREVIEW ONLY: ... was NOT done.` and exits 0.
5. **Ask before irreversible actions** (see Confirmation). When the user has already said to go ahead, pass `-Confirm`.

## Which script when

| Need | Script | Key parameters |
|---|---|---|
| List / filter emails | `outlook-find.ps1` | `-Folder -Days -DateFrom -DateTo -From -To -Subject -BodyContains [-FastBody] -UnreadOnly -ReadOnly -HasAttachment -Flagged -Importance -Category -Limit -Detailed -Json` |
| Latest emails | `outlook-read.ps1` | `-Days 1 -Limit 20 -UnreadOnly -Folder` |
| Quick text search | `outlook-search.ps1` | `-Subject -From -Body [-FastBody] -Days -Folder` |
| Search many folders / whole mailbox / all mailboxes | `outlook-search-advanced.ps1` | `-Folders "Inbox;Archive" -IncludeSubfolders -WholeMailbox -AllMailboxes` + find filters, `-SaveAs "name"` (Outlook Search Folder) |
| What the user has selected in Outlook | `outlook-selection.ps1` | `-Body` |
| One email in full | `outlook-read-body.ps1` | `-EntryID [-Links domain|full|none] [-Format html] -Json` |
| Whole thread, oldest first | `outlook-conversation.ps1` | `-EntryID [-TrimQuoted] [-Preview 300] [-HeadersOnly]` |
| Full URLs behind `[URL-00N]` | `outlook-links.ps1` | `-EntryID [-Ref URL-003]` |
| Internet headers, SPF/DKIM/DMARC, phishing signs | `outlook-headers.ps1` | `-EntryID [-Raw]` |
| Attachments: list / save | `outlook-attachment-save.ps1` | `-EntryID -ListOnly | -Path -AttachmentIndex -IncludeInline` |
| New draft | `outlook-draft.ps1` | `-To -Subject -Body [-Native|-HTML|-HtmlFile] -CC -BCC -Attachment -Account -Importance -Sensitivity` |
| Change a draft | `outlook-draft-update.ps1` | `-EntryID -To/-AddTo -CC/-AddCC -Subject -Body [-AppendBody] -AddAttachment -RemoveAttachment -Importance -Account` |
| Send a reviewed draft | `outlook-send-draft.ps1` | `-EntryID -Confirm` |
| Send new email | `outlook-send.ps1` / `outlook-send-as.ps1` (`-Account`) | same as draft + `-Confirm` |
| Send later | `outlook-schedule-send.ps1` | `-SendAt "2026-11-02 09:00" -Confirm` |
| Voting buttons / receipts | `outlook-voting.ps1` (`-Options "A;B"`), `outlook-request-receipt.ps1` | `-Confirm` |
| Reply / reply all in thread | `outlook-reply.ps1` | `-EntryID -Body | -HTML/-HtmlFile [-ReplyAll] [-RemoveAutoSignature] [-DraftEntryID] [-Send -Confirm]` |
| Forward | `outlook-forward.ps1` | `-EntryID -To [-Body] [-DraftEntryID] [-Send -Confirm]` |
| Move / copy one email | `outlook-move.ps1`, `outlook-copy.ps1` | `-EntryID -Folder` |
| Delete | `outlook-delete.ps1` | `-EntryID [-Permanent] -Confirm` |
| Many emails at once | `outlook-bulk.ps1` | `-Action Move|Copy|MarkRead|MarkUnread|Flag|Unflag|AddCategory|RemoveCategory|ClearCategories|Importance|Delete -Value` and `-EntryID a,b,c` or `-Folder` + find filters, `-Confirm` |
| Read / unread, flag, importance, sensitivity | `outlook-mark-read.ps1` (`-Unread`), `outlook-flag.ps1` (`-When Tomorrow -DueDate -ReminderDate`), `outlook-unflag.ps1` (`-Complete`), `outlook-set-importance.ps1`, `outlook-set-sensitivity.ps1` | `-EntryID` |
| Categories on an email | `outlook-assign-category.ps1`, `outlook-remove-category.ps1` | `-EntryID -Category "A","B"` / `-All` |
| Master category list | `outlook-list-categories.ps1`, `outlook-create-category.ps1` (`-Color Green`), `outlook-update-category.ps1` (`-NewName -Color -Shortcut F2`), `outlook-delete-category.ps1` | |
| Folders: list | `outlook-folders.ps1` | `-Depth -Under "Inbox" -AllStores -NoCounts -Json` |
| Folders: create, rename, move, copy, delete, empty, mark all read | `outlook-folder-manage.ps1` | `-Action -Folder -Name -To -Confirm` |
| Thread-wide actions | `outlook-conversation-actions.ps1` | `-EntryID -Action Status|MarkRead|MarkUnread|AlwaysMove -Folder|AlwaysDelete|AlwaysCategorize -Category|Stop...` |
| Rules: list | `outlook-rules-list.ps1` | `-Detailed -Mailbox -Name -Enabled` |
| Rules: create | `outlook-rules-create.ps1` | conditions (`-FromAddress -SubjectContains -BodyContains -SenderAddressContains -HeaderContains -SentTo -HasAttachment -Importance -Category -OnlyToMe ...`), exceptions (`-Except...`), actions (`-MoveToFolder -CopyToFolder -AssignCategory -ForwardTo -RedirectTo -MarkAsTask -Delete -StopProcessing ...`), `-SendRule`, `-RunNow`, `-Confirm` |
| Rules: enable, disable, rename, reorder, delete, run now | `outlook-rules-manage.ps1` | `-Name|-Order -Enable|-Disable|-NewName|-MoveTo|-Delete|-Run [-Folder -Messages] -Confirm` |
| Junk | `outlook-junk.ps1` | `-Action MarkJunk|NotJunk|BlockSender|SafeSender|Lists|OpenOptions -EntryID` |
| Signatures | `outlook-signatures.ps1` | `-Action List|Show|Create|Delete|SetDefault|Insert -Name [-EntryID draft]` |
| Reminders | `outlook-reminders.ps1` | `-Action List|Snooze|Dismiss -Caption|-Index -Minutes -Days -All` |
| Out of office | `outlook-out-of-office.ps1` | `-Status | -Enable | -Disable [-InternalMessage -ExternalMessage -StartTime -EndTime] -Confirm` |
| Accounts, mailboxes, sizes | `outlook-accounts-list.ps1`, `outlook-stores.ps1` (`-TopFolders 15`) | |
| Save / print / import | `outlook-save-as.ps1` (`-Format msgunicode|msg|txt|html|mht|rtf|doc|template`), `outlook-print.ps1 -Confirm`, `outlook-import-msg.ps1 -Path` | |
| Recall | `outlook-recall.ps1` | `-EntryID -OpenRecallDialog` (opens Outlook's recall dialog) or `-DeleteUnread -Confirm` (Sent copy only) |

## Output

Text is the default and keeps everything needed to act: `EntryID`, `ConversationID`, folder, `From` with address, `To`/`CC`, date with UTC offset (`2026-10-22 09:15 -04:00`), status, importance, flag, categories, attachments, preview or body. List views print the folder once at the top. Add `-Json` to any script for one compact JSON object (dates ISO 8601 with offset; `StoreID` included).

**Links are citations.** Bodies show links as `[URL-001]`, `[URL-002]`, ... numbered by first appearance (repeats reuse the number). Full views (`read-body`, `conversation`, `selection -Body`) add a table: `URL-001  teams.microsoft.com  "Join meeting"`. Microsoft Safe Links are decoded to the real site. Get the full address with `outlook-links.ps1 -EntryID <id> -Ref URL-001` (numbering is computed from the message, so references from earlier calls still resolve), or show all links inline with `-Links full`. On `outlook-read-body.ps1`, `-StripLinks:$false` from older callers means `-Links full`; elsewhere `-StripLinks` is accepted and ignored.

**Searching body text:** `-BodyContains` is an exact substring match over the emails that pass the other filters (nothing missed; about 2 s for 30 days). `-FastBody` uses Outlook's content index instead: much faster (under 1 s for a year) but matches whole words and can miss mail not indexed yet.

## Replies in the thread (HTML)

`outlook-reply.ps1` / `outlook-forward.ps1` put your text at the top of the reply, inside Outlook's own body and above the signature and the quoted thread, so the thread is never overwritten. Plain `-Body` text is rendered the way Outlook writes new mail (Aptos 11pt `<div class="elementToProof">` blocks); `-HTML` / `-HtmlFile` take your own HTML. The text sits in a marked block, so `-DraftEntryID <draft> -Body "..." -Confirm` replaces it in the same draft instead of creating another. Outlook inserts the account's reply signature automatically; `-RemoveAutoSignature` drops it (use it when your text already ends with a signature), or add a saved one with `outlook-signatures.ps1 -Action Insert`.

Typical flow: `outlook-reply.ps1 -EntryID <id> -HtmlFile reply.html -ReplyAll -RemoveAutoSignature` (draft saved, EntryID printed) -> user reviews in Outlook -> `outlook-send-draft.ps1 -EntryID <draft> -Confirm`.

## Confirmation

Every irreversible or outward-facing action is available, and each one checks a confirmation flag. Without `-Confirm` it previews exactly what would happen and does nothing (exit 0).

Gated actions, by the key used in config.json:

| Key | Actions |
|---|---|
| `send` | send, send-as, send-draft, schedule-send, voting, request-receipt, reply/forward `-Send` |
| `delete` | delete (and `-Permanent`), recall `-DeleteUnread` |
| `edit-draft` | replacing draft text, removing draft attachments |
| `rules` | rules create, enable/disable, rename, reorder, delete, run now |
| `folders` | folder rename, move, delete, empty, mark all read |
| `bulk` | every `outlook-bulk` action |
| `category-delete` | deleting a category from the master list |
| `out-of-office` | turning auto-replies on or off |
| `junk` | block sender, safe sender |
| `signatures` | replacing, deleting or setting the default signature |
| `reminders` | dismissing reminders |
| `conversation` | always-move, always-delete, always-categorize a thread |
| `print` | printing |
| `overwrite` | save-as `-Overwrite` of an existing file |

Not gated (reversible or creates something new): find/read, draft, reply/forward drafts, move, copy, mark read, flag, categories, importance, sensitivity, appending to drafts, folder create/copy.

Switching confirmation off (the user's choice):
- per call: `-Confirm`
- for this shell: `$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "0"` (`"1"` forces it on)
- permanently: `%APPDATA%\outlook-suite\config.json` with `{ "requireConfirm": false }`, or per action: `{ "requireConfirm": true, "actions": { "send": true, "bulk": false } }`

## Speed: the warm host

`scripts/outlook-call.ps1 <script> [arguments]` runs any script through one long-lived PowerShell + Outlook session (`outlook-host.ps1`, started automatically on first use, stops after 30 idle minutes, current Windows user only). Same output and exit code as a direct run. Call it with `&` from the PowerShell you are already in: that takes about half the time of a direct run (find 7 days: ~320 ms vs ~600 ms). Starting a new `pwsh -File outlook-call.ps1` per call gains nothing, because starting pwsh costs what the host saves. The host needs PowerShell 7 (`pwsh`); under Windows PowerShell 5.1 `outlook-call.ps1` simply runs the script directly. All scripts themselves work in both. `outlook-call.ps1 --batch requests.json` runs many requests in one round trip; `--ping`, `--stop` manage the host. Concurrent callers are queued (Outlook serves one call at a time).

```powershell
& scripts/outlook-call.ps1 find -Days 7 -UnreadOnly
& scripts/outlook-call.ps1 read-body -EntryID "00000000..."
```

## Outlook behaviours worth knowing

- Drafts cannot be marked "flag complete" (Outlook rule); `outlook-unflag.ps1 -Complete` says so. Clearing works.
- "Has attachment" uses Outlook's attachment flag: inline images such as signature logos do not count.
- Rules: if Outlook reports the mailbox's rules are "in conflict", no rule can be read or changed by script until the user opens File > Manage Rules & Alerts once and picks which copy to keep. Some accounts (Outlook.com) refuse move/copy actions in scripted rules, and no account accepts moving into a folder inside Deleted Items.
- Exchange does not let scripts permanently delete folders that are already in Deleted Items.
- Out-of-office text/schedule, and reading the junk lists, need the Exchange Online module (`Install-Module ExchangeOnlineManagement -Scope CurrentUser; Connect-ExchangeOnline`); on/off and block/safe sender work without it.
- With Microsoft 365 roaming signatures, Outlook may keep the cloud default instead of `-Action SetDefault`.
- Reminders are asynchronous: Outlook registers and pops them up on its own schedule.
- Debugging a script: `$env:OUTLOOK_SUITE_DEBUG = "1"` adds the script stack trace to `Error:` lines.

## Tests

`tests/` (run against the live profile; they print counts and PASS/FAIL only, write only to test items and a hidden `Deleted Items/_suite-test` folder, and send nothing): `Test-MailParity.ps1` (same results as a baseline copy), `Invoke-MailBench.ps1`, `Test-MailWrites.ps1`, `Test-MailCompose.ps1`, `Test-MailFolders.ps1`, `Test-MailRules.ps1 -Mailbox <test mailbox>`, `Test-MailExtras.ps1`, `Test-MailHost.ps1`, `Test-MailPS51.ps1` (same results under Windows PowerShell 5.1 and PowerShell 7).
