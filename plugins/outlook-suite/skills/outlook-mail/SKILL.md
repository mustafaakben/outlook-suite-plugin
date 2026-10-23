---
name: outlook-mail
description: "Manage classic Outlook email on Windows through PowerShell COM: triage (untagged, by tag, not replied, Copilot priority), Copilot insights (priority ranking, headline, reason, suggested replies as drafts), find, read, full threads, reply/forward in thread with HTML, draft, send, move, bulk actions, folders, rules, categories, flags, attachments, headers, junk, signatures, reminders, out-of-office. Use for any Outlook mail task on the local Windows PC. Fast (one Table query per search), agent-friendly (EntryIDs everywhere, -Json, exit codes), with a confirm switch on every irreversible action."
---

# Outlook Mail

56 PowerShell scripts in `scripts/` that drive classic Outlook (running on this Windows PC) through COM. Every script works standalone (`pwsh -File scripts/<name>.ps1 ...`); a warm host makes repeated calls about twice as fast.

## How to work

1. **Find first, then act by EntryID.** `outlook-triage.ps1` (what needs attention), `outlook-find.ps1` (or `outlook-read`, `outlook-search`, `outlook-search-advanced`, `outlook-selection`) prints an `EntryID` for every email. Pass it to any action script with `-EntryID`. `-Index` still exists but positions shift as mail arrives; do not use it for actions.
2. **EntryIDs can change when an email moves.** Scripts that move or copy print the new EntryID (and the old one). Use the new one afterwards.
3. **Folders are names or paths:** `Inbox`, `Sent`, `Drafts`, `Deleted`, `Junk`, `Outbox`, `Inbox/Projects/2026`, or another mailbox: `someone@x.edu/Inbox`. `outlook-folders.ps1` lists the paths. An unknown folder is an error (exit 1), never a silent fallback to Inbox.
4. **Read the output contract:** success exits 0; any failure prints `Error: <reason>` (often a `Hint:` line) and exits 1. A gated action without confirmation prints `PREVIEW ONLY: ... was NOT done.` and exits 0.
5. **Ask before irreversible actions** (see Confirmation). When the user has already said to go ahead, pass `-Confirm`.

## Which script when

| Need | Script | Key parameters |
|---|---|---|
| What needs attention: untagged, by tag, not replied, Copilot priority | `outlook-triage.ps1` | `untagged|tagged|reply -Days -DateFrom -DateTo -All -Unread -Read -Category "@A,@B" -Untagged -NotReplied -Replied -Priority High,Normal,Low -From -Subject -Folder -Sort Newest|Oldest|Priority -Limit -Json` (see Triage) |
| Mail ranked by Copilot priority | `outlook-priority.ps1` | `-Days 1 -DateFrom -DateTo -All -Unread -Read -Category -Untagged -Priority High,Normal,Low,None -From -Subject -Folder | -AllFolders -Limit [-Why] [-Summary] -Json` (see Copilot insights) |
| Copilot's suggested replies | `outlook-smart-reply.ps1` | `-EntryID [-Use 1-3 [-ReplyAll] [-RemoveAutoSignature]]`; list: `-Days 2 -DateFrom -DateTo -UnreadOnly -From -Subject -Folder -Limit -MaxScan -Json` |
| List / filter emails | `outlook-find.ps1` | `-Folder -Days -DateFrom -DateTo -From -To -Subject -BodyContains [-FastBody] -UnreadOnly -ReadOnly -HasAttachment -Flagged -Importance -Category -Priority High,Normal,Low -Limit -Detailed -Json` |
| Latest emails | `outlook-read.ps1` | `-Days 1 -Limit 20 -UnreadOnly -Folder -Priority High,Normal,Low` |
| Quick text search | `outlook-search.ps1` | `-Subject -From -Body [-FastBody] -Days -Folder` |
| Search many folders / whole mailbox / all mailboxes | `outlook-search-advanced.ps1` | `-Folders "Inbox;Archive" -IncludeSubfolders -WholeMailbox -AllMailboxes` + find filters, `-SaveAs "name"` (Outlook Search Folder) |
| Everything about a topic: mail, calendar, contacts, tasks, notes AND local files | `outlook-everything.ps1` | `"words" [-Kind Mail,Calendar,Contact,Task,Note,File] -Days -From -Path <folder> -Top 10 -NoResolve -Json` (see Everything search) |
| What the user has selected in Outlook | `outlook-selection.ps1` | `-Body` |
| One email in full | `outlook-read-body.ps1` | `-EntryID [-Links domain|full|none] [-Format html] -Json` |
| Whole thread, oldest first | `outlook-conversation.ps1` | `-EntryID [-TrimQuoted] [-Preview 300] [-HeadersOnly]` |
| Full URLs behind `[URL-00N]` | `outlook-links.ps1` | `-EntryID [-Ref URL-003]` |
| Internet headers, SPF/DKIM/DMARC, phishing signs | `outlook-headers.ps1` | `-EntryID [-Raw]` |
| Attachments: list / save | `outlook-attachment-save.ps1` | `-EntryID -ListOnly | -Path -AttachmentIndex -IncludeInline` |
| Attachments: read the text (Word, Excel, PowerPoint, PDF, .msg, text) | `outlook-attachment-read.ps1` | `-EntryID [-AttachmentIndex] -MaxChars 8000 -MaxRows 200 -IncludeInline -Json`, or `-FilePath <local file>` (see Reading attachments) |
| New draft | `outlook-draft.ps1` | `-To -Subject -Body [-Native|-HTML|-HtmlFile] -CC -BCC -Attachment -Account -Importance -Sensitivity` |
| Change a draft | `outlook-draft-update.ps1` | `-EntryID -To/-AddTo -CC/-AddCC -Subject -Body [-AppendBody] -AddAttachment -RemoveAttachment -Importance -Account` |
| Send a reviewed draft | `outlook-send-draft.ps1` | `-EntryID -Confirm` |
| Send new email | `outlook-send.ps1` / `outlook-send-as.ps1` (`-Account`) | same as draft + `-Confirm` |
| Send later | `outlook-schedule-send.ps1` | `-SendAt "2026-11-02 09:00" -Confirm` |
| Voting buttons / receipts | `outlook-voting.ps1` (`-Options "A;B"`), `outlook-request-receipt.ps1` | `-Confirm` |
| Reply / reply all in thread | `outlook-reply.ps1` | `-EntryID -Body | -HTML/-HtmlFile | -SmartReply 1-3 [-ReplyAll] [-RemoveAutoSignature] [-DraftEntryID] [-Send -Confirm]` |
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

## Everything search (`outlook-everything.ps1`)

One full-text query over the Windows Search index, the same index behind Outlook's and Explorer's search boxes. It covers every indexed mail and calendar item, contacts, tasks, notes, and local files (Documents, Dropbox, and whatever else Indexing Options lists), usually in 1 to 3 s.

- Words are ANDed; `"quoted phrases"` stay together; `budg*` is a prefix match.
- Each kind is queried on its own (`-Top` each, default 10), so files cannot crowd out mail.
- Outlook hits come back with a real `EntryID` (the index's short-term id is opened through Outlook), ready for `outlook-read-body.ps1`, `outlook-attachment-read.ps1` or `outlook-calendar-prep.ps1`. A hit moved since indexing shows no EntryID. `-NoResolve` skips that lookup.
- `-Days` filters by received date (mail), start (calendar; a recurring series counts by its first start) or last change (others). `-From` applies to mail; `-Path` limits files to one folder tree.
- Only what Windows Search has indexed is found. Mail from the last minutes may not be indexed yet; use `outlook-find.ps1` for that.

## Reading attachments (`outlook-attachment-read.ps1`)

Returns the text of each attachment (inline images skipped) so the agent can triage or prep without opening files. Read-only. Each file is copied to a private temp folder that is deleted afterwards.

| Type | How |
|---|---|
| `.docx` `.xlsx` `.pptx` | Read from the file's XML; no Office app starts. Word tables come out as tab-separated cells, slides as `--- Slide N ---` with `[Notes]`, sheets as `--- Sheet: name (N rows) ---` with tab-separated cells (`-MaxRows`, default 200). |
| `.pdf` `.doc` `.rtf` `.odt` | Word, hidden and read-only, macros forced off. Scanned PDFs have no text. |
| `.xls` / `.ppt` | Excel / PowerPoint, the same way. |
| `.msg` (attached email) | Headers and body through Outlook. |
| `.txt` `.csv` `.md` `.html` and other text | Read directly; HTML tags stripped. |
| Images, archives, other binaries | Listed as not readable. |

Text is cut at `-MaxChars` (default 8000, `0` for all) with the full length reported. `-FilePath` reads a local file the same way, e.g. a File hit from `outlook-everything.ps1`. Password-protected files fail at once instead of prompting.

## Triage

`outlook-triage.ps1` lists facts for deciding what to do with mail. It applies no content rules and calls no AI; the agent reading the list decides. Three views, each a set of defaults:

| View | Defaults | Answers |
|---|---|---|
| `untagged` (default) | Inbox, last 3 days, unread, no category | what nobody has sorted yet |
| `tagged` | Inbox, last 3 days, read and unread, any category, grouped by tag set | where everything stands |
| `reply` | Inbox, last 7 days, not replied (any tag or none), oldest first, "waiting Nd" | what the user has not answered |

Every filter works with every view and overrides its default: `-Days N` (`-Days 0` = no date limit), `-DateFrom/-DateTo`, `-All` (no date limit and read + unread; an explicit `-Days` still wins), `-Unread`/`-Read`, `-Category "@Action,@Review"` (any of them) or `-Untagged`, `-NotReplied`/`-Replied`, `-Priority High,Normal,Low`, `-From`, `-Subject`, `-Folder`, `-Sort Newest|Oldest|Priority`, `-Limit`, `-Json`.

Each row: date, exact subject, sender, tags, unread, replied / replied all / forwarded, Copilot priority, whether the user is in To or only CC and the recipient count, flag and due date, preview, EntryID. The last line counts matches, unread, untagged, not replied, Copilot high and flagged.

- **Replied** is Outlook's own record of the last action on the email (reply, reply all, forward). Forwarded is not replied.
- **Copilot priority** is the High / Low arrow Copilot shows in Outlook ("Marked as high priority by Copilot"), read from the message property `CopilotInboxScore` (5 high, 3 normal, 1 low). It is empty on mail Copilot did not score.

```powershell
& scripts/outlook-call.ps1 triage                                   # unread, untagged, last 3 days
& scripts/outlook-call.ps1 triage tagged -Category "@Action,@Review" -Unread
& scripts/outlook-call.ps1 triage reply -Priority High -Days 10 -Json
```

Each row also shows Copilot's one-line headline when it wrote one. `-Copilot` adds Copilot's reason and suggested replies for the rows shown (one extra read per email).

## Copilot insights

Copilot stores its analysis on each email in Exchange. These scripts read it; they do not call Copilot or any AI. The fields are undocumented Microsoft properties (`PSETID_Common`, names in `$script:CopilotProp` in `_common.ps1`) and may change. Missing means Copilot did not analyze that email, never a guess.

| Field | What it is | Coverage (typical) |
|---|---|---|
| Priority | High / Normal / Low, the arrow Copilot shows in Outlook (`CopilotInboxScore` 5 / 3 / 1) | almost every recent email |
| Headline | Copilot's one-line summary of why the email matters | about 1 in 5 |
| Reason | Copilot's explanation of the priority | emails with a headline |
| Suggested replies | up to 3 short replies, in Copilot's rank order | about 1 in 4 |

- `outlook-priority.ps1` ranks mail High > Normal > Low > not scored, newest first within each level (default: last 1 day). `-Why` adds the reason, `-Summary` prints counts only. `-AllFolders` covers all received mail: the Inbox and every folder rules move mail to (not Sent, Drafts, Deleted Items, Junk, Outbox or system folders); each row then shows its folder. Mail that rules file before it reaches the Inbox is often not scored by Copilot.
- `outlook-smart-reply.ps1 -EntryID <id>` shows the suggestions. `-Use N` saves a reply draft in the thread with suggestion N as the text (same path as `outlook-reply.ps1 -SmartReply N`). It never sends; use `outlook-send-draft.ps1 -Confirm`.
- `outlook-read-body.ps1` prints a "Copilot" block (priority, headline, reason, suggestions) and `-Json` carries it as `Copilot`.
- `outlook-find.ps1` / `outlook-read.ps1` take `-Priority` and show `CopilotPriority` on each row.

```powershell
& scripts/outlook-call.ps1 priority -Days 1 -Unread               # today's unread, ranked
& scripts/outlook-call.ps1 priority -Days 7 -Priority High -Why    # only High, with Copilot's reason
& scripts/outlook-call.ps1 smart-reply -EntryID <id> -Use 2        # draft from suggestion 2
```

## Output

Text is the default and keeps everything needed to act: `EntryID`, `ConversationID`, folder, `From` with address, `To`/`CC`, date with UTC offset (`2026-10-22 09:15 -04:00`), status, importance, flag, categories, attachments, preview or body. List views print the folder once at the top. Add `-Json` to any script for one compact JSON object (dates ISO 8601 with offset; `StoreID` included).

**Links are citations.** Bodies show links as `[URL-001]`, `[URL-002]`, ... numbered by first appearance (repeats reuse the number). Full views (`read-body`, `conversation`, `selection -Body`) add a table: `URL-001  teams.microsoft.com  "Join meeting"`. Microsoft Safe Links are decoded to the real site. Get the full address with `outlook-links.ps1 -EntryID <id> -Ref URL-001` (numbering is computed from the message, so references from earlier calls still resolve), or show all links inline with `-Links full`. On `outlook-read-body.ps1`, `-StripLinks:$false` from older callers means `-Links full`; elsewhere `-StripLinks` is accepted and ignored.

**Searching body text:** `-BodyContains` is an exact substring match over the emails that pass the other filters (nothing missed; about 2 s for 30 days). `-FastBody` uses Outlook's content index instead: much faster (under 1 s for a year) but matches whole words and can miss mail not indexed yet.

## Replies in the thread (HTML)

`outlook-reply.ps1` / `outlook-forward.ps1` put your text at the top of the reply, inside Outlook's own body and above the signature and the quoted thread, so the thread is never overwritten. Plain `-Body` text is rendered the way Outlook writes new mail (Aptos 11pt `<div class="elementToProof">` blocks); `-HTML` / `-HtmlFile` take your own HTML. The text sits in a marked block, so `-DraftEntryID <draft> -Body "..." -Confirm` replaces it in the same draft instead of creating another. Outlook inserts the account's reply signature automatically; `-RemoveAutoSignature` drops it (use it when your text already ends with a signature), or add a saved one with `outlook-signatures.ps1 -Action Insert`.

Typical flow: `outlook-reply.ps1 -EntryID <id> -HtmlFile reply.html -ReplyAll -RemoveAutoSignature` (draft saved, EntryID printed) -> user reviews in Outlook -> `outlook-send-draft.ps1 -EntryID <draft> -Confirm`.

## Confirmation

Every irreversible or outward-facing action is available, and each one checks a confirmation flag. Without `-Confirm` it previews exactly what would happen and does nothing (exit 0).

Gated actions, by the key used in config.json. The keys cover every skill in the suite, not only mail:

| Key | Actions |
|---|---|
| `send` | mail: send, send-as, send-draft, schedule-send, voting, request-receipt, reply/forward `-Send`. Calendar: create and recurring with attendees, hold-to-invite, update `-SendUpdate`, forward `-Send`, cancel. Tasks: assign |
| `respond` | calendar respond (accept, tentative, decline) |
| `delete` | mail delete (and `-Permanent`), recall `-DeleteUnread`; calendar, contact, task and note delete; contacts dedupe `-Merge` |
| `import` | contacts import |
| `edit-draft` | replacing draft text, removing draft attachments |
| `rules` | rules create, enable/disable, rename, reorder, delete, run now |
| `folders` | folder rename, move, delete, empty, mark all read |
| `bulk` | every `outlook-bulk` action; tasks triage |
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
- If Outlook is closed or crashes, the next call reconnects (the warm host notices too). Scripts start Outlook with its window, never hidden: after a crash Outlook may want to show a recovery prompt, and a hidden Outlook would wait on it forever. If Outlook does not come up within 90 s the call fails with `Error:`; open Outlook once by hand.
- Debugging a script: `$env:OUTLOOK_SUITE_DEBUG = "1"` adds the script stack trace to `Error:` lines.

## Tests

`tests/` (run against the live profile; they print counts and PASS/FAIL only, write only to test items and a hidden `Deleted Items/_suite-test` folder, and send nothing): `Test-MailParity.ps1` (same results as a baseline copy), `Invoke-MailBench.ps1`, `Test-MailWrites.ps1`, `Test-MailCompose.ps1`, `Test-MailFolders.ps1`, `Test-MailRules.ps1 -Mailbox <test mailbox>`, `Test-MailExtras.ps1`, `Test-MailHost.ps1`, `Test-MailPS51.ps1` (same results under Windows PowerShell 5.1 and PowerShell 7), `Test-MailTriage.ps1` (triage views against independent item-by-item counts), `Test-MailSmartReply.ps1` (suggested replies against independent reads; one test draft, deleted with `-Confirm`).
