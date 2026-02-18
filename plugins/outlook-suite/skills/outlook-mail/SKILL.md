---
name: outlook-mail
description: "Manage Outlook email via PowerShell: read, search, draft, send, reply, forward, move, delete, flag, categorize, schedule send, rules, attachments, and out-of-office. Requires Outlook running."
---

# Outlook Mail

Manage Microsoft Outlook email using PowerShell COM objects.

## Prerequisites

- Microsoft Outlook installed and running
- PowerShell 5.1+ (included in Windows 10/11)

## Safety Rules

1. **Never auto-send** - Always create drafts first for user review
2. `outlook-send.ps1` requires `-Confirm` flag to actually send
3. `outlook-send-as.ps1`, `outlook-reply.ps1`, `outlook-forward.ps1`, `outlook-voting.ps1`, `outlook-request-receipt.ps1` also require `-Confirm`
4. Confirm with user before any send operation
5. `outlook-delete.ps1` requires `-Confirm` to delete; `-Permanent` attempts identity-verified purge and safe-fails when identity cannot be verified
6. `outlook-recall.ps1 -DeleteUnread` only removes from your Sent Items and requires `-ConfirmDeleteFromSent`

## Email Targeting

**Always use EntryID as the primary method to target emails.** Index is a fallback only.

- **Listing scripts** (`outlook-find.ps1`, `outlook-read.ps1`, `outlook-search.ps1`, `outlook-search-advanced.ps1`) output an `EntryID` for each email
- **Action scripts** accept `-EntryID` as the preferred parameter (alternative to `-Index`)
- EntryID is stable for exact targeting in the current mailbox context, but can change if an item is moved/copied across stores
- Index numbers are volatile — they can shift between commands if new emails arrive

**Workflow:**
1. Run `outlook-find.ps1`, `outlook-read.ps1`, `outlook-search.ps1`, or `outlook-search-advanced.ps1` to get EntryIDs
2. Use `-EntryID` on any action script to target the exact email

```powershell
# Step 1: Find emails matching criteria, get stable EntryIDs
& "./scripts/outlook-find.ps1" -Days 1 -UnreadOnly

# Step 2: Act on email using stable EntryID
& "./scripts/outlook-reply.ps1" -EntryID "00000000..." -Body "Got it, thanks!"
```

## Link Stripping

Scripts that output email body content support a `-StripLinks` parameter (type: `bool`, default: `$true`). When enabled, all URLs in the output are replaced with `[URL]` to reduce context window bloat from long tracking URLs, Safe Links wrappers, and marketing links.

- **Default ON**: No action needed — URLs are stripped automatically
- **To preserve URLs**: Pass `-StripLinks $false`

**Scripts with `-StripLinks` support:**
`outlook-read-body.ps1`, `outlook-find.ps1`, `outlook-draft.ps1`, `outlook-forward.ps1`, `outlook-reply.ps1`, `outlook-out-of-office.ps1`, `outlook-request-receipt.ps1`, `outlook-rules-create.ps1`, `outlook-rules-list.ps1`, `outlook-schedule-send.ps1`, `outlook-send-as.ps1`, `outlook-send.ps1`, `outlook-voting.ps1`

## Scripts Overview

**Find and Target**
- `outlook-find.ps1` - Find emails by criteria, return stable EntryIDs for action scripts

**Read and Search**
- `outlook-read.ps1` - List recent inbox emails
- `outlook-read-body.ps1` - Read full email body by EntryID or index
- `outlook-search.ps1` - Search by subject, sender, or body
- `outlook-search-advanced.ps1` - Multi-filter advanced search

**Compose and Send**
- `outlook-draft.ps1` - Create draft email
- `outlook-draft-update.ps1` - Modify existing draft
- `outlook-send.ps1` - Send email directly (requires -Confirm)
- `outlook-send-as.ps1` - Send from specific account (requires -Confirm)
- `outlook-reply.ps1` - Reply to email (requires -Confirm to send)
- `outlook-forward.ps1` - Forward email (requires -Confirm to send)
- `outlook-voting.ps1` - Send email with voting buttons (requires -Confirm)

**Organize**
- `outlook-move.ps1` - Move email to folder
- `outlook-delete.ps1` - Delete email (requires -Confirm)
- `outlook-mark-read.ps1` - Mark email read or unread
- `outlook-conversation.ps1` - View email conversation thread
- `outlook-print.ps1` - Print email

**Flags, Importance and Sensitivity**
- `outlook-flag.ps1` - Flag email for follow-up
- `outlook-unflag.ps1` - Remove or complete flag
- `outlook-set-importance.ps1` - Set email priority level
- `outlook-set-sensitivity.ps1` - Set email sensitivity level
- `outlook-save-as.ps1` - Export email to file

**Categories**
- `outlook-list-categories.ps1` - List all categories
- `outlook-create-category.ps1` - Create new category
- `outlook-assign-category.ps1` - Assign category to email
- `outlook-remove-category.ps1` - Remove category from email
- `outlook-delete-category.ps1` - Delete category from master list

**Scheduling and Receipts**
- `outlook-schedule-send.ps1` - Schedule delayed send
- `outlook-request-receipt.ps1` - Send with read/delivery receipts
- `outlook-recall.ps1` - Recall guidance and optional Sent Items deletion

**Account and Rule Management**
- `outlook-folders.ps1` - List all mail folders
- `outlook-accounts-list.ps1` - List configured email accounts
- `outlook-rules-list.ps1` - List email rules
- `outlook-rules-create.ps1` - Create new email rule
- `outlook-out-of-office.ps1` - Configure auto-replies
- `outlook-attachment-save.ps1` - Save email attachments

## Find and Target

Search and filter emails to retrieve stable EntryIDs for use with action scripts.

### outlook-find.ps1
Find emails matching search criteria and return their EntryIDs. This is the primary way to get stable email identifiers before performing actions (reply, forward, delete, etc.).

**Required Parameters:** None

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search |
| `-DateFrom` | datetime | — | Start date (overrides -Days when used with -DateTo) |
| `-DateTo` | datetime | — | End date (inclusive, end-of-day) |
| `-From` | string | — | Sender name or email address contains |
| `-To` | string | — | Recipient name or email contains |
| `-Subject` | string | — | Subject line contains |
| `-BodyContains` | string | — | Body text contains (plain text search) |
| `-UnreadOnly` | switch | off | Only unread emails |
| `-ReadOnly` | switch | off | Only read emails |
| `-HasAttachment` | switch | off | Only emails with attachments |
| `-Flagged` | switch | off | Only flagged/follow-up emails |
| `-Importance` | string | — | Filter by importance: High, Normal, Low |
| `-Category` | string | — | Only emails assigned this category |
| `-Folder` | string | Inbox | Folder to search (e.g., "Sent Items", "Archive") |
| `-Limit` | int | 50 | Max results to return |
| `-StripLinks` | bool | `$true` | Replace URLs in body snippets with `[URL]` |

**Output per result:** EntryID, Subject, From (name + email), Date, Read/Unread status, Body snippet (~150 chars), Attachments (count + filenames), Importance (if non-normal), Flag status, Categories.

**Examples:**

```powershell
# Find all unread emails from today
& "./scripts/outlook-find.ps1" -Days 1 -UnreadOnly

# Find emails about "lunch" from today
& "./scripts/outlook-find.ps1" -Days 1 -Subject "lunch"

# Find emails with "invoice" in the body
& "./scripts/outlook-find.ps1" -Days 30 -BodyContains "invoice"

# Find flagged high-importance emails
& "./scripts/outlook-find.ps1" -Days 7 -Flagged -Importance High

# Find emails from a specific sender with attachments
& "./scripts/outlook-find.ps1" -From "john@example.com" -HasAttachment

# Find read emails in a specific category
& "./scripts/outlook-find.ps1" -ReadOnly -Category "Project Alpha"

# Find emails sent to a specific person
& "./scripts/outlook-find.ps1" -To "team@company.com" -Days 14

# Search in Sent Items folder
& "./scripts/outlook-find.ps1" -Folder "Sent Items" -Days 3 -Subject "report"

# Then act on a found email using its EntryID
& "./scripts/outlook-reply.ps1" -EntryID "00000000..." -Body "Got it, thanks!"
```

---

## Read and Search

Read recent emails, view full message bodies, and search with simple or advanced filters.

### outlook-read.ps1
List recent emails from the Inbox with sender name, email address, date, unread status, and EntryID. Results include a stable EntryID for use with action scripts (e.g., `outlook-read-body.ps1 -EntryID "..."`). Also numbered with an index as a fallback.

**Required Parameters:** None

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 1 | Number of days to look back |
| `-Limit` | int | 20 | Maximum number of emails to display |
| `-UnreadOnly` | switch | off | Show only unread emails |

**Examples:**

```powershell
# Read today's emails (default: last 1 day, limit 20)
& "./scripts/outlook-read.ps1"

# Read last 2 days of emails
& "./scripts/outlook-read.ps1" -Days 2

# Read last 7 days, limit to 50
& "./scripts/outlook-read.ps1" -Days 7 -Limit 50

# Show only unread emails from last 7 days
& "./scripts/outlook-read.ps1" -Days 7 -UnreadOnly
```

### outlook-read-body.ps1
Read the full body content of a specific email. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (1-based, fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | Number of days to look back (only used with -Index) |
| `-StripLinks` | bool | `$true` | Replace URLs in email body with `[URL]` |

**Examples:**

```powershell
# Read body of the 3rd most recent email
& "./scripts/outlook-read-body.ps1" -Index 3

# Read body with wider date range
& "./scripts/outlook-read-body.ps1" -Index 5 -Days 14
```

### outlook-search.ps1
Search emails by subject, sender, or body text. At least one search filter (`-Subject`, `-From`, or `-Body`) is required. Uses server-side date pre-filtering with client-side matching for text filters. `-From` matches sender display name or sender email address. Results include EntryID for stable follow-up actions.

**Required Parameters:** At least one of `-Subject`, `-From`, or `-Body` must be provided.

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | | Search in email subject |
| `-From` | string | | Search by sender name or sender email address |
| `-Body` | string | | Search in email body (client-side, slower on large result sets) |
| `-Days` | int | 30 | Number of days to look back |
| `-Limit` | int | 20 | Maximum results to display |

**Examples:**

```powershell
# Search by subject
& "./scripts/outlook-search.ps1" -Subject "meeting"

# Search by sender with custom date range
& "./scripts/outlook-search.ps1" -From "John" -Days 14

# Search by body content
& "./scripts/outlook-search.ps1" -Body "quarterly report" -Days 30

# Combine multiple filters
& "./scripts/outlook-search.ps1" -Subject "review" -From "Jane" -Days 7

# Limit results
& "./scripts/outlook-search.ps1" -Subject "report" -Limit 5
```

### outlook-search-advanced.ps1
Advanced multi-filter search with attachment, importance, flag, category, recipient, and date range filters. Uses server-side JET filtering for dates and client-side matching for all other criteria. `-From` matches sender display name or sender email address. Searches up to 1000 items per run. Supports recursive folder lookup by name. Results include EntryID for stable follow-up actions.

**Required Parameters:** None (but at least one filter recommended)

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Subject` | string | | Search in email subject |
| `-From` | string | | Search by sender name or sender email address |
| `-To` | string | | Search by recipient name |
| `-Body` | string | | Search in email body (client-side, slower on large result sets) |
| `-HasAttachment` | switch | off | Only emails with attachments |
| `-Unread` | switch | off | Only unread emails |
| `-Flagged` | switch | off | Only flagged emails |
| `-Importance` | string | | Filter by importance: High, Normal, Low |
| `-Category` | string | | Filter by category name |
| `-DateFrom` | datetime | | Start date (e.g., "2026-01-01") |
| `-DateTo` | datetime | | End date, inclusive full day (e.g., "2026-01-31" includes all of Jan 31) |
| `-Folder` | string | Inbox | Search in specific folder (Inbox, Sent, Drafts, or any folder name) |
| `-Limit` | int | 20 | Maximum results to display |

**Result Markers:** `[*]` unread, `[A]` has attachment, `[F]` flagged, `[!]` high importance

**Examples:**

```powershell
# Find unread emails with attachments
& "./scripts/outlook-search-advanced.ps1" -HasAttachment -Unread

# Find high importance emails from a sender
& "./scripts/outlook-search-advanced.ps1" -From "John" -Importance High

# Search within a date range (both dates inclusive)
& "./scripts/outlook-search-advanced.ps1" -DateFrom "2026-01-01" -DateTo "2026-01-31"

# Search flagged emails in a date range
& "./scripts/outlook-search-advanced.ps1" -DateFrom "2026-01-01" -DateTo "2026-01-31" -Flagged

# Search by recipient with attachments
& "./scripts/outlook-search-advanced.ps1" -To "team@example.com" -HasAttachment

# Search in a specific folder with category filter
& "./scripts/outlook-search-advanced.ps1" -Subject "Report" -Category "Important" -Folder "Sent"

# Combine text and property filters
& "./scripts/outlook-search-advanced.ps1" -Body "quarterly" -Importance High -Unread -Limit 50
```

## Compose and Send

Create drafts, send emails, reply, forward, and send voting polls. All send operations require `-Confirm`.

### outlook-draft.ps1
Create a draft email in the Drafts folder. Safe operation — nothing is sent. Supports plain text and HTML body, CC/BCC recipients, and one or multiple file attachments.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address(es) |
| `-Subject` | string | Email subject line |
| `-Body` | string | Email body content (plain text or HTML) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CC` | string | | CC recipient(s) |
| `-BCC` | string | | BCC recipient(s) |
| `-Attachment` | string[] | | File path(s) to attach. Warns if file not found. |
| `-HTML` | switch | off | Treat body as HTML instead of plain text |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**

```powershell
# Create a simple plain text draft
& "./scripts/outlook-draft.ps1" -To "john@example.com" -Subject "Meeting" -Body "Let's meet tomorrow"

# Create draft with CC and BCC
& "./scripts/outlook-draft.ps1" -To "john@example.com" -Subject "Update" -Body "Status update" -CC "team@example.com" -BCC "boss@example.com"

# Create HTML draft
& "./scripts/outlook-draft.ps1" -To "john@example.com" -Subject "Report" -Body "<h1>Monthly Report</h1><p>See below.</p>" -HTML

# Create draft with single attachment
& "./scripts/outlook-draft.ps1" -To "john@example.com" -Subject "Files" -Body "Attached" -Attachment "C:\docs\report.pdf"

# Create draft with multiple attachments
& "./scripts/outlook-draft.ps1" -To "john@example.com" -Subject "Files" -Body "See attached" -Attachment "C:\a.pdf","C:\b.xlsx"
```

### outlook-draft-update.ps1
Modify an existing draft in the Drafts folder. Update recipients, subject, body (replace or append), and add attachments. Shows old vs new values for each change. Use `-EntryID` (preferred) or `-Index` (fallback). Drafts are sorted by most recently modified when using -Index.

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable draft ID from listing output (preferred) |
| `-Index` | int | Draft number (1-based, sorted by most recently modified, fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-To` | string | | Replace recipient(s) |
| `-CC` | string | | Replace CC recipient(s) |
| `-BCC` | string | | Replace BCC recipient(s) |
| `-Subject` | string | | Replace subject line |
| `-Body` | string | | Replace or append body content |
| `-AppendBody` | switch | off | Append body text instead of replacing |
| `-AddAttachment` | string[] | | File path(s) to attach. Warns if file not found. |
| `-HTML` | switch | off | Treat body as HTML when replacing or appending |
| `-Limit` | int | 20 | Maximum drafts to scan when finding by index |

**Examples:**

```powershell
# Update the subject of the most recent draft
& "./scripts/outlook-draft-update.ps1" -Index 1 -Subject "Updated Subject Line"

# Change recipients on draft #2
& "./scripts/outlook-draft-update.ps1" -Index 2 -To "new@example.com" -CC "cc@example.com"

# Append a note to the body of draft #1
& "./scripts/outlook-draft-update.ps1" -Index 1 -Body "PS: One more thing..." -AppendBody

# Replace body with HTML content
& "./scripts/outlook-draft-update.ps1" -Index 1 -Body "<h1>New Content</h1><p>Updated report.</p>" -HTML

# Add an attachment to draft #3
& "./scripts/outlook-draft-update.ps1" -Index 3 -AddAttachment "C:\docs\report.pdf"

# Add multiple attachments
& "./scripts/outlook-draft-update.ps1" -Index 1 -AddAttachment "C:\a.pdf","C:\b.xlsx"
```

### outlook-send.ps1
Send an email directly. **Requires `-Confirm` flag** as a safety measure — without it, only a preview is shown and nothing is sent. Shows full preview before sending including all recipients, format, and attachments.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address(es) |
| `-Subject` | string | Email subject line |
| `-Body` | string | Email body content (plain text or HTML) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CC` | string | | CC recipient(s) |
| `-BCC` | string | | BCC recipient(s) |
| `-Attachment` | string[] | | File path(s) to attach. Warns if file not found. |
| `-HTML` | switch | off | Treat body as HTML instead of plain text |
| `-Confirm` | switch | off | **Required to actually send.** Without it, only preview is shown. |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**

```powershell
# Preview only (safe - nothing sent)
& "./scripts/outlook-send.ps1" -To "john@example.com" -Subject "Meeting" -Body "Let's meet tomorrow"

# Actually send (requires -Confirm)
& "./scripts/outlook-send.ps1" -To "john@example.com" -Subject "Meeting" -Body "Let's meet tomorrow" -Confirm

# Send with CC, BCC, and attachment
& "./scripts/outlook-send.ps1" -To "john@example.com" -Subject "Report" -Body "See attached" -CC "team@example.com" -BCC "boss@example.com" -Attachment "C:\report.pdf" -Confirm

# Send HTML email with multiple attachments
& "./scripts/outlook-send.ps1" -To "john@example.com" -Subject "Update" -Body "<h1>Status</h1><p>All good.</p>" -HTML -Attachment "C:\a.pdf","C:\b.xlsx" -Confirm
```

### outlook-send-as.ps1
Send an email from a specific Outlook account. Useful when multiple email accounts are configured. Matches account by SMTP address or display name. **Requires `-Confirm` flag** — without it, only a preview is shown. Use `outlook-accounts-list.ps1` to see available accounts.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address(es) |
| `-Subject` | string | Email subject line |
| `-Body` | string | Email body content (plain text or HTML) |
| `-Account` | string | Account SMTP address or display name to send from |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CC` | string | | CC recipient(s) |
| `-BCC` | string | | BCC recipient(s) |
| `-Attachment` | string[] | | File path(s) to attach. Warns if file not found. |
| `-HTML` | switch | off | Treat body as HTML instead of plain text |
| `-Confirm` | switch | off | **Required to actually send.** Without it, only preview is shown. |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**

```powershell
# Preview only (safe - nothing sent)
& "./scripts/outlook-send-as.ps1" -Account "work@company.com" -To "john@example.com" -Subject "Hello" -Body "From my work account"

# Send from specific account
& "./scripts/outlook-send-as.ps1" -Account "work@company.com" -To "john@example.com" -Subject "Hello" -Body "From my work account" -Confirm

# Send from account with attachment
& "./scripts/outlook-send-as.ps1" -Account "sender@example.com" -To "john@example.com" -Subject "Files" -Body "See attached" -Attachment "C:\report.pdf" -Confirm

# List available accounts first
& "./scripts/outlook-accounts-list.ps1"
```

### outlook-reply.ps1
Reply to an email. Creates a reply draft by default. **Requires `-Confirm` flag** to actually send — without it, the reply is saved as a draft. Supports Reply and Reply All. Detects HTML vs plain text format and inserts reply body accordingly (HTML-encoded for HTML emails, plain text for plain text emails). Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (1-based, fallback) |
| `-Body` | string | Reply body content (plain text; HTML-encoded automatically for HTML emails) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | Number of days to look back (only used with -Index) |
| `-ReplyAll` | switch | off | Reply to all recipients instead of just the sender |
| `-Confirm` | switch | off | **Required to actually send.** Without it, reply is saved as draft. |
| `-StripLinks` | bool | `$true` | Replace URLs in reply body preview with `[URL]` |

**Examples:**

```powershell
# Reply to most recent email (saved as draft)
& "./scripts/outlook-reply.ps1" -Index 1 -Body "Thank you for your email."

# Reply All (saved as draft)
& "./scripts/outlook-reply.ps1" -Index 1 -Body "Thanks everyone for the update." -ReplyAll

# Reply and send immediately (requires -Confirm)
& "./scripts/outlook-reply.ps1" -Index 1 -Body "Got it, thanks!" -Confirm

# Reply All and send immediately
& "./scripts/outlook-reply.ps1" -Index 1 -Body "Sounds good to me." -ReplyAll -Confirm

# Reply to an older email (wider date range)
& "./scripts/outlook-reply.ps1" -Index 3 -Body "Following up on this." -Days 14
```

### outlook-forward.ps1
Forward an email. Creates a forward draft by default. **Requires `-Confirm` flag** to actually send — without it, the forward is saved as a draft. Detects HTML vs plain text format and inserts your message accordingly (HTML-encoded for HTML emails, plain text for plain text emails). Supports CC/BCC recipients. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (1-based, fallback) |
| `-To` | string | Recipient email address(es) to forward to |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Body` | string | "" | Message to prepend above forwarded content |
| `-CC` | string | "" | CC recipient(s) |
| `-BCC` | string | "" | BCC recipient(s) |
| `-Days` | int | 7 | Number of days to look back (only used with -Index) |
| `-Confirm` | switch | off | **Required to actually send.** Without it, forward is saved as draft. |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**

```powershell
# Forward most recent email (saved as draft)
& "./scripts/outlook-forward.ps1" -Index 1 -To "colleague@example.com"

# Forward with a message (saved as draft)
& "./scripts/outlook-forward.ps1" -Index 1 -To "colleague@example.com" -Body "FYI - please review"

# Forward with CC and BCC (saved as draft)
& "./scripts/outlook-forward.ps1" -Index 2 -To "colleague@example.com" -CC "team@example.com" -BCC "boss@example.com" -Body "Looping in the team"

# Forward and send immediately (requires -Confirm)
& "./scripts/outlook-forward.ps1" -Index 1 -To "colleague@example.com" -Body "FYI" -Confirm

# Forward an older email (wider date range)
& "./scripts/outlook-forward.ps1" -Index 5 -To "colleague@example.com" -Days 14
```

### outlook-voting.ps1
Send an email with voting buttons (polls). Recipients see clickable voting options in their email client. Responses appear in your Inbox with their votes. **Requires `-Confirm` flag** to actually send — without it, only a preview is shown. Supports HTML body, CC/BCC, and custom voting options.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address(es) |
| `-Subject` | string | Email subject line |
| `-Body` | string | Email body content (plain text or HTML) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Options` | string | "Yes;No" | Semicolon-separated voting options |
| `-CC` | string | "" | CC recipient(s) |
| `-BCC` | string | "" | BCC recipient(s) |
| `-HTML` | switch | off | Treat body as HTML instead of plain text |
| `-Confirm` | switch | off | **Required to actually send.** Without it, only preview is shown. |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**

```powershell
# Preview a voting email (safe - nothing sent)
& "./scripts/outlook-voting.ps1" -To "team@example.com" -Subject "Lunch poll" -Body "Where should we eat?" -Options "Pizza;Tacos;Sushi"

# Send voting email with default Yes/No options
& "./scripts/outlook-voting.ps1" -To "team@example.com" -Subject "Approve?" -Body "Please vote on this proposal." -Confirm

# Send with custom options
& "./scripts/outlook-voting.ps1" -To "team@example.com" -Subject "Meeting time" -Body "Pick a time" -Options "9am;10am;11am;None work" -Confirm

# Send with CC and BCC
& "./scripts/outlook-voting.ps1" -To "team@example.com" -Subject "Project decision" -Body "Which approach?" -Options "Option A;Option B;Option C" -CC "manager@example.com" -BCC "archive@example.com" -Confirm

# Send HTML voting email
& "./scripts/outlook-voting.ps1" -To "team@example.com" -Subject "Survey" -Body "<h2>Quick Poll</h2><p>Please vote below.</p>" -Options "Agree;Disagree;Abstain" -HTML -Confirm
```

## Organize

Move, delete, mark read/unread, view conversation threads, and print emails.

### outlook-move.ps1
Move an email to a specified folder by name or path. Searches Inbox subfolders first, then all accounts recursively. Supports path-based navigation with `/` separator (e.g., `"Projects/2026"` navigates from the account root). Shows available folders if the target is not found. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (1-based, fallback) |
| `-Folder` | string | Target folder name (e.g., `"Archive"`) or path (e.g., `"Projects/2026"`) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | Number of days to look back (only used with -Index) |

**Examples:**

```powershell
# Move most recent email to Archive
& "./scripts/outlook-move.ps1" -Index 1 -Folder "Archive"

# Move email #3 to a subfolder by path (navigates from account root)
& "./scripts/outlook-move.ps1" -Index 3 -Folder "Projects/2026"

# Move an older email (wider date range)
& "./scripts/outlook-move.ps1" -Index 5 -Folder "Archive" -Days 14

# Move to Sent Items (or any top-level folder)
& "./scripts/outlook-move.ps1" -Index 1 -Folder "Sent Items"

# Move to a nested folder by name (found recursively)
& "./scripts/outlook-move.ps1" -Index 2 -Folder "Newsletters"
```

| Script | Purpose | Key Parameters |
|--------|---------|----------------|
| `outlook-delete.ps1` | Delete email (requires -Confirm) | `-EntryID "..." -Confirm -Permanent` |
| `outlook-mark-read.ps1` | Mark read/unread | `-EntryID "..." -Unread` |
| `outlook-conversation.ps1` | View email thread | `-EntryID "..." -Limit 20` |
| `outlook-print.ps1` | Print email (requires -ConfirmPrint) | `-EntryID "..." -ConfirmPrint` |

### outlook-delete.ps1
Delete an email. Without `-Confirm`, shows a preview only. Soft delete moves to Deleted Items. `-Permanent` uses identity-based targeting (StoreID/EntryID where available) and safe-fails instead of deleting an ambiguous item. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |
| `-Permanent` | switch | false | Attempt permanent purge after identity verification; safe-fails if identity cannot be verified |
| `-Confirm` | switch | false | Actually perform the delete (safety gate) |

**Examples:**
```powershell
# Preview which email will be deleted (no action taken)
& "./scripts/outlook-delete.ps1" -Index 1

# Soft delete - move to Deleted Items
& "./scripts/outlook-delete.ps1" -Index 1 -Confirm

# Preview permanent delete (shows warning, no action)
& "./scripts/outlook-delete.ps1" -Index 3 -Permanent

# Permanent delete - attempts purge with identity checks
& "./scripts/outlook-delete.ps1" -Index 3 -Permanent -Confirm

# Delete an older email (search last 30 days)
& "./scripts/outlook-delete.ps1" -Index 5 -Days 30 -Confirm
```

### outlook-mark-read.ps1
Mark an email as read or unread. Detects if the email is already in the desired state and skips the unnecessary save. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |
| `-Unread` | switch | false | Mark as unread instead of read |

**Examples:**
```powershell
# Mark email as read
& "./scripts/outlook-mark-read.ps1" -Index 1

# Mark email as unread
& "./scripts/outlook-mark-read.ps1" -Index 1 -Unread

# Mark an older email as read (search last 30 days)
& "./scripts/outlook-mark-read.ps1" -Index 5 -Days 30
```

### outlook-conversation.ps1
View all emails in the same conversation thread as the specified email. Uses Outlook's native conversation API when available, falls back to subject-matching search. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |
| `-Limit` | int | 20 | Maximum number of thread messages to display |

**Examples:**
```powershell
# View conversation thread for email at index 2
& "./scripts/outlook-conversation.ps1" -Index 2

# Search further back for older threads
& "./scripts/outlook-conversation.ps1" -Index 1 -Days 30

# Limit output to first 5 messages in thread
& "./scripts/outlook-conversation.ps1" -Index 2 -Limit 5
```

### outlook-print.ps1
Print an email using the system's default printer. Without `-ConfirmPrint`, the script runs in preview mode and shows what would be printed. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |
| `-ConfirmPrint` | switch | false | Actually send the email to the default printer |

**Examples:**
```powershell
# Preview the most recent email (no print submitted)
& "./scripts/outlook-print.ps1" -Index 1

# Print an older email (search last 30 days)
& "./scripts/outlook-print.ps1" -Index 3 -Days 30 -ConfirmPrint
```

## Flags, Importance and Sensitivity

Flag emails for follow-up, set priority levels, control sensitivity, and export emails.

| Script | Purpose | Key Parameters |
|--------|---------|----------------|
| `outlook-flag.ps1` | Flag email for follow-up | `-EntryID "..." -Flag "Follow up" -DueDate "2026-02-15"` |
| `outlook-unflag.ps1` | Clear or complete flag | `-EntryID "..." -Complete` |
| `outlook-set-importance.ps1` | Set email priority | `-EntryID "..." -Importance High` |
| `outlook-set-sensitivity.ps1` | Set email sensitivity | `-EntryID "..." -Sensitivity Confidential` |
| `outlook-save-as.ps1` | Export email to file | `-EntryID "..." -Format msg -Path "C:\email.msg"` |

### outlook-flag.ps1
Flag an email for follow-up with optional due date and reminder. Sets the flag text, due date, and Outlook reminder. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Flag` | string | "Follow up" | Flag text (e.g., "Follow up", "Call back", "Review") |
| `-DueDate` | datetime | none | Due date for the flagged task |
| `-ReminderDate` | datetime | none | When to show an Outlook reminder |
| `-Days` | int | 7 | How many days back to search (only used with -Index) |

**Examples:**
```powershell
# Basic flag with default "Follow up" text
& "./scripts/outlook-flag.ps1" -Index 1

# Custom flag text
& "./scripts/outlook-flag.ps1" -Index 2 -Flag "Call back"

# Flag with due date
& "./scripts/outlook-flag.ps1" -Index 1 -DueDate "2026-02-15"

# Flag with due date and reminder
& "./scripts/outlook-flag.ps1" -Index 1 -DueDate "2026-02-15" -ReminderDate "2026-02-14 09:00"

# Flag an older email
& "./scripts/outlook-flag.ps1" -Index 3 -Days 30
```

### outlook-unflag.ps1
Remove a flag or mark it as complete. Clears the flag text, status, and any reminder. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |
| `-Complete` | switch | false | Mark flag as complete instead of removing it |

**Examples:**
```powershell
# Remove flag from email
& "./scripts/outlook-unflag.ps1" -Index 1

# Mark flag as complete (shows checkmark in Outlook)
& "./scripts/outlook-unflag.ps1" -Index 1 -Complete

# Unflag an older email
& "./scripts/outlook-unflag.ps1" -Index 3 -Days 30
```

### outlook-set-importance.ps1
Set the importance/priority level of an email (Low, Normal, High). Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided, plus `-Importance`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |
| `-Importance` | string | Priority level: `Low`, `Normal`, or `High` |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |

**Examples:**
```powershell
# Set email as high importance
& "./scripts/outlook-set-importance.ps1" -Index 1 -Importance High

# Set email as low importance
& "./scripts/outlook-set-importance.ps1" -Index 2 -Importance Low

# Reset to normal importance
& "./scripts/outlook-set-importance.ps1" -Index 1 -Importance Normal
```

### outlook-set-sensitivity.ps1
Set the sensitivity level of an email (Normal, Personal, Private, Confidential). Note: sensitivity can typically only be changed on drafts/new messages; received emails may be read-only depending on Exchange settings. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided, plus `-Sensitivity`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |
| `-Sensitivity` | string | Sensitivity level: `Normal`, `Personal`, `Private`, or `Confidential` |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | 7 | How many days back to search (only used with -Index) |

**Examples:**
```powershell
# Set email as confidential
& "./scripts/outlook-set-sensitivity.ps1" -Index 1 -Sensitivity Confidential

# Set email as private
& "./scripts/outlook-set-sensitivity.ps1" -Index 2 -Sensitivity Private

# Reset to normal sensitivity
& "./scripts/outlook-set-sensitivity.ps1" -Index 1 -Sensitivity Normal
```

### outlook-save-as.ps1
Export an email to a file in various formats (`msg`, `txt`, `html`, `mht`). Use `-EntryID` (preferred) or `-Index` (fallback). If no file path is provided, the script saves to Downloads and auto-generates a safe filename from subject (timestamp fallback when subject is empty).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from outlook-read.ps1 listing (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Format` | string | "msg" | Output format: `msg`, `txt`, `html`, `mht` |
| `-Path` | string | Downloads | Full path for the output file or output directory |
| `-Days` | int | 7 | How many days back to search (only used with -Index) |

**Examples:**
```powershell
# Save as Outlook message file (default)
& "./scripts/outlook-save-as.ps1" -Index 1

# Save as HTML file
& "./scripts/outlook-save-as.ps1" -Index 1 -Format html -Path "C:\Emails\email.html"

# Save as plain text
& "./scripts/outlook-save-as.ps1" -Index 2 -Format txt -Path "C:\Emails\email.txt"
```

## Categories

Manage Outlook categories: list, create, assign to emails, remove from emails, and delete.

### outlook-list-categories.ps1
List all Outlook categories with their color names and color index numbers.

**Required Parameters:** None

**Examples:**
```powershell
# List all categories
& "./scripts/outlook-list-categories.ps1"
```

### outlook-create-category.ps1
Create a new Outlook category with a specified name and color.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Name` | string | Name of the new category |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Color` | int | `1` (Red) | Color index (0-25). See color reference table below |

**Color Reference:** 0=None, 1=Red, 2=Orange, 3=Peach, 4=Yellow, 5=Green, 6=Teal, 7=Olive, 8=Blue, 9=Purple, 10=Maroon, 11=Steel, 12=DarkSteel, 13=Gray, 14=DarkGray, 15=Black, 16=DarkRed, 17=DarkOrange, 18=DarkPeach, 19=DarkYellow, 20=DarkGreen, 21=DarkTeal, 22=DarkOlive, 23=DarkBlue, 24=DarkPurple, 25=DarkMaroon

**Examples:**
```powershell
# Create a green category
& "./scripts/outlook-create-category.ps1" -Name "Project Alpha" -Color 5

# Create a blue category (default color is Red if -Color omitted)
& "./scripts/outlook-create-category.ps1" -Name "Meetings" -Color 8

# Create with default red color
& "./scripts/outlook-create-category.ps1" -Name "Urgent"
```

### outlook-assign-category.ps1
Assign a category to an email. Appends to existing categories (does not replace). Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided, plus `-Category`.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (fallback) |
| `-Category` | string | Category name (must already exist) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Days` | int | `7` | How far back to look (only used with -Index) |

**Examples:**
```powershell
# Assign "Urgent" to email #3
& "./scripts/outlook-assign-category.ps1" -Index 3 -Category "Urgent"

# Assign to older email
& "./scripts/outlook-assign-category.ps1" -Index 5 -Category "Project Alpha" -Days 30
```

### outlook-remove-category.ps1
Remove a specific category or all categories from an email. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (fallback) |

**Optional Parameters (one required):**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Category` | string | `""` | Specific category name to remove |
| `-All` | switch | `$false` | Remove all categories from the email |
| `-Days` | int | `7` | How far back to look (only used with -Index) |

**Examples:**
```powershell
# Remove specific category
& "./scripts/outlook-remove-category.ps1" -Index 1 -Category "Urgent"

# Remove all categories from an email
& "./scripts/outlook-remove-category.ps1" -Index 1 -All

# Remove from older email
& "./scripts/outlook-remove-category.ps1" -Index 5 -Category "Old Project" -Days 30
```

### outlook-delete-category.ps1
Delete a category from the master category list (does not remove it from emails that already have it).

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Name` | string | Exact name of the category to delete |

**Examples:**
```powershell
# Delete a category
& "./scripts/outlook-delete-category.ps1" -Name "Old Project"
```

## Scheduling and Receipts

Schedule delayed sends, request read/delivery receipts, and recall sent emails.

### outlook-schedule-send.ps1
Schedule an email to be sent at a future date/time. Email goes to Outbox and sends automatically at the specified time. Requires `-Confirm` to schedule.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address |
| `-Subject` | string | Email subject |
| `-Body` | string | Email body text |
| `-SendAt` | datetime | Future date/time to send (e.g., `"2026-02-10 09:00"`) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CC` | string | `""` | CC recipients |
| `-BCC` | string | `""` | BCC recipients |
| `-Attachment` | string[] | `@()` | File path(s) to attach |
| `-HTML` | switch | `$false` | Send as HTML format |
| `-ReadReceipt` | switch | `$false` | Request read receipt |
| `-DeliveryReceipt` | switch | `$false` | Request delivery receipt |
| `-Confirm` | switch | `$false` | Required to actually schedule the email |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**
```powershell
# Preview scheduled email (no -Confirm = dry run)
& "./scripts/outlook-schedule-send.ps1" -To "john@example.com" -Subject "Reminder" -Body "Don't forget the meeting" -SendAt "2026-02-10 09:00"

# Actually schedule the email
& "./scripts/outlook-schedule-send.ps1" -To "john@example.com" -Subject "Reminder" -Body "Don't forget" -SendAt "2026-02-10 09:00" -Confirm

# Schedule with read receipt
& "./scripts/outlook-schedule-send.ps1" -To "john@example.com" -Subject "Report" -Body "See attached" -SendAt "2026-02-10 09:00" -ReadReceipt -Confirm
```

### outlook-request-receipt.ps1
Send an email with read and/or delivery receipt requests. Requires `-Confirm` to send and at least one receipt type.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-To` | string | Recipient email address |
| `-Subject` | string | Email subject |
| `-Body` | string | Email body text |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-CC` | string | `""` | CC recipients |
| `-BCC` | string | `""` | BCC recipients |
| `-Attachment` | string[] | `@()` | File path(s) to attach |
| `-HTML` | switch | `$false` | Send as HTML format |
| `-ReadReceipt` | switch | `$false` | Request notification when email is read |
| `-DeliveryReceipt` | switch | `$false` | Request notification when email is delivered |
| `-Confirm` | switch | `$false` | Required to actually send the email |
| `-StripLinks` | bool | `$true` | Replace URLs in body preview with `[URL]` |

**Examples:**
```powershell
# Preview with read receipt (no -Confirm = dry run)
& "./scripts/outlook-request-receipt.ps1" -To "john@example.com" -Subject "Important" -Body "Please confirm" -ReadReceipt

# Send with read receipt
& "./scripts/outlook-request-receipt.ps1" -To "john@example.com" -Subject "Important" -Body "Please confirm" -ReadReceipt -Confirm

# Send with both receipt types
& "./scripts/outlook-request-receipt.ps1" -To "john@example.com" -Subject "Contract" -Body "Review attached" -ReadReceipt -DeliveryReceipt -Confirm
```

### outlook-recall.ps1
Recall helper for sent email (Exchange/Office 365 only). True recipient recall must be done in Outlook UI; this script cannot submit a server-side recall request via COM. `-DeleteUnread` is Sent Items cleanup only (not recipient recall). Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from Sent Items (most recent first, fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-DeleteUnread` | switch | `$false` | Delete the email from your Sent Items only (not recipient recall) |
| `-ConfirmDeleteFromSent` | switch | `$false` | Required guard to perform `-DeleteUnread` action |
| `-Days` | int | `7` | How far back to look in Sent Items (only used with -Index) |

**Examples:**
```powershell
# View recall options for most recent sent email
& "./scripts/outlook-recall.ps1" -Index 1

# Delete sent email from your Sent Items (explicit confirmation required)
& "./scripts/outlook-recall.ps1" -Index 1 -DeleteUnread -ConfirmDeleteFromSent

# Look further back in sent items
& "./scripts/outlook-recall.ps1" -Index 3 -Days 30
```

## Account and Rule Management

List folders and accounts, manage email rules, configure out-of-office, and save attachments.

### outlook-folders.ps1
List all mail folders across all accounts with item counts and unread counts, including subfolders.

**Required Parameters:** None

**Examples:**
```powershell
# List all folders
& "./scripts/outlook-folders.ps1"
```

### outlook-accounts-list.ps1
List all email accounts configured in Outlook. Shows each account's display name, SMTP email address, and type (Exchange, IMAP, POP3, EAS, HTTP). Also shows the first account in the session collection (which may differ from the effective default send account). Falls back to listing mail stores if the Accounts collection is empty.

**Required Parameters:** None

**Output per account:**
- Display name (numbered)
- SMTP email address
- Account type (human-readable: Exchange, IMAP, POP3, EAS, HTTP)
- First account in session collection highlighted at the end (informational)

**Examples:**
```powershell
# List all configured accounts
& "./scripts/outlook-accounts-list.ps1"

# Use with outlook-send-as.ps1 to send from a specific account
& "./scripts/outlook-accounts-list.ps1"   # find the account name/email first
& "./scripts/outlook-send-as.ps1" -To "user@example.com" -Subject "Test" -Body "Hello" -Account "myother@email.com" -Confirm
```

### outlook-rules-list.ps1
List all email rules with [ON]/[OFF] status indicators. Use `-Detailed` to see conditions (From, Subject, SentTo, HasAttachment, Importance) and actions (MoveToFolder, CopyToFolder, Delete, Forward, AssignCategory, MarkRead, Stop) for each rule. Summary shows displayed/enabled/total counts.

**Required Parameters:** None

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Enabled` | switch | `$false` | Only show enabled rules |
| `-Detailed` | switch | `$false` | Show conditions, actions, and execution order for each rule |
| `-StripLinks` | bool | `$true` | Replace URLs in rule conditions/actions with `[URL]` |

**Examples:**
```powershell
# List all rules with ON/OFF status
& "./scripts/outlook-rules-list.ps1"

# Only enabled rules
& "./scripts/outlook-rules-list.ps1" -Enabled

# Full details: conditions, actions, execution order
& "./scripts/outlook-rules-list.ps1" -Detailed

# Enabled rules with full details
& "./scripts/outlook-rules-list.ps1" -Enabled -Detailed
```

### outlook-rules-create.ps1
Create a new Outlook rule for incoming email with flexible conditions and actions. Validates inputs before connecting to Outlook. Shows a preview before saving. Duplicate rule names are warned but allowed with `-Force`.

**Required Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Name` | string | Name for the rule |

**Optional Parameters - Conditions (at least one required):**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-FromAddress` | string | `""` | Match emails from this sender address |
| `-SubjectContains` | string | `""` | Match emails with this text in subject |
| `-BodyContains` | string | `""` | Match emails with this text in body |
| `-SentTo` | string | `""` | Match emails sent to this recipient |
| `-HasAttachment` | switch | `$false` | Match emails that have attachments |

**Optional Parameters - Actions (at least one required):**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-MoveToFolder` | string | `""` | Move matching emails to this folder |
| `-CopyToFolder` | string | `""` | Copy matching emails to this folder |
| `-ForwardTo` | string | `""` | Forward matching emails to this address |
| `-RedirectTo` | string | `""` | Redirect matching emails to this address |
| `-AssignCategory` | string | `""` | Assign this category to matching emails |
| `-Delete` | switch | `$false` | Move matching emails to Deleted Items |
| `-DeletePermanently` | switch | `$false` | Permanently delete matching emails |
| `-StopProcessing` | switch | `$false` | Stop processing further rules |
| `-DesktopAlert` | switch | `$false` | Show desktop notification |

**Optional Parameters - Other:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Disabled` | switch | `$false` | Create the rule in disabled state |
| `-Force` | switch | `$false` | Allow creation even if rule name already exists |
| `-StripLinks` | bool | `$true` | Replace URLs in rule preview output with `[URL]` |

**Examples:**
```powershell
# Move newsletters to Archive folder
& "./scripts/outlook-rules-create.ps1" -Name "Archive newsletters" -FromAddress "newsletter@example.com" -MoveToFolder "Archive"

# Forward reports to your boss
& "./scripts/outlook-rules-create.ps1" -Name "Forward reports" -SubjectContains "weekly report" -ForwardTo "boss@example.com"

# Flag emails with attachments
& "./scripts/outlook-rules-create.ps1" -Name "Flag attachments" -HasAttachment -AssignCategory "Has Files" -DesktopAlert

# Delete spam and stop processing
& "./scripts/outlook-rules-create.ps1" -Name "Spam filter" -FromAddress "spam@example.com" -Delete -StopProcessing

# Create rule for a distribution list
& "./scripts/outlook-rules-create.ps1" -Name "Team emails" -SentTo "team@example.com" -MoveToFolder "Team"

# Create disabled rule for later activation
& "./scripts/outlook-rules-create.ps1" -Name "Archive old" -SubjectContains "archive" -MoveToFolder "Archive" -Disabled

# Allow duplicate rule name
& "./scripts/outlook-rules-create.ps1" -Name "Archive newsletters" -BodyContains "unsubscribe" -AssignCategory "Newsletter" -Force
```

**Note:** The `MarkAsRead` action is not supported by Outlook's COM rule API for programmatic rule creation (per Microsoft docs). Use the Rules and Alerts Wizard in Outlook for that action.

### outlook-out-of-office.ps1
Configure Out of Office / automatic replies (Exchange / Microsoft 365). Detects the Exchange account via Outlook COM. If ExchangeOnlineManagement module is connected, runs commands directly; otherwise generates ready-to-run PowerShell commands.

**Required Parameters:** None (but one of `-Enable`, `-Disable`, or `-Status` must be specified)

**Action Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Enable` | switch | Enable out-of-office auto-reply |
| `-Disable` | switch | Disable out-of-office auto-reply |
| `-Status` | switch | Check current out-of-office status |

**Message Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-InternalMessage` | string | `""` | Auto-reply for internal senders (required with `-Enable`) |
| `-ExternalMessage` | string | `""` | Auto-reply for external senders |
| `-ExternalAudienceAll` | switch | `$false` | Send external reply to all senders (default: known contacts only) |

**Schedule Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `-StartTime` | datetime | Schedule start (requires `-EndTime`; sets state to Scheduled) |
| `-EndTime` | datetime | Schedule end (requires `-StartTime`; must be later than `-StartTime`) |

**Other Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-StripLinks` | bool | `$true` | Replace URLs in auto-reply message output with `[URL]` |

**Examples:**
```powershell
# Check current OOF status
& "./scripts/outlook-out-of-office.ps1" -Status

# Enable with internal message
& "./scripts/outlook-out-of-office.ps1" -Enable -InternalMessage "I'm out of office until Monday."

# Enable with scheduled dates
& "./scripts/outlook-out-of-office.ps1" -Enable -InternalMessage "Away until March 15" -StartTime "2026-03-10" -EndTime "2026-03-15"

# Enable with both internal and external messages
& "./scripts/outlook-out-of-office.ps1" -Enable -InternalMessage "Away" -ExternalMessage "Out of office, back Monday." -ExternalAudienceAll

# Disable out of office
& "./scripts/outlook-out-of-office.ps1" -Disable
```

**Note:** Out of Office is a server-side Exchange feature. The Outlook COM API cannot access these settings directly. If ExchangeOnlineManagement module is connected (`Connect-ExchangeOnline`), the script runs commands directly. Otherwise it generates ready-to-run `Set-MailboxAutoReplyConfiguration` / `Get-MailboxAutoReplyConfiguration` commands.

### outlook-attachment-save.ps1
Save email attachments to disk. Lists attachments with sizes, detects inline/embedded images (signature logos, etc.) and skips them by default. Handles duplicate filenames automatically. Use `-ListOnly` to preview without saving. Use `-EntryID` (preferred) or `-Index` (fallback).

**Required Parameters:** One of `-EntryID` or `-Index` must be provided.

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EntryID` | string | Stable email ID from listing output (preferred) |
| `-Index` | int | Email number from `outlook-read.ps1` output (fallback) |

**Optional Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Path` | string | Downloads folder | Directory to save attachments to (created if missing) |
| `-AttachmentIndex` | int | `0` (all) | Save only this attachment (1-based) |
| `-Days` | int | `7` | How far back to look (only used with -Index) |
| `-ListOnly` | switch | `$false` | List attachments without saving |
| `-IncludeInline` | switch | `$false` | Also save inline/embedded images (skipped by default) |

**Examples:**
```powershell
# List attachments without saving
& "./scripts/outlook-attachment-save.ps1" -Index 1 -ListOnly

# Save all regular attachments (inline images skipped)
& "./scripts/outlook-attachment-save.ps1" -Index 1

# Save to specific directory
& "./scripts/outlook-attachment-save.ps1" -Index 1 -Path "C:\Downloads\project"

# Save only the second attachment
& "./scripts/outlook-attachment-save.ps1" -Index 3 -AttachmentIndex 2

# Save everything including inline images
& "./scripts/outlook-attachment-save.ps1" -Index 1 -IncludeInline

# Save with wider date range
& "./scripts/outlook-attachment-save.ps1" -Index 5 -Days 30
```

**Note:** Inline/embedded attachments (e.g., signature images) are detected via the `PR_ATTACH_CONTENT_ID` MAPI property and marked with `[inline]` in the listing. They are skipped by default when saving all attachments. Use `-AttachmentIndex` to save a specific one, or `-IncludeInline` to save all.
