---
name: outlook-contacts
description: "Find people and manage classic Outlook contacts on Windows through PowerShell COM: find anyone by name, partial name or email across your Contacts, your organization's directory (Global Address List) and recent mail, with contact points (emails, phones, title, department, office, manager); a 'who is this' profile with mail exchanged and meetings together; members of distribution lists and contact groups; org chart (manager chain, direct reports, peers); directory search by department or title; duplicate contact report and merge; vCard/CSV import; list, search, create, update, delete, photo and vCard export of contacts. Fast (one table query per folder, a local directory name index), agent-friendly (EntryIDs, -Json, exit codes), with a confirm switch on delete, dedupe -Merge and import."
---

# Outlook Contacts and People

10 PowerShell scripts in `scripts/` that drive classic Outlook (running on this Windows PC) through COM. Every script works standalone (`pwsh -File scripts/<name>.ps1 ...`) and through the mail skill's warm host (`../outlook-mail/scripts/outlook-call.ps1 people "jane"`), under PowerShell 7 and Windows PowerShell 5.1.

## How to work

1. **Looking for a person? Use `outlook-people.ps1`.** It searches your Contacts in every mailbox, your organization's directory and the people in your Inbox and Sent Items, and ranks the matches. `outlook-contacts-*` scripts work only on your own Contacts folder.
2. **Find first, then act by EntryID.** List and search print an `EntryID` for every contact; pass it to update, photo, export or delete with `-EntryID`. `-Index` still works (the `[n]` position in `outlook-contacts-list.ps1` with the same `-Search`), but positions shift as contacts change.
3. **Output contract:** success exits 0; a failure prints `Error: <reason>` (often a `Hint:` line) and exits 1. Every script takes `-Json` for compact JSON.
4. **Delete, dedupe `-Merge` and import are gated.** Delete and dedupe `-Merge` use the confirmation key "delete", import uses "import". Without `-Confirm` they print a preview and `PREVIEW ONLY: ... was NOT done.` and exits 0 (see the mail skill's Confirmation section for turning this off). As in mail, `-Confirm` is a plain switch: `-Confirm:$false` only previews, and the old `-WhatIf` and interactive "Are you sure?" prompt are gone.

## Which script when

| Need | Script | Key parameters |
|---|---|---|
| Find anyone (name, first or last name, partial, email) with contact points | `outlook-people.ps1` | `"<query>" -Limit 10 -Sources Contacts,Directory,Recent -Days 180 [-Detail] -Json` |
| Who is this person: contact points, mail exchanged, meetings together | `outlook-people.ps1` | `"<query>" -Profile [-ProfileDays 365 -UpcomingDays 60] -Json` |
| Members of a distribution list or contact group | `outlook-people.ps1` | `-Members "<name or address>" [-Limit 100] -Json` |
| Org chart: manager chain, direct reports, peers | `outlook-people.ps1` | `-Org "<person>" [-Levels 4] [-Limit 25] -Json` |
| Directory by department and/or title ("who in IT", "deans") | `outlook-people.ps1` | `-Department "<text>" -Title "<text>" ["<name>"] [-Limit 25] -Json` |
| Rebuild the local directory name index now | `outlook-people.ps1` | `-RefreshDirectory` |
| List your contacts | `outlook-contacts-list.ps1` | `-Limit 20 -Search "text" [-AllStores] -Json` |
| Search your contacts by field (AND) | `outlook-contacts-search.ps1` | `-Name -Email -Company -Limit 10 [-AllStores] -Json` |
| Create a contact | `outlook-contacts-create.ps1` | `-FullName (required) -Email -Email2 -Phone -Mobile -HomePhone -Company -JobTitle -Department -Office -Categories -Notes -Json` |
| Change a contact | `outlook-contacts-update.ps1` | `-EntryID` (or `-Index [-Search]`) plus any of `-FullName -Email -Email2 -Phone -Mobile -HomePhone -Company -JobTitle -Department -Office -Categories -Notes`; pass `""` to clear a field |
| Add or remove a photo | `outlook-contacts-photo.ps1` | `-EntryID -PhotoPath <jpg/png/gif/bmp>` or `-Remove` |
| Export as vCard (.vcf) | `outlook-contacts-export.ps1` | `-EntryID [-Path file-or-folder]` (default Downloads; never overwrites) |
| Duplicate contacts: report, or merge same-email groups | `outlook-contacts-dedupe.ps1` | `[-Search "text"] [-AllStores] [-MaxGroups 10] [-Merge -Confirm -Plan <hash>] -Json` |
| Import contacts from .vcf or .csv | `outlook-contacts-import.ps1` | `-Path <file> [-Map "Header=Field; ..."] [-Update [-Overwrite]] [-Confirm] [-Show 20] [-MaxRows 2000] [-MaxFileMB 20] -Json` |
| Delete a contact (to Deleted Items) | `outlook-contacts-delete.ps1` | `-EntryID` or `-Name "exact full name"` (must be unique), `-Confirm` |

## Finding people

`outlook-people.ps1 "<query>"` returns the best matches first. Each person shows the sources that found them (`Contacts`, `Directory`, `Recent`), a score, title, department, company, every email and phone, office, and how much recent mail you exchanged. `-Detail` adds the manager (one extra directory lookup per person).

- **Contacts:** the default Contacts folder of every mailbox in the profile.
- **Directory:** the Exchange Global Address List (your organization's directory; about 17,000 entries in the one this was built on). Names are matched against a local name index, `%LOCALAPPDATA%\outlook-suite\directory-names.txt`, which is built on first use (about 12 to 15 s), refreshed automatically after 7 days, and kept in memory by the warm host. The index is saved only when the walk read at least 95% of the directory. Each matching name then opens one directory entry by name; when several entries share a display name, only the first can be opened this way, so the result carries `Namesakes` (how many share it) and the name alone never resolves to an address. If a name is missing from the index (a new hire, or an alias), Exchange resolves it directly. An outside address (one-off SMTP) is never reported as a directory hit.
- **Recent:** senders in every received folder (the Inbox and the folders your rules file mail into; not Sent, Drafts, Deleted, Junk or Outbox) and recipients in Sent Items, in the last `-Days` (default 180). One indexed table query per folder; Sent Items recipients come from the To/CC lines. This finds people outside your organization who are not in your Contacts.

Matching: an exact name or email scores 100; every query word starting a name word (`"jane"`, `"doe"`, `"ja do"`) scores 80; a substring scores 60. People found by several sources and people you correspond with often rank higher. A query that is a full email address is resolved directly.

`-Profile` picks the best match (Contacts and directory first, recent mail only if they find nobody) and shows contact points with the manager, mail exchanged (every received folder from them, Sent Items to them: counts, last dates and the last five subjects, never bodies) and meetings together (count in the past `-ProfileDays`, count in the next `-UpcomingDays`, last and next meeting; canceled meetings skipped). When other people match nearly as well, they are listed so you can rerun with the exact email. The counts are approximate: Sent Items and meetings are matched by display name, so a recipient or attendee shown only as a raw address or as "Last, First" is missed. About 1 s through the warm host.

`-Members "<name or address>"` lists the members of a local contact group (any Contacts folder, exact name) or an Exchange distribution list (by address or directory name): name, email and kind (`Person` or `Group` for nested lists), read-only. The total is always shown; only the first `-Limit` (default 100) members are read, because each member is a separate directory lookup (about 25 ms each). Large lists also take a few seconds to open.

```powershell
& "./scripts/outlook-people.ps1" "doe"                    # last name
& "./scripts/outlook-people.ps1" "jane d" -Detail -Json   # partial first and last name, with manager
& "./scripts/outlook-people.ps1" "jane@example.org" -Profile
& "./scripts/outlook-people.ps1" -Members "team-list@example.org" -Limit 50
```

## Org chart and directory attributes

`-Org "<person>"` finds the person in the directory (name, partial name or email; other close matches are listed) and shows their manager chain up to `-Levels` (default 4, closest first), their direct reports (first `-Limit`, default 25) and their peers, meaning the manager's other direct reports (first 10 unless `-Limit` is given). Totals are always shown. Each manager level is one directory server lookup (about 0.5 s), so a 4-level chain takes 2 to 4 s. Read-only.

`-Department "<text>"` and/or `-Title "<text>"` search the directory by attribute (contains, case-insensitive; with both, both must match; a name query narrows further), for example `-Department "information tech"` or `-Title "dean"`. They use a second local index, `%LOCALAPPDATA%\outlook-suite\directory-attributes.tsv` (name, department and title per entry; about 450 KB for 17,000 entries), built on the first attribute search (about 20 to 35 s, once), refreshed after 7 days, and saved only when complete. After that a search takes about 0.5 s. Email, phone and office are looked up for the people shown; a display name shared by several directory entries gets no address. Many directory entries (students, for example) have no department or title, so they never match an attribute search.

## Contact hygiene and import

`outlook-contacts-dedupe.ps1` reports likely duplicates in your Contacts. Exchange addresses are compared as their resolved SMTP address.

- **Same email, mergeable:** contacts that share an email address, are in one mailbox, and where every duplicate shares at least one name word with the keeper.
- **Same email, report only:** the same address but no shared name word (a shared or role address, for example), or contacts in different mailboxes. Each group shows its reason.
- **Same name only:** the same name (case, punctuation and word order ignored) with no shared email. Report only: two people can share a name.

The keeper of a group is its most recently modified contact. The report lists every contact that would be deleted (name and email), every conflicting field (the kept value and the dropped one), notes that will be appended, and what cannot be carried over (extra emails, postal addresses, birthday, anniversary, photo, web page and similar). It ends with a plan hash covering the keepers, the contacts to delete and the filters; at most `-MaxGroups` (default 10) groups are planned per run.

`-Merge -Confirm -Plan <hash>` merges exactly that plan: it fills the keeper's empty fields and free email slots (never an unresolved Exchange address), appends each duplicate's notes under a divider, combines the categories, and deletes the duplicates to Deleted Items. It refuses without `-Plan`, or when the contacts changed since the report (run the report again). Each group is merged on its own: a failed group does not stop the others, the result lists merged, failed and untouched groups, and the script exits 1 when any group failed. Without `-Confirm` (or with `-Confirm:$false`) `-Merge` only previews.

`outlook-contacts-import.ps1 -Path <file>` reads a vCard file (one or many cards: name, emails, phones by type including `tel:` URIs, organization and department with escaped `\;`, title, quoted-printable notes, categories; a card without `END:VCARD` counts as invalid) or a CSV file (read as UTF-8; a byte-order mark is fine). CSV headers are matched by common names (Full Name or Name, First Name and Last Name, E-mail Address or Email, Business Phone, Mobile Phone, Home Phone, Company, Job Title, Department, Office, Categories, Notes); `-Map "Work Email=Email; Cell=Mobile"` adds or overrides header mappings (fields: FullName, FirstName, LastName, Email, Email2, Email3, Phone, Mobile, HomePhone, Company, JobTitle, Department, Office, Categories, Notes). Control characters are removed from single-line fields, and notes are capped at 32 KB.

Every email of a row is matched against your contacts. One match: the row is skipped, or with `-Update` that contact is updated. Several matches: the row is reported as ambiguous and never written. A row without an email is skipped when a contact with the same name and company exists. `-Update` fills only empty fields by default; add `-Overwrite` to replace differing values too. Notes are always appended under a divider, never replaced. The preview shows counts (create, update, unchanged, skipped, ambiguous, duplicate in file, invalid) and the first `-Show` rows with each update's field changes; contacts are written only with `-Confirm`. Files over `-MaxRows` (default 2000) rows or `-MaxFileMB` (default 20) are refused. The script exits 1 when any row fails.
## For other scripts: `_people.ps1`

Any skill can turn a name into a person. Dot-source after `_common.ps1`:

```powershell
. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "..\..\outlook-contacts\scripts\_people.ps1")
$smtp = Resolve-PersonSmtp -Namespace $Namespace -Query "jane doe"    # one address, or throws listing the candidates
$list = Resolve-Person -Namespace $Namespace -Query "doe" -Limit 5    # ranked records (Name, Email, Score, Sources, Title, ...)
```

`Resolve-PersonSmtp` returns an email address unchanged. For a name it returns an address only when `Select-ClearPerson` finds one clear person, judged on `Match` (the name or email match quality before any ranking bonus): the best `Match` is at least 80 (every query word starts a word of the name), exactly one person has it, that person has no namesakes in the directory, and the match is exact (100) or leads the next person by 20 or more. Otherwise it throws `'<query>' matches more than one person: Name <email>; ...` (or "matches no one closely enough", or "No person found") so the caller can ask which one. It never guesses, and how often you correspond with someone (which only orders the list) never decides it.

## Examples

```powershell
& "./scripts/outlook-contacts-list.ps1" -Search "acme" -Json
& "./scripts/outlook-contacts-search.ps1" -Name "john" -Company "acme"
& "./scripts/outlook-contacts-create.ps1" -FullName "John Doe" -Email "john@example.com" -Mobile "555-1234" -Company "Acme Inc"
& "./scripts/outlook-contacts-update.ps1" -EntryID "00000000..." -JobTitle "Director" -Phone ""
& "./scripts/outlook-contacts-photo.ps1" -EntryID "00000000..." -PhotoPath "C:\Photos\john.jpg"
& "./scripts/outlook-contacts-export.ps1" -EntryID "00000000..." -Path "C:\Contacts\"
& "./scripts/outlook-contacts-delete.ps1" -EntryID "00000000..." -Confirm
& "./scripts/outlook-people.ps1" -Org "jane doe"
& "./scripts/outlook-people.ps1" -Title "director" -Department "information"
& "./scripts/outlook-contacts-dedupe.ps1" -Json
& "./scripts/outlook-contacts-import.ps1" -Path "C:\in\people.csv" -Map "Cell=Mobile" -Confirm
```

## Notes

- Contact lists are read with one `Folder.GetTable()` query per Contacts folder (about 0.5 s for a call), and EntryIDs are long-term IDs that stay valid across processes.
- Contacts with Exchange (EX) addresses are shown with their SMTP address. People search resolves those only for contacts that match, and a mailbox that cannot be read is skipped instead of failing the search.
- Test: `pwsh -File tests/Test-PeopleContacts.ps1 [-Shell powershell]` creates one `[suite-test]` contact and contact group, runs every script against them (including the clear-person rule, directory namesakes, outside addresses, received folders, `-Members`, `-Org`, attribute search, dedupe and import), uses only undeliverable `@suite-test.invalid` addresses, forces the confirm gate on, prints PASS/FAIL and timings only, and removes and purges them.
