---
name: outlook-contacts
description: "Manage Outlook contacts via PowerShell: list, create, search, update, delete contacts, add/remove photos, and export as vCard. Requires Outlook running."
---

# Outlook Contacts

Manage Microsoft Outlook contacts using PowerShell COM objects.

## Prerequisites

- Microsoft Outlook installed and running
- PowerShell 5.1+ (included in Windows 10/11)

## List Contacts

### outlook-contacts-list.ps1

Browse contacts with optional search filtering. Returns EntryIDs for use with action scripts.

**Required Parameters:** None (lists all contacts by default)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Limit` | int | 20 | Maximum contacts to display |
| `-Search` | string | — | Filter by name, email, or company (case-insensitive) |

```powershell
# List all contacts (first 20)
& "./scripts/outlook-contacts-list.ps1" -Limit 20

# List with search filter
& "./scripts/outlook-contacts-list.ps1" -Search "John"
```

## Search Contacts

### outlook-contacts-search.ps1

Search contacts by name, email, and/or company. Multiple criteria use AND logic. Returns EntryIDs for use with action scripts.

**Required Parameters:** At least one of `-Name`, `-Email`, or `-Company`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Name` | string | — | Search by contact name (case-insensitive contains) |
| `-Email` | string | — | Search by email address (case-insensitive contains) |
| `-Company` | string | — | Search by company name (case-insensitive contains) |
| `-Limit` | int | 10 | Maximum results to return |

```powershell
# Search by name
& "./scripts/outlook-contacts-search.ps1" -Name "John"

# Search by company
& "./scripts/outlook-contacts-search.ps1" -Company "Acme"

# Search by email domain
& "./scripts/outlook-contacts-search.ps1" -Email "example.com"

# Combined search (AND logic)
& "./scripts/outlook-contacts-search.ps1" -Name "John" -Company "Acme"
```

## Create Contact

### outlook-contacts-create.ps1

Create a new contact in the default Contacts folder.

**Required Parameters:** `-FullName`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-FullName` | string | **Required** | Contact's full name |
| `-Email` | string | — | Email address |
| `-Phone` | string | — | Business phone number |
| `-Mobile` | string | — | Mobile phone number |
| `-HomePhone` | string | — | Home phone number |
| `-Company` | string | — | Company name |
| `-JobTitle` | string | — | Job title |
| `-Notes` | string | — | Notes/body text |

```powershell
# Create a new contact
& "./scripts/outlook-contacts-create.ps1" -FullName "John Doe" -Email "john@example.com" -Phone "555-1234" -Company "Acme Inc" -JobTitle "Manager"

# Create with mobile and notes
& "./scripts/outlook-contacts-create.ps1" -FullName "Jane Smith" -Email "jane@example.com" -Mobile "555-5678" -Notes "Met at conference"
```

## Update Contact

### outlook-contacts-update.ps1

Edit an existing contact's fields. Uses EntryID (preferred) or Index as targeting method.

**Required Parameters:** `-EntryID` or `-Index`, plus at least one field to update

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Contact's unique EntryID (preferred, from list/search) |
| `-Index` | int | 0 | Contact position number (fallback) |
| `-FullName` | string | — | New full name |
| `-Email` | string | — | New email address |
| `-Phone` | string | — | New business phone |
| `-Mobile` | string | — | New mobile phone |
| `-Company` | string | — | New company name |
| `-JobTitle` | string | — | New job title |
| `-Notes` | string | — | New notes/body text |
| `-Search` | string | — | Filter contacts before applying Index (narrows list) |

```powershell
# Update by EntryID (preferred)
& "./scripts/outlook-contacts-update.ps1" -EntryID "00000000..." -Email "newemail@example.com"

# Update by index
& "./scripts/outlook-contacts-update.ps1" -Index 1 -Phone "555-9999" -Company "New Corp"

# Update with search filter
& "./scripts/outlook-contacts-update.ps1" -Index 1 -Search "John" -JobTitle "Director"
```

## Contact Photo

### outlook-contacts-photo.ps1

Add or remove a contact's photo. Supports JPG, JPEG, PNG, GIF, BMP formats.

**Required Parameters:** `-EntryID` or `-Index`, plus `-PhotoPath` or `-Remove`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Contact's unique EntryID (preferred) |
| `-Index` | int | 0 | Contact position number (fallback) |
| `-PhotoPath` | string | — | Path to photo file (JPG, PNG, GIF, BMP) |
| `-Remove` | switch | false | Remove existing photo instead of adding |
| `-Search` | string | — | Filter contacts before applying Index |

```powershell
# Add photo by EntryID
& "./scripts/outlook-contacts-photo.ps1" -EntryID "00000000..." -PhotoPath "C:\Photos\john.jpg"

# Add photo by index
& "./scripts/outlook-contacts-photo.ps1" -Index 1 -PhotoPath "C:\Photos\john.jpg"

# Remove photo
& "./scripts/outlook-contacts-photo.ps1" -EntryID "00000000..." -Remove
```

## Export Contact

### outlook-contacts-export.ps1

Export a contact as vCard (.vcf) file. Default save location is Downloads folder.

**Required Parameters:** `-EntryID` or `-Index`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Contact's unique EntryID (preferred) |
| `-Index` | int | 0 | Contact position number (fallback) |
| `-Path` | string | ~/Downloads | Save path (file or directory) |
| `-Search` | string | — | Filter contacts before applying Index |

```powershell
# Export by EntryID
& "./scripts/outlook-contacts-export.ps1" -EntryID "00000000..."

# Export to specific path
& "./scripts/outlook-contacts-export.ps1" -EntryID "00000000..." -Path "C:\Contacts\john.vcf"

# Export by index
& "./scripts/outlook-contacts-export.ps1" -Index 1
```

## Delete Contact

### outlook-contacts-delete.ps1

Delete a contact from Outlook. Supports targeting by EntryID (preferred), Index, or Name.

**Required Parameters:** One of `-EntryID`, `-Index`, or `-Name`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-EntryID` | string | — | Contact's unique EntryID (preferred) |
| `-Index` | int | 0 | Contact position number (fallback) |
| `-Name` | string | — | Exact full name match (fallback) |

```powershell
# Delete by EntryID (preferred)
& "./scripts/outlook-contacts-delete.ps1" -EntryID "00000000..."

# Delete by index
& "./scripts/outlook-contacts-delete.ps1" -Index 1

# Delete by name
& "./scripts/outlook-contacts-delete.ps1" -Name "John Doe"
```
