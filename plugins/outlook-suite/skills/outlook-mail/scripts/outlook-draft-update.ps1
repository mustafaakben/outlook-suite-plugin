param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$To = "",
    [string]$CC = "",
    [string]$BCC = "",
    [string]$Subject = "",
    [string]$Body = "",
    [switch]$AppendBody,
    [string[]]$AddAttachment = @(),
    [switch]$HTML,
    [int]$Limit = 20
)

# Outlook Draft Update Script
# Usage: .\outlook-draft-update.ps1 -EntryID "00000000..." -Subject "Updated Subject" (preferred)
# Usage: .\outlook-draft-update.ps1 -Index 1 -Subject "Updated Subject" (fallback)

$Outlook = $null
$targetDraft = $null
try {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    if ($EntryID) {
        # EntryID = primary (direct lookup)
        $targetDraft = $Namespace.GetItemFromID($EntryID)
        if (-not $targetDraft) {
            Write-Host "ERROR: Draft not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-find.ps1 -Folder Drafts to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        # Index = fallback (scan)
        $Drafts = $Namespace.GetDefaultFolder(16)  # olFolderDrafts = 16
        $items = $Drafts.Items
        $items.Sort("[LastModificationTime]", $true)

        $count = 0
        foreach ($draft in $items) {
            $count++
            if ($count -eq $Index) { $targetDraft = $draft }
            if ($count -ge $Limit) { break }
        }

        if (-not $targetDraft) {
            Write-Host "`nDraft at index $Index not found." -ForegroundColor Red
            Write-Host "There are $count drafts available (showing up to $Limit)." -ForegroundColor Gray
            Write-Host "Tip: Use -Limit parameter to see more drafts." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target a draft." -ForegroundColor Red
        Write-Host "Usage: .\outlook-draft-update.ps1 -EntryID ""00000000..."" -Subject ""New Subject""" -ForegroundColor Gray
        Write-Host "       .\outlook-draft-update.ps1 -Index 1 -Subject ""New Subject""" -ForegroundColor Gray
    }

    if ($targetDraft) {
        # Track changes
        $changes = @()

        # Update fields if provided
        if ($To) {
            $oldTo = $targetDraft.To
            $targetDraft.To = $To
            if ($oldTo) {
                $changes += "To: '$oldTo' -> '$To'"
            } else {
                $changes += "To: Set to '$To'"
            }
        }

        if ($CC) {
            $oldCC = $targetDraft.CC
            $targetDraft.CC = $CC
            if ($oldCC) {
                $changes += "CC: '$oldCC' -> '$CC'"
            } else {
                $changes += "CC: Set to '$CC'"
            }
        }

        if ($BCC) {
            $oldBCC = $targetDraft.BCC
            $targetDraft.BCC = $BCC
            if ($oldBCC) {
                $changes += "BCC: '$oldBCC' -> '$BCC'"
            } else {
                $changes += "BCC: Set to '$BCC'"
            }
        }

        if ($Subject) {
            $oldSubject = $targetDraft.Subject
            $targetDraft.Subject = $Subject
            $changes += "Subject: '$oldSubject' -> '$Subject'"
        }

        if ($Body) {
            if ($AppendBody) {
                if ($HTML) {
                    $targetDraft.HTMLBody = $targetDraft.HTMLBody + "<br><br>" + $Body
                } else {
                    $targetDraft.Body = $targetDraft.Body + "`n`n" + $Body
                }
                $changes += "Body: Appended text"
            } else {
                if ($HTML) {
                    $targetDraft.HTMLBody = $Body
                } else {
                    $targetDraft.Body = $Body
                }
                $changes += "Body: Replaced"
            }
        }

        # Attach files with validation
        foreach ($file in $AddAttachment) {
            if (-not $file) { continue }
            if (Test-Path $file) {
                $targetDraft.Attachments.Add((Resolve-Path $file).Path) | Out-Null
                $fileName = [System.IO.Path]::GetFileName($file)
                $changes += "Attachment: Added '$fileName'"
            } else {
                Write-Host "Warning: Attachment not found - $file" -ForegroundColor Yellow
            }
        }

        if ($changes.Count -eq 0) {
            Write-Host "`nNo changes specified. Use parameters like -To, -Subject, -Body, -AddAttachment, etc." -ForegroundColor Yellow
        } else {
            $targetDraft.Save()

            Write-Host "`n=== DRAFT UPDATED ===" -ForegroundColor Green

            $subjectDisplay = if ($targetDraft.Subject) { $targetDraft.Subject } else { "(No Subject)" }
            Write-Host "Subject: $subjectDisplay" -ForegroundColor Yellow

            if ($targetDraft.To) { Write-Host "To: $($targetDraft.To)" }
            if ($targetDraft.CC) { Write-Host "CC: $($targetDraft.CC)" }
            if ($targetDraft.BCC) { Write-Host "BCC: $($targetDraft.BCC)" }

            $attachCount = $targetDraft.Attachments.Count
            if ($attachCount -gt 0) {
                Write-Host "Attachments: $attachCount file(s)" -ForegroundColor Gray
            }

            Write-Host "`nChanges made:" -ForegroundColor Cyan
            foreach ($change in $changes) {
                Write-Host "  - $change" -ForegroundColor Gray
            }

            Write-Host "`nDraft saved. Use outlook-send.ps1 to send when ready." -ForegroundColor Cyan
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($targetDraft) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetDraft) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
