param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Path = "",
    [int]$Days = 7,
    [int]$AttachmentIndex = 0,
    [switch]$ListOnly,
    [switch]$IncludeInline
)

# Outlook Attachment Save Script
# Save email attachments to disk
# Usage: .\outlook-attachment-save.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-attachment-save.ps1 -Index 1 (fallback)
# Save to specific directory: .\outlook-attachment-save.ps1 -EntryID "00000000..." -Path "C:\Downloads"
# Save specific attachment: .\outlook-attachment-save.ps1 -EntryID "00000000..." -AttachmentIndex 2
# List attachments only: .\outlook-attachment-save.ps1 -EntryID "00000000..." -ListOnly
# Include inline images: .\outlook-attachment-save.ps1 -EntryID "00000000..." -IncludeInline

if (-not $Path) {
    $Path = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
}

$Outlook = $null
$Namespace = $null
$Inbox = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")
    $targetEmail = $null

    if ($EntryID) {
        $targetEmail = $Namespace.GetItemFromID($EntryID)
        if (-not $targetEmail) {
            Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
        $Inbox = $Namespace.GetDefaultFolder(6)

        # Get emails with locale-safe date
        $since = (Get-Date).AddDays(-$Days).ToString("g")
        $filter = "[ReceivedTime] >= '$since'"

        $items = $Inbox.Items.Restrict($filter)
        $items.Sort("[ReceivedTime]", $true)

        $count = 0
        foreach ($email in $items) {
            $count++
            if ($count -eq $Index) {
                $targetEmail = $email
                break
            }
        }

        if (-not $targetEmail) {
            Write-Host "`nEmail at index $Index not found in last $Days days." -ForegroundColor Red
            Write-Host "Use outlook-read.ps1 -Days $Days to see available emails." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-attachment-save.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if (-not $targetEmail) {
        # Errors already displayed above
    }
    elseif ($targetEmail.Attachments.Count -eq 0) {
        Write-Host "`nEmail has no attachments." -ForegroundColor Yellow
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Gray
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }
        Write-Host "From: $($targetEmail.SenderName) <$senderAddr>" -ForegroundColor Gray
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray
    }
    else {
        # Validate AttachmentIndex range
        if ($AttachmentIndex -gt $targetEmail.Attachments.Count) {
            Write-Host "`nAttachment index $AttachmentIndex is out of range." -ForegroundColor Red
            Write-Host "Email has $($targetEmail.Attachments.Count) attachment(s). Use -AttachmentIndex 1 to $($targetEmail.Attachments.Count)." -ForegroundColor Gray
        }
        else {
            # Show email info
            $senderAddr = $targetEmail.SenderEmailAddress
            if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

            Write-Host "`n=== ATTACHMENTS ===" -ForegroundColor Cyan
            Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "From: $($targetEmail.SenderName) <$senderAddr>" -ForegroundColor Gray
            Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))" -ForegroundColor Gray

            # Build attachment list, detecting inline images
            $attachments = [System.Collections.ArrayList]::new()
            $contentIdProp = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

            for ($i = 1; $i -le $targetEmail.Attachments.Count; $i++) {
                $att = $targetEmail.Attachments.Item($i)
                $isInline = $false
                try {
                    $cid = $att.PropertyAccessor.GetProperty($contentIdProp)
                    if ($cid) { $isInline = $true }
                } catch { }

                $sizeKB = [math]::Round($att.Size / 1KB, 1)
                [void]$attachments.Add(@{
                    Index    = $i
                    FileName = $att.FileName
                    SizeKB   = $sizeKB
                    IsInline = $isInline
                    Object   = $att
                })
            }

            $regularCount = ($attachments | Where-Object { -not $_.IsInline }).Count
            $inlineCount  = ($attachments | Where-Object { $_.IsInline }).Count

            Write-Host "Attachments: $regularCount file(s)" -NoNewline -ForegroundColor Gray
            if ($inlineCount -gt 0) {
                Write-Host " + $inlineCount inline/embedded" -NoNewline -ForegroundColor DarkGray
            }
            Write-Host ""

            # List all attachments
            Write-Host ""
            foreach ($att in $attachments) {
                $marker = ""
                $color = "White"
                if ($att.IsInline) {
                    $marker = " [inline]"
                    $color = "DarkGray"
                }
                Write-Host "  [$($att.Index)] $($att.FileName) ($($att.SizeKB) KB)$marker" -ForegroundColor $color
            }

            if ($ListOnly) {
                Write-Host "`nUse -AttachmentIndex <number> to save a specific attachment." -ForegroundColor Gray
                Write-Host "Use -IncludeInline to also save inline/embedded images." -ForegroundColor Gray
            }
            else {
                # Ensure path exists
                if (-not (Test-Path $Path)) {
                    New-Item -ItemType Directory -Path $Path -Force | Out-Null
                }

                Write-Host "`nSaving to: $Path" -ForegroundColor Gray

                $savedCount = 0
                foreach ($att in $attachments) {
                    # Skip if specific attachment requested and this isn't it
                    if ($AttachmentIndex -gt 0 -and $att.Index -ne $AttachmentIndex) {
                        continue
                    }

                    # Skip inline attachments unless -IncludeInline
                    if ($att.IsInline -and -not $IncludeInline -and $AttachmentIndex -eq 0) {
                        continue
                    }

                    $fileName = $att.FileName
                    $filePath = Join-Path $Path $fileName

                    # Handle duplicate filenames
                    $counter = 1
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                    $extension = [System.IO.Path]::GetExtension($fileName)
                    while (Test-Path $filePath) {
                        $fileName = "${baseName}_${counter}${extension}"
                        $filePath = Join-Path $Path $fileName
                        $counter++
                    }

                    $att.Object.SaveAsFile($filePath)
                    $savedCount++

                    $actualSize = [math]::Round((Get-Item $filePath).Length / 1KB, 1)
                    Write-Host "  Saved: $fileName ($actualSize KB)" -ForegroundColor Green
                }

                if ($savedCount -eq 0 -and $inlineCount -gt 0 -and $AttachmentIndex -eq 0) {
                    Write-Host "`nNo regular attachments saved ($inlineCount inline skipped)." -ForegroundColor Yellow
                    Write-Host "Use -IncludeInline to save inline/embedded images." -ForegroundColor Gray
                }
                elseif ($savedCount -eq 0) {
                    Write-Host "`nNo attachments saved." -ForegroundColor Yellow
                }
                else {
                    Write-Host "`n--- Saved $savedCount attachment(s) ---" -ForegroundColor Cyan
                }
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Inbox) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Inbox) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
