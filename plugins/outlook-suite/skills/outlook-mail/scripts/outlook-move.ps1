param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [string]$Folder,

    [int]$Days = 7
)

# Outlook Move Email Script
# Usage: .\outlook-move.ps1 -EntryID "00000000..." -Folder "Archive" (preferred)
# Usage: .\outlook-move.ps1 -Index 1 -Folder "Archive" (fallback)

# Function to find folder recursively
function Find-OutlookFolder {
    param(
        $ParentFolder,
        [string]$FolderName
    )

    # Check direct children first
    foreach ($folder in $ParentFolder.Folders) {
        if ($folder.Name -eq $FolderName) {
            return $folder
        }
    }

    # Then search recursively
    foreach ($folder in $ParentFolder.Folders) {
        $found = Find-OutlookFolder -ParentFolder $folder -FolderName $FolderName
        if ($found) {
            return $found
        }
    }

    return $null
}

$Outlook = $null
try {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)
    $targetEmail = $null

    if ($EntryID) {
        $targetEmail = $Namespace.GetItemFromID($EntryID)
        if (-not $targetEmail) {
            Write-Host "ERROR: Email not found with provided EntryID." -ForegroundColor Red
            Write-Host "Tip: Use outlook-find.ps1 to get valid EntryIDs." -ForegroundColor Gray
        }
    } elseif ($Index -gt 0) {
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
            Write-Host "Tip: Use outlook-read.ps1 -Days $Days to see available emails." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-move.ps1 -EntryID ""00000000..."" -Folder ""Archive""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Find target folder
        $targetFolder = $null

        if ($Folder.Contains("/")) {
            # Path-based navigation: traverse folder hierarchy from account root
            $parts = $Folder -split "/"
            $rootFolder = $Inbox.Parent
            $currentFolder = $rootFolder
            $pathValid = $true
            foreach ($part in $parts) {
                try {
                    $currentFolder = $currentFolder.Folders.Item($part)
                } catch {
                    $pathValid = $false
                    break
                }
            }
            if ($pathValid) {
                $targetFolder = $currentFolder
            }
        } else {
            # Simple name: search recursively
            # First check Inbox subfolders
            $targetFolder = Find-OutlookFolder -ParentFolder $Inbox -FolderName $Folder

            # If not found, search all accounts
            if (-not $targetFolder) {
                foreach ($account in $Namespace.Folders) {
                    $targetFolder = Find-OutlookFolder -ParentFolder $account -FolderName $Folder
                    if ($targetFolder) { break }
                }
            }
        }

        if (-not $targetFolder) {
            Write-Host "`nFolder '$Folder' not found." -ForegroundColor Red
            Write-Host "`nAvailable folders:" -ForegroundColor Yellow
            # Show top-level folders from the default account
            $rootFolder = $Inbox.Parent
            foreach ($f in $rootFolder.Folders) {
                Write-Host "  - $($f.Name)" -ForegroundColor Gray
            }
            Write-Host "`nTip: Use outlook-folders.ps1 to see all folders." -ForegroundColor Gray
        } else {
            # Store email info before moving (COM reference may be invalid after Move)
            $subject = $targetEmail.Subject
            $senderName = $targetEmail.SenderName
            $senderAddr = $targetEmail.SenderEmailAddress
            if ($senderAddr -match "/O=") { $senderAddr = $senderName }
            $receivedDate = $targetEmail.ReceivedTime.ToString("g")

            # Move the email
            $targetEmail.Move($targetFolder) | Out-Null

            Write-Host "`n=== EMAIL MOVED ===" -ForegroundColor Green
            Write-Host "Email: $subject" -ForegroundColor Yellow
            Write-Host "From: $senderName <$senderAddr>"
            Write-Host "Date: $receivedDate"
            Write-Host "Moved to: $($targetFolder.Name)" -ForegroundColor Cyan
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
