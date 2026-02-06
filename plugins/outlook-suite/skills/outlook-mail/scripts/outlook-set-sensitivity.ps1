param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Normal", "Personal", "Private", "Confidential")]
    [string]$Sensitivity,

    [int]$Days = 7
)

# Outlook Set Sensitivity Script
# Usage: .\outlook-set-sensitivity.ps1 -EntryID "00000000..." -Sensitivity Confidential (preferred)
# Usage: .\outlook-set-sensitivity.ps1 -Index 1 -Sensitivity Confidential (fallback)

$sensitivityMap = @{
    "Normal" = 0
    "Personal" = 1
    "Private" = 2
    "Confidential" = 3
}

$sensitivityNames = @{
    0 = "Normal"
    1 = "Personal"
    2 = "Private"
    3 = "Confidential"
}

$Outlook = $null
$Namespace = $null

try {
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
            Write-Host "Tip: Try increasing -Days if the email is older." -ForegroundColor Gray
        }
    } else {
        Write-Host "ERROR: Provide -EntryID or -Index to target an email." -ForegroundColor Red
        Write-Host "Usage: .\outlook-set-sensitivity.ps1 -EntryID ""00000000..."" -Sensitivity Confidential" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        # Get previous sensitivity
        $previousSensitivity = $sensitivityNames[$targetEmail.Sensitivity]

        # Set new sensitivity
        try {
            $targetEmail.Sensitivity = $sensitivityMap[$Sensitivity]
            $targetEmail.Save()
        } catch {
            Write-Host "`n[!] Cannot set sensitivity on this email." -ForegroundColor Red
            Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
            Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
            Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
            Write-Host "Current sensitivity: $previousSensitivity" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Note: Sensitivity can only be set on emails you compose (drafts/new messages)." -ForegroundColor Gray
            Write-Host "Received emails may be read-only depending on your Exchange/account settings." -ForegroundColor Gray
            return
        }

        # Color based on sensitivity
        $color = switch ($Sensitivity) {
            "Confidential" { "Red" }
            "Private" { "Magenta" }
            "Personal" { "Yellow" }
            "Normal" { "Gray" }
        }

        Write-Host "`n=== SENSITIVITY SET ===" -ForegroundColor Green
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
        Write-Host "Sensitivity: $Sensitivity" -ForegroundColor $color
        if ($previousSensitivity -ne $Sensitivity) {
            Write-Host "Previous: $previousSensitivity" -ForegroundColor Gray
        }

        # Show what sensitivity means
        $description = switch ($Sensitivity) {
            "Normal" { "No special restrictions" }
            "Personal" { "Treat as personal information" }
            "Private" { "Do not forward or share" }
            "Confidential" { "Highly sensitive - restricted access" }
        }
        Write-Host "($description)" -ForegroundColor Gray
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
