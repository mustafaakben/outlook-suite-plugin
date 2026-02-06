param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Low", "Normal", "High")]
    [string]$Importance,

    [int]$Days = 7
)

# Outlook Set Importance Script
# Usage: .\outlook-set-importance.ps1 -EntryID "00000000..." -Importance High (preferred)
# Usage: .\outlook-set-importance.ps1 -Index 1 -Importance High (fallback)

$importanceMap = @{
    "Low" = 0
    "Normal" = 1
    "High" = 2
}

$importanceNames = @{
    0 = "Low"
    1 = "Normal"
    2 = "High"
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
        Write-Host "Usage: .\outlook-set-importance.ps1 -EntryID ""00000000..."" -Importance High" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        # Get previous importance
        $previousImportance = $importanceNames[$targetEmail.Importance]

        # Set new importance
        $targetEmail.Importance = $importanceMap[$Importance]
        $targetEmail.Save()

        # Color based on importance
        $color = switch ($Importance) {
            "High" { "Red" }
            "Normal" { "Yellow" }
            "Low" { "Gray" }
        }

        Write-Host "`n=== IMPORTANCE SET ===" -ForegroundColor Green
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
        Write-Host "Importance: $Importance" -ForegroundColor $color
        if ($previousImportance -ne $Importance) {
            Write-Host "Previous: $previousImportance" -ForegroundColor Gray
        }
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
