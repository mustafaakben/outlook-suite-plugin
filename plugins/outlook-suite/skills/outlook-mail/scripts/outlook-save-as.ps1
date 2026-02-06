param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Path = "",
    [ValidateSet("msg", "txt", "html", "mht")]
    [string]$Format = "msg",
    [int]$Days = 7
)

# Outlook Save Email As File Script
# Usage: .\outlook-save-as.ps1 -EntryID "00000000..." -Format txt (preferred)
# Usage: .\outlook-save-as.ps1 -Index 1 -Path "C:\Emails\email.msg" (fallback)

$formatMap = @{
    "txt" = 0
    "msg" = 3
    "html" = 5
    "mht" = 9
}

$formatExtensions = @{
    "txt" = ".txt"
    "msg" = ".msg"
    "html" = ".html"
    "mht" = ".mht"
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
        Write-Host "Usage: .\outlook-save-as.ps1 -EntryID ""00000000..."" -Format txt" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback
        $senderAddr = $targetEmail.SenderEmailAddress
        if ($senderAddr -match "^/O=") { $senderAddr = $targetEmail.SenderName }

        # Generate filename if path not provided or is a directory
        if (-not $Path -or (Test-Path $Path -PathType Container)) {
            $folder = if ($Path) { $Path } else { [Environment]::GetFolderPath("UserProfile") + "\Downloads" }

            # Clean subject for filename
            $cleanSubject = ""
            if (-not [string]::IsNullOrWhiteSpace($targetEmail.Subject)) {
                $cleanSubject = $targetEmail.Subject -replace '[\\/:*?"<>|]', '_'
                $cleanSubject = $cleanSubject.Trim()
                if ($cleanSubject.Length -gt 50) {
                    $cleanSubject = $cleanSubject.Substring(0, 50)
                }
            }
            if ([string]::IsNullOrWhiteSpace($cleanSubject)) {
                $cleanSubject = "email_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            }

            $fileName = $cleanSubject + $formatExtensions[$Format]
            $Path = Join-Path $folder $fileName
        }

        # Ensure directory exists
        $directory = [System.IO.Path]::GetDirectoryName($Path)
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Handle duplicate filenames
        $basePath = [System.IO.Path]::GetDirectoryName($Path)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $extension = [System.IO.Path]::GetExtension($Path)
        $counter = 1
        while (Test-Path $Path) {
            $Path = Join-Path $basePath "${baseName}_${counter}${extension}"
            $counter++
        }

        # Save the email
        $targetEmail.SaveAs($Path, $formatMap[$Format])

        $fileSize = [math]::Round((Get-Item $Path).Length / 1KB, 1)

        Write-Host "`n=== EMAIL SAVED ===" -ForegroundColor Green
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($targetEmail.SenderName) <$senderAddr>"
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
        Write-Host "Format: $($Format.ToUpper())" -ForegroundColor Cyan
        Write-Host "Size: $fileSize KB"
        Write-Host "Path: $Path" -ForegroundColor Gray
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
