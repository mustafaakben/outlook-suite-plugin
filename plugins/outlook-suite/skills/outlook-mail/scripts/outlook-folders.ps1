# Outlook Folder Lister
# Usage: .\outlook-folders.ps1

$Outlook = $null
$Namespace = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    Write-Host "`n=== OUTLOOK FOLDERS ===" -ForegroundColor Cyan

    foreach ($account in $Namespace.Folders) {
        Write-Host "`n$($account.Name)" -ForegroundColor Yellow
        Write-Host ("=" * 40)

        foreach ($folder in $account.Folders) {
            $count = $folder.Items.Count
            $unread = $folder.UnReadItemCount
            $unreadInfo = if ($unread -gt 0) { " ($unread unread)" } else { "" }

            Write-Host "  $($folder.Name): $count items$unreadInfo"

            # Show subfolders
            foreach ($subfolder in $folder.Folders) {
                $subUnread = $subfolder.UnReadItemCount
                $subUnreadInfo = if ($subUnread -gt 0) { " ($subUnread unread)" } else { "" }
                Write-Host "    - $($subfolder.Name): $($subfolder.Items.Count) items$subUnreadInfo" -ForegroundColor Gray
            }
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
