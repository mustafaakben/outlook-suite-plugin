# Outlook Accounts List Script
# Usage: .\outlook-accounts-list.ps1
# Lists all email accounts configured in Outlook

$Outlook = $null

# OlAccountType enum mapping (official values from Microsoft docs)
$AccountTypes = @{
    0 = "Exchange"
    1 = "IMAP"
    2 = "POP3"
    3 = "HTTP"
    4 = "EAS"       # Exchange ActiveSync
    5 = "Other"
}

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    Write-Host "`n=== OUTLOOK ACCOUNTS ===" -ForegroundColor Cyan
    Write-Host ""

    $count = 0

    # Method 1: List accounts from Accounts collection
    foreach ($account in $Outlook.Session.Accounts) {
        $count++
        $typeCode = [int]$account.AccountType
        $typeName = if ($AccountTypes.ContainsKey($typeCode)) { $AccountTypes[$typeCode] } else { "Unknown ($typeCode)" }

        Write-Host "$count. $($account.DisplayName)" -ForegroundColor Yellow
        Write-Host "      Email: $($account.SmtpAddress)" -ForegroundColor Gray
        Write-Host "      Type:  $typeName" -ForegroundColor Gray
        Write-Host ""

        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($account) | Out-Null } catch {}
    }

    if ($count -eq 0) {
        # Fallback: List from Folders (stores)
        Write-Host "Accounts from mail stores:" -ForegroundColor Gray
        Write-Host ""
        foreach ($folder in $Outlook.Session.Folders) {
            $count++
            Write-Host "$count. $($folder.Name)" -ForegroundColor Yellow

            # Try to get the associated store
            try {
                $store = $folder.Store
                if ($store) {
                    Write-Host "      Store: $($store.DisplayName)" -ForegroundColor Gray
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null } catch {}
                }
            } catch {}
            Write-Host ""

            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) | Out-Null } catch {}
        }
    }

    if ($count -eq 0) {
        Write-Host "No accounts found." -ForegroundColor Gray
        Write-Host "Check that Outlook is configured with at least one email account." -ForegroundColor Gray
    } else {
        Write-Host "--- Total: $count account(s) ---" -ForegroundColor Cyan
    }

    # Show first account in collection (often default, but not guaranteed)
    Write-Host ""
    Write-Host "First account in Outlook session collection:" -ForegroundColor Cyan
    Write-Host "(This may differ from your effective default send account.)" -ForegroundColor Gray
    try {
        $defaultAccount = $Outlook.Session.Accounts.Item(1)
        Write-Host "  $($defaultAccount.DisplayName) <$($defaultAccount.SmtpAddress)>" -ForegroundColor Green
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($defaultAccount) | Out-Null } catch {}
    } catch {
        Write-Host "  Could not determine the first account entry." -ForegroundColor Gray
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
