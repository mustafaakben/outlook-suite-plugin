param(
    [switch]$Json
)

# Outlook Accounts: every account in the profile (address, type, mailbox) and which one sends
# by default. Use the address with -Account in outlook-send-as / outlook-draft.
# Usage: .\outlook-accounts-list.ps1

. (Join-Path $PSScriptRoot "_common.ps1")

$types = @{ 0 = "Exchange"; 1 = "IMAP"; 2 = "POP3"; 3 = "HTTP"; 4 = "EAS"; 5 = "Other" }

try {
    $null = Connect-Outlook
    $defaultStore = $Namespace.DefaultStore.StoreID
    $list = @(foreach ($a in $Namespace.Accounts) {
        $store = $null; try { $store = $a.DeliveryStore } catch { }
        [ordered]@{
            DisplayName = $a.DisplayName
            Smtp        = $a.SmtpAddress
            Type        = $(if ($types.ContainsKey([int]$a.AccountType)) { $types[[int]$a.AccountType] } else { "Unknown" })
            Mailbox     = $(if ($store) { $store.DisplayName } else { "" })
            Default     = [bool]($store -and $store.StoreID -eq $defaultStore)
        }
    })
    if ($Json) { Write-Json ([ordered]@{ count = $list.Count; accounts = $list }); exit 0 }
    Write-Host "=== OUTLOOK ACCOUNTS ===" -ForegroundColor Cyan
    $n = 0
    foreach ($a in $list) {
        $n++
        Write-Host "$n. $($a.DisplayName)$(if ($a.Default) { '  (default mailbox)' })" -ForegroundColor Yellow
        Write-Host "   Email: $($a.Smtp)   Type: $($a.Type)   Mailbox: $($a.Mailbox)"
    }
    if ($n -eq 0) { Write-Host "No accounts found." -ForegroundColor Gray }
}
catch {
    Stop-WithError $_
}
