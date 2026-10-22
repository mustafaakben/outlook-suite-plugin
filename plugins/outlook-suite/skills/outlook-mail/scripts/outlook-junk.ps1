param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("MarkJunk", "NotJunk", "BlockSender", "SafeSender", "Lists", "OpenOptions")]
    [string]$Action,
    [string]$EntryID = "",
    [string]$StoreID = "",
    [string]$Address = "",
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Junk: junk-mail handling.
#   MarkJunk -EntryID      move the email to Junk Email
#   NotJunk  -EntryID      move it back to the Inbox
#   BlockSender -EntryID   add the email's sender to Blocked Senders (Outlook moves the email to Junk)
#   SafeSender  -EntryID   add the sender to Safe Senders ("Never Block Sender")
#       With the Exchange Online module connected you can pass -Address instead of -EntryID.
#       Without it, Outlook's own Block/Never Block command is run on the email (a message
#       window opens briefly). Both are gated: preview unless -Confirm.
#   Lists        show Blocked / Safe senders (needs the Exchange Online module)
#   OpenOptions  open Outlook's Junk E-mail Options window for the user
# Usage: .\outlook-junk.ps1 -Action BlockSender -EntryID "00000000..." -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $exo = [bool](Get-Command Set-MailboxJunkEmailConfiguration -ErrorAction SilentlyContinue)
    $smtp = ""; try { $smtp = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { }

    switch ($Action) {
        "OpenOptions" {
            $exp = $Outlook.ActiveExplorer()
            if (-not $exp) { Stop-WithError "No Outlook window is open." }
            $exp.CommandBars.ExecuteMso("JunkEmailOptions")
            if ($Json) { Write-Json ([ordered]@{ Action = $Action; Result = "opened" }) } else { Write-Host "Junk E-mail Options opened in Outlook." -ForegroundColor Green }
            exit 0
        }
        "Lists" {
            if (-not $exo) { Stop-WithError "Reading the Blocked/Safe sender lists needs the Exchange Online module (Install-Module ExchangeOnlineManagement -Scope CurrentUser; Connect-ExchangeOnline)." "Or open them in Outlook: -Action OpenOptions." }
            $c = Get-MailboxJunkEmailConfiguration -Identity $smtp
            $o = [ordered]@{ Mailbox = $smtp; Enabled = $c.Enabled; Blocked = @($c.BlockedSendersAndDomains); Safe = @($c.TrustedSendersAndDomains) }
            if ($Json) { Write-Json $o; exit 0 }
            Write-Host "=== JUNK LISTS: $smtp ===" -ForegroundColor Cyan
            Write-Host "  Blocked ($($o.Blocked.Count)): $($o.Blocked -join ', ')"
            Write-Host "  Safe ($($o.Safe.Count)): $($o.Safe -join ', ')"
            exit 0
        }
    }

    if ($Action -in "MarkJunk", "NotJunk") {
        $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
        $target = if ($Action -eq "MarkJunk") { $email.Parent.Store.GetDefaultFolder(23) } else { $email.Parent.Store.GetDefaultFolder(6) }
        $from = $email.Parent.FolderPath
        $moved = $email.Move($target)
        Write-ActionResult -Title $(if ($Action -eq "MarkJunk") { "MOVED TO JUNK" } else { "MOVED BACK TO INBOX" }) -Item $moved -Changes ([ordered]@{ From = $from; To = $target.FolderPath }) -Json:$Json -Action "junk"
        exit 0
    }

    # BlockSender / SafeSender
    $email = $null
    if ($EntryID) { $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID }
    $addr = $Address
    if (-not $addr -and $email) { $addr = Get-SenderSmtp $email }
    if (-not $addr) { Stop-WithError "Give -EntryID of an email from the sender (or -Address with the Exchange Online module)." }
    if (-not $exo -and -not $email) { Stop-WithError "Without the Exchange Online module, give -EntryID of an email from that sender (Outlook's own command runs on it)." }
    $list = if ($Action -eq "BlockSender") { "Blocked Senders" } else { "Safe Senders" }

    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "junk"))) {
        Write-Host "=== $($Action.ToUpper()) (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Add $addr to $list$(if ($email) { " (from: $($email.Subject))" })"
        Write-Host "  Via: $(if ($exo) { 'Exchange Online module' } else { 'Outlook command on the email (a message window opens briefly)' })"
        if ($Action -eq "BlockSender") { Write-Host "  Outlook also moves the email to Junk Email." -ForegroundColor Gray }
        Approve-Action -Action "junk" -Confirmed $false -What "changing the junk lists" | Out-Null
        exit 0
    }

    if ($exo) {
        $p = @{ Identity = $smtp }
        if ($Action -eq "BlockSender") { $p.BlockedSendersAndDomainsToAdd = @($addr) } else { $p.TrustedSendersAndDomainsToAdd = @($addr) }
        Set-MailboxJunkEmailConfiguration @p
    } else {
        $id = if ($Action -eq "BlockSender") { "JunkEmailAddToBlockedSendersList" } else { "JunkEmailAddToSafeSendersList" }
        $email.Display()
        $insp = $email.GetInspector
        if (-not $insp.CommandBars.GetEnabledMso($id)) { $insp.Close(1); Stop-WithError "Outlook's '$id' command is not available for this item." }
        $insp.CommandBars.ExecuteMso($id)
        try { $insp.Close(1) } catch { }
    }
    $o = [ordered]@{ Action = $Action; Address = $addr; List = $list; Via = $(if ($exo) { "Exchange Online module" } else { "Outlook command" }) }
    if ($Json) { Write-Json $o; exit 0 }
    Write-Host "=== $addr ADDED TO $($list.ToUpper()) ===" -ForegroundColor Green
}
catch {
    Stop-WithError $_
}
