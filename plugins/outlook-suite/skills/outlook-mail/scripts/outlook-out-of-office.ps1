param(
    [switch]$Enable,
    [switch]$Disable,
    [switch]$Status,
    [string]$InternalMessage = "",
    [string]$ExternalMessage = "",
    [datetime]$StartTime,
    [datetime]$EndTime,
    [switch]$ExternalAudienceAll,
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Out of Office (automatic replies), Exchange / Microsoft 365 mailboxes.
#   -Status                 on/off from the mailbox itself (no extra modules needed)
#   -Enable / -Disable      switch automatic replies on/off. Uses the reply text already saved
#                           in Outlook. Outward-facing (people get auto-replies): previews unless -Confirm.
#   Messages, schedule, audience (-InternalMessage, -ExternalMessage, -StartTime/-EndTime,
#   -ExternalAudienceAll) need the Exchange Online module, connected:
#     Install-Module ExchangeOnlineManagement -Scope CurrentUser; Connect-ExchangeOnline
#   When it is connected this script uses it for everything.
# Usage: .\outlook-out-of-office.ps1 -Status
# Usage: .\outlook-out-of-office.ps1 -Enable -Confirm
# Usage: .\outlook-out-of-office.ps1 -Enable -InternalMessage "Away until Monday." -StartTime "2026-11-01" -EndTime "2026-11-05" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")
$PR_OOF_STATE = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"

try {
    if (-not ($Enable -or $Disable -or $Status)) { Stop-WithError "Choose -Status, -Enable or -Disable." }
    if ($Enable -and $Disable) { Stop-WithError "Choose -Enable or -Disable, not both." }
    if (($StartTime -and -not $EndTime) -or ($EndTime -and -not $StartTime)) { Stop-WithError "Give both -StartTime and -EndTime." }
    if ($StartTime -and $EndTime -and $EndTime -le $StartTime) { Stop-WithError "-EndTime must be after -StartTime." }

    $null = Connect-Outlook
    $store = $Namespace.DefaultStore
    $smtp = ""; try { $smtp = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { }
    if (-not $smtp) { Stop-WithError "Automatic replies need an Exchange / Microsoft 365 mailbox as the default store." }
    $exo = [bool](Get-Command Get-MailboxAutoReplyConfiguration -ErrorAction SilentlyContinue)
    $needsExo = $InternalMessage -or $ExternalMessage -or $StartTime -or $ExternalAudienceAll

    $state = $null; try { $state = [bool]$store.PropertyAccessor.GetProperty($PR_OOF_STATE) } catch { }

    if ($Status) {
        $o = [ordered]@{ Mailbox = $smtp; AutomaticReplies = $(if ($state) { "On" } elseif ($state -eq $false) { "Off" } else { "Unknown" }) }
        if ($exo) {
            $c = Get-MailboxAutoReplyConfiguration -Identity $smtp
            $o.State = "$($c.AutoReplyState)"
            if ($c.AutoReplyState -eq "Scheduled") { $o.Start = Format-MailDate $c.StartTime; $o.End = Format-MailDate $c.EndTime }
            $o.ExternalAudience = "$($c.ExternalAudience)"
            $o.InternalMessage = (Format-MailLinks -Text (($c.InternalMessage -replace '<[^>]+>', ' ') -replace '\s+', ' ').Trim()).Text
            $o.ExternalMessage = (Format-MailLinks -Text (($c.ExternalMessage -replace '<[^>]+>', ' ') -replace '\s+', ' ').Trim()).Text
        } else {
            $o.Note = "Message text and schedule need the Exchange Online module (Connect-ExchangeOnline)."
        }
        if ($Json) { Write-Json $o; exit 0 }
        Write-Host "=== AUTOMATIC REPLIES ===" -ForegroundColor Cyan
        foreach ($k in $o.Keys) { Write-Host "  ${k}: $($o[$k])" }
        exit 0
    }

    if ($needsExo -and -not $exo) {
        Stop-WithError "Setting the reply text, schedule or audience needs the Exchange Online module." "Install-Module ExchangeOnlineManagement -Scope CurrentUser; Connect-ExchangeOnline -UserPrincipalName $smtp ; then run this again. To just switch replies on with the saved text: -Enable -Confirm (no message parameters)."
    }

    $want = [bool]$Enable
    $what = if ($want) { "turning automatic replies ON (senders will get auto-replies)" } else { "turning automatic replies OFF" }
    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "out-of-office"))) {
        Write-Host "=== AUTOMATIC REPLIES (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Mailbox: $smtp"
        Write-Host "  Now: $(if ($state) { 'On' } else { 'Off' })   Requested: $(if ($want) { 'On' } else { 'Off' })"
        if ($StartTime) { Write-Host "  Schedule: $(Format-MailDate $StartTime) to $(Format-MailDate $EndTime)" }
        if ($InternalMessage) { Write-Host "  Internal message: $(Format-Preview $InternalMessage 300)" }
        if ($ExternalMessage) { Write-Host "  External message: $(Format-Preview $ExternalMessage 300) (audience: $(if ($ExternalAudienceAll) { 'All' } else { 'Known contacts' }))" }
        Approve-Action -Action "out-of-office" -Confirmed $false -What $what | Out-Null
        exit 0
    }

    if ($exo) {
        $p = @{ Identity = $smtp }
        if (-not $want) { $p.AutoReplyState = "Disabled" }
        else {
            $p.AutoReplyState = if ($StartTime) { "Scheduled" } else { "Enabled" }
            if ($StartTime) { $p.StartTime = $StartTime; $p.EndTime = $EndTime }
            if ($InternalMessage) { $p.InternalMessage = $InternalMessage }
            if ($ExternalMessage) { $p.ExternalMessage = $ExternalMessage; $p.ExternalAudience = $(if ($ExternalAudienceAll) { "All" } else { "Known" }) }
        }
        Set-MailboxAutoReplyConfiguration @p
    } else {
        $store.PropertyAccessor.SetProperty($PR_OOF_STATE, $want)
    }
    $now = $null; try { $now = [bool]$store.PropertyAccessor.GetProperty($PR_OOF_STATE) } catch { }
    $o = [ordered]@{ Mailbox = $smtp; AutomaticReplies = $(if ($now) { "On" } else { "Off" }); Via = $(if ($exo) { "Exchange Online module" } else { "mailbox property (saved reply text)" }) }
    if ($Json) { Write-Json $o; exit 0 }
    Write-Host "=== AUTOMATIC REPLIES $(if ($want) { 'ON' } else { 'OFF' }) ===" -ForegroundColor Green
    foreach ($k in $o.Keys) { Write-Host "  ${k}: $($o[$k])" }
}
catch {
    Stop-WithError $_
}
