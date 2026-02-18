param(
    [switch]$Enable,
    [switch]$Disable,
    [switch]$Status,
    [string]$InternalMessage = "",
    [string]$ExternalMessage = "",
    [datetime]$StartTime,
    [datetime]$EndTime,
    [switch]$ExternalAudienceAll,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Out of Office Script
# Configure automatic reply / out of office (Exchange / Microsoft 365)
# Usage: .\outlook-out-of-office.ps1 -Status
# Enable: .\outlook-out-of-office.ps1 -Enable -InternalMessage "I'm out of office until Monday."
# With dates: .\outlook-out-of-office.ps1 -Enable -InternalMessage "Away" -StartTime "2026-02-10" -EndTime "2026-02-15"
# Disable: .\outlook-out-of-office.ps1 -Disable

# Note: Outlook COM does not expose Out of Office settings directly (server-side Exchange feature).
# This script detects the Exchange account via Outlook COM, then:
#   - If ExchangeOnlineManagement module is connected: runs commands directly
#   - Otherwise: generates ready-to-run PowerShell commands

$Outlook = $null
$Namespace = $null

try {
    # Validate: must specify an action
    if (-not $Enable -and -not $Disable -and -not $Status) {
        Write-Host "`n=== OUT OF OFFICE ===" -ForegroundColor Cyan
        Write-Host "`nUsage:" -ForegroundColor Yellow
        Write-Host "  Check status:  .\outlook-out-of-office.ps1 -Status" -ForegroundColor Gray
        Write-Host "  Enable:        .\outlook-out-of-office.ps1 -Enable -InternalMessage 'I am away'" -ForegroundColor Gray
        Write-Host "  Disable:       .\outlook-out-of-office.ps1 -Disable" -ForegroundColor Gray
        Write-Host "`nWith scheduled times:" -ForegroundColor Gray
        Write-Host "  .\outlook-out-of-office.ps1 -Enable -InternalMessage 'Away' -StartTime '2026-02-10' -EndTime '2026-02-15'" -ForegroundColor Gray
        Write-Host "`nWith external message:" -ForegroundColor Gray
        Write-Host "  .\outlook-out-of-office.ps1 -Enable -InternalMessage 'Away' -ExternalMessage 'Out of office' -ExternalAudienceAll" -ForegroundColor Gray
    }
    elseif ($Enable -and $Disable) {
        Write-Host "`nCannot use -Enable and -Disable together." -ForegroundColor Red
    }
    elseif ($Enable -and -not $InternalMessage) {
        Write-Host "`n-InternalMessage is required when enabling Out of Office." -ForegroundColor Red
        Write-Host "Example: .\outlook-out-of-office.ps1 -Enable -InternalMessage 'I am away until Monday.'" -ForegroundColor Gray
    }
    elseif ($Enable -and $StartTime -and -not $EndTime) {
        Write-Host "`n-EndTime is required when -StartTime is specified." -ForegroundColor Red
        Write-Host "Example: .\outlook-out-of-office.ps1 -Enable -InternalMessage 'Away' -StartTime '2026-02-10' -EndTime '2026-02-15'" -ForegroundColor Gray
    }
    elseif ($Enable -and -not $StartTime -and $EndTime) {
        Write-Host "`n-StartTime is required when -EndTime is specified." -ForegroundColor Red
        Write-Host "Example: .\outlook-out-of-office.ps1 -Enable -InternalMessage 'Away' -StartTime '2026-02-10' -EndTime '2026-02-15'" -ForegroundColor Gray
    }
    elseif ($Enable -and $StartTime -and $EndTime -and $EndTime -le $StartTime) {
        Write-Host "`n-EndTime must be later than -StartTime." -ForegroundColor Red
        Write-Host "Example: .\outlook-out-of-office.ps1 -Enable -InternalMessage 'Away' -StartTime '2026-02-10 09:00' -EndTime '2026-02-15 17:00'" -ForegroundColor Gray
    }
    else {
        # Connect to Outlook - try active instance first
        try {
            $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $Outlook = New-Object -ComObject Outlook.Application
            Start-Sleep -Milliseconds 500
        }

        $Namespace = $Outlook.GetNamespace("MAPI")

        # Detect Exchange account
        $exchangeUser = $null
        try {
            $exchangeUser = $Namespace.CurrentUser.AddressEntry.GetExchangeUser()
        } catch { }

        if (-not $exchangeUser) {
            Write-Host "`nOut of Office requires Microsoft Exchange or Office 365." -ForegroundColor Red
            Write-Host "This feature is not available for POP/IMAP accounts." -ForegroundColor Gray
        }
        else {
            $smtpAddress = $exchangeUser.PrimarySmtpAddress

            Write-Host "`n=== OUT OF OFFICE ===" -ForegroundColor Cyan
            Write-Host "Account: $smtpAddress" -ForegroundColor Gray

            # Check if Exchange Online Management module is connected
            $exchangeConnected = $false
            try {
                $null = Get-Command Get-MailboxAutoReplyConfiguration -ErrorAction Stop
                $exchangeConnected = $true
            } catch { }

            if ($Status) {
                if ($exchangeConnected) {
                    Write-Host "`nQuerying current settings..." -ForegroundColor Gray
                    try {
                        $config = Get-MailboxAutoReplyConfiguration -Identity $smtpAddress
                        $stateColor = if ($config.AutoReplyState -eq "Disabled") { "Gray" } else { "Green" }
                        Write-Host "`nStatus: $($config.AutoReplyState)" -ForegroundColor $stateColor

                        if ($config.AutoReplyState -ne "Disabled") {
                            if ($config.AutoReplyState -eq "Scheduled") {
                                Write-Host "Schedule: $($config.StartTime.ToString('g')) - $($config.EndTime.ToString('g'))" -ForegroundColor Yellow
                            }
                            Write-Host "External Audience: $($config.ExternalAudience)" -ForegroundColor Gray

                            if ($config.InternalMessage) {
                                Write-Host "`nInternal Reply:" -ForegroundColor Cyan
                                $internalText = $config.InternalMessage -replace '<[^>]+>', ''
                                $internalText = $internalText.Trim()
                                if ($StripLinks) { $internalText = Strip-Links $internalText }
                                Write-Host "  $internalText" -ForegroundColor Gray
                            }
                            if ($config.ExternalMessage) {
                                Write-Host "`nExternal Reply:" -ForegroundColor Cyan
                                $externalText = $config.ExternalMessage -replace '<[^>]+>', ''
                                $externalText = $externalText.Trim()
                                if ($StripLinks) { $externalText = Strip-Links $externalText }
                                Write-Host "  $externalText" -ForegroundColor Gray
                            }
                        }
                    } catch {
                        Write-Host "Failed to query: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nCommand to check status:" -ForegroundColor Yellow
                    Write-Host "  Get-MailboxAutoReplyConfiguration -Identity $smtpAddress" -ForegroundColor Cyan
                }
            }
            elseif ($Disable) {
                if ($exchangeConnected) {
                    Write-Host "`nDisabling Out of Office..." -ForegroundColor Gray
                    try {
                        Set-MailboxAutoReplyConfiguration -Identity $smtpAddress -AutoReplyState Disabled
                        Write-Host "`nOut of Office DISABLED." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nCommand to disable:" -ForegroundColor Yellow
                    Write-Host "  Set-MailboxAutoReplyConfiguration -Identity $smtpAddress -AutoReplyState Disabled" -ForegroundColor Cyan
                }
            }
            elseif ($Enable) {
                # Determine state
                $replyState = "Enabled"
                if ($StartTime -and $EndTime) {
                    $replyState = "Scheduled"
                }

                # Show preview
                Write-Host "`n--- Preview ---" -ForegroundColor Yellow
                Write-Host "State: $replyState" -ForegroundColor Green
                if ($replyState -eq "Scheduled") {
                    Write-Host "Schedule: $($StartTime.ToString('g')) to $($EndTime.ToString('g'))" -ForegroundColor Yellow
                }
                $internalPreview = $InternalMessage
                if ($StripLinks) { $internalPreview = Strip-Links $internalPreview }
                Write-Host "Internal Message: $internalPreview" -ForegroundColor Gray
                if ($ExternalMessage) {
                    $audience = if ($ExternalAudienceAll) { "All" } else { "Known contacts only" }
                    $externalPreview = $ExternalMessage
                    if ($StripLinks) { $externalPreview = Strip-Links $externalPreview }
                    Write-Host "External Message: $externalPreview" -ForegroundColor Gray
                    Write-Host "External Audience: $audience" -ForegroundColor Gray
                }

                if ($exchangeConnected) {
                    Write-Host "`nEnabling Out of Office..." -ForegroundColor Gray
                    try {
                        $params = @{
                            Identity       = $smtpAddress
                            AutoReplyState = $replyState
                            InternalMessage = $InternalMessage
                        }
                        if ($StartTime -and $EndTime) {
                            $params.StartTime = $StartTime
                            $params.EndTime = $EndTime
                        }
                        if ($ExternalMessage) {
                            $params.ExternalMessage = $ExternalMessage
                            $params.ExternalAudience = if ($ExternalAudienceAll) { "All" } else { "Known" }
                        }
                        Set-MailboxAutoReplyConfiguration @params
                        Write-Host "`nOut of Office ENABLED." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    # Generate ready-to-run command
                    $psCommand = "Set-MailboxAutoReplyConfiguration -Identity $smtpAddress"
                    $psCommand += " -AutoReplyState $replyState"

                    if ($StartTime -and $EndTime) {
                        $psCommand += " -StartTime `"$($StartTime.ToString('g'))`""
                        $psCommand += " -EndTime `"$($EndTime.ToString('g'))`""
                    }

                    $escapedInternal = $InternalMessage -replace '"', '""'
                    $psCommand += " -InternalMessage `"$escapedInternal`""

                    if ($ExternalMessage) {
                        $escapedExternal = $ExternalMessage -replace '"', '""'
                        $psCommand += " -ExternalMessage `"$escapedExternal`""
                        if ($ExternalAudienceAll) {
                            $psCommand += " -ExternalAudience All"
                        } else {
                            $psCommand += " -ExternalAudience Known"
                        }
                    }

                    Write-Host "`nCommand to enable:" -ForegroundColor Yellow
                    Write-Host "  $psCommand" -ForegroundColor Cyan
                }
            }

            # Show setup tip if Exchange module not connected
            if (-not $exchangeConnected) {
                Write-Host "`n--- Setup ---" -ForegroundColor DarkGray
                Write-Host "Out of Office is a server-side Exchange feature." -ForegroundColor DarkGray
                Write-Host "Outlook COM cannot access these settings directly." -ForegroundColor DarkGray
                Write-Host "`nTo run commands directly, connect Exchange Online:" -ForegroundColor DarkGray
                Write-Host "  Install-Module ExchangeOnlineManagement" -ForegroundColor DarkGray
                Write-Host "  Connect-ExchangeOnline -UserPrincipalName $smtpAddress" -ForegroundColor DarkGray
                Write-Host "`nOr use Outlook: File > Automatic Replies" -ForegroundColor DarkGray
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nNote: Some features require Exchange or Office 365." -ForegroundColor Gray
}
finally {
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
