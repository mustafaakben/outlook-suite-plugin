param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [int]$Days = 7,
    [int]$Limit = 20
)

# Outlook Conversation Thread Script
# Usage: .\outlook-conversation.ps1 -EntryID "00000000..." (preferred)
# Usage: .\outlook-conversation.ps1 -Index 1 (fallback)

$Outlook = $null
$Namespace = $null
$Inbox = $null

try {
    # Connect to Outlook - try active instance first
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
        Write-Host "Usage: .\outlook-conversation.ps1 -EntryID ""00000000...""" -ForegroundColor Gray
    }

    if ($targetEmail) {
        # Get sender email with Exchange fallback for the target email
        $targetSenderAddr = $targetEmail.SenderEmailAddress
        if ($targetSenderAddr -match "^/O=") { $targetSenderAddr = $targetEmail.SenderName }

        Write-Host "`n=== CONVERSATION THREAD ===" -ForegroundColor Cyan
        Write-Host "Subject: $($targetEmail.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($targetEmail.SenderName) <$targetSenderAddr>"
        Write-Host "Date: $($targetEmail.ReceivedTime.ToString('g'))"
        Write-Host ""

        # Try GetConversation first (works with Exchange/IMAP accounts)
        $conversationDone = $false
        try {
            $conversation = $targetEmail.GetConversation()
            if ($conversation) {
                $conversationTable = $conversation.GetTable()
                $threadCount = 0

                while (-not $conversationTable.EndOfTable) {
                    $row = $conversationTable.GetNextRow()
                    $threadCount++

                    if ($threadCount -gt $Limit) {
                        Write-Host "... (limited to $Limit messages, use -Limit to increase)" -ForegroundColor Gray
                        break
                    }

                    $rowSubject = $row["Subject"]
                    $rowSender = $row["SenderName"]
                    $rowReceived = $row["ReceivedTime"]
                    $rowDate = if ($rowReceived) { $rowReceived.ToString("g") } else { "N/A" }

                    Write-Host "$threadCount. $rowSubject" -ForegroundColor Yellow
                    Write-Host "      From: $rowSender"
                    Write-Host "      Date: $rowDate"
                    Write-Host ""
                }

                if ($threadCount -eq 0) {
                    Write-Host "No messages found in conversation." -ForegroundColor Gray
                } else {
                    Write-Host "--- Total: $threadCount message(s) in thread ---" -ForegroundColor Cyan
                }

                Write-Host ""
                Write-Host "Conversation Topic: $($conversation.ConversationTopic)" -ForegroundColor Gray
                $conversationDone = $true
            }
        } catch {
            # GetConversation or GetTable failed - fall through to subject-based search
        }

        if (-not $conversationDone) {
            # Fallback: Search by subject with date filter
            Write-Host "Note: Full conversation lookup not available. Searching by subject..." -ForegroundColor Yellow
            Write-Host ""

            if (-not $Inbox) {
                try {
                    $Inbox = $Namespace.GetDefaultFolder(6)
                } catch {
                    Write-Host "Unable to access Inbox for fallback search." -ForegroundColor Red
                }
            }

            if (-not $Inbox) {
                Write-Host "Cannot run subject-based fallback without Inbox access." -ForegroundColor Red
                return
            }

            $baseSubject = $targetEmail.Subject -replace "^(RE:|FW:|FWD:|Re:|Fw:|Fwd:)\s*", ""

            # Use date-filtered items instead of scanning entire mailbox
            $fallbackSince = (Get-Date).AddDays(-$Days).ToString("g")
            $fallbackFilter = "[ReceivedTime] >= '$fallbackSince'"
            $fallbackItems = $Inbox.Items.Restrict($fallbackFilter)
            $fallbackItems.Sort("[ReceivedTime]", $false)

            $threadCount = 0
            foreach ($item in $fallbackItems) {
                $itemSubject = $item.Subject -replace "^(RE:|FW:|FWD:|Re:|Fw:|Fwd:)\s*", ""

                if ($itemSubject -eq $baseSubject) {
                    $threadCount++

                    $itemSenderAddr = $item.SenderEmailAddress
                    if ($itemSenderAddr -match "^/O=") { $itemSenderAddr = $item.SenderName }

                    Write-Host "$threadCount. $($item.Subject)" -ForegroundColor Yellow
                    Write-Host "      From: $($item.SenderName) <$itemSenderAddr>"
                    Write-Host "      Date: $($item.ReceivedTime.ToString('g'))"
                    Write-Host ""

                    if ($threadCount -ge $Limit) {
                        Write-Host "... (limited to $Limit messages, use -Limit to increase)" -ForegroundColor Gray
                        break
                    }
                }
            }

            if ($threadCount -eq 0) {
                Write-Host "No related messages found in last $Days days." -ForegroundColor Gray
                Write-Host "Tip: Try increasing -Days to search further back." -ForegroundColor Gray
            } else {
                Write-Host "--- Found: $threadCount related message(s) ---" -ForegroundColor Cyan
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
