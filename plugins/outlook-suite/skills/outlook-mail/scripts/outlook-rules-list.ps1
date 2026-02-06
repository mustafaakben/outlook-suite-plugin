param(
    [switch]$Enabled,
    [switch]$Detailed
)

# Outlook Rules List Script
# List email rules
# Usage: .\outlook-rules-list.ps1
# Only enabled rules: .\outlook-rules-list.ps1 -Enabled
# With details: .\outlook-rules-list.ps1 -Detailed

$Outlook = $null
$Namespace = $null
$rules = $null

try {
    # Connect to Outlook - try active instance first
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $Outlook = New-Object -ComObject Outlook.Application
        Start-Sleep -Milliseconds 500
    }

    $Namespace = $Outlook.GetNamespace("MAPI")

    # Get rules
    $rules = $Namespace.DefaultStore.GetRules()

    Write-Host "`n=== OUTLOOK RULES ===" -ForegroundColor Cyan

    if ($Enabled) {
        Write-Host "(Showing only enabled rules)" -ForegroundColor Gray
    }

    Write-Host ""

    $count = 0
    $displayed = 0
    $enabledCount = 0

    foreach ($rule in $rules) {
        $count++

        if ($rule.Enabled) { $enabledCount++ }

        # Filter by enabled status
        if ($Enabled -and -not $rule.Enabled) {
            continue
        }

        $displayed++

        # Status indicator
        $status = if ($rule.Enabled) { "[ON] " } else { "[OFF]" }
        $statusColor = if ($rule.Enabled) { "Green" } else { "Gray" }

        Write-Host "$displayed. " -NoNewline
        Write-Host $status -NoNewline -ForegroundColor $statusColor
        Write-Host " $($rule.Name)" -ForegroundColor Yellow

        if ($Detailed) {
            # Show conditions
            $conditions = $rule.Conditions
            $conditionsList = [System.Collections.ArrayList]::new()

            try {
                if ($conditions.From.Enabled) {
                    $fromAddresses = [System.Collections.ArrayList]::new()
                    foreach ($recipient in $conditions.From.Recipients) {
                        [void]$fromAddresses.Add($recipient.Address)
                    }
                    if ($fromAddresses.Count -gt 0) {
                        [void]$conditionsList.Add("From: $($fromAddresses -join ', ')")
                    }
                }
            } catch {}

            try {
                if ($conditions.Subject.Enabled) {
                    [void]$conditionsList.Add("Subject contains: $($conditions.Subject.Text -join ', ')")
                }
            } catch {}

            try {
                if ($conditions.SentTo.Enabled) {
                    $toAddresses = [System.Collections.ArrayList]::new()
                    foreach ($recipient in $conditions.SentTo.Recipients) {
                        [void]$toAddresses.Add($recipient.Address)
                    }
                    if ($toAddresses.Count -gt 0) {
                        [void]$conditionsList.Add("Sent to: $($toAddresses -join ', ')")
                    }
                }
            } catch {}

            try {
                if ($conditions.HasAttachment.Enabled) {
                    [void]$conditionsList.Add("Has attachment")
                }
            } catch {}

            try {
                if ($conditions.Importance.Enabled) {
                    $imp = switch ($conditions.Importance.Importance) {
                        0 { "Low" }
                        1 { "Normal" }
                        2 { "High" }
                    }
                    [void]$conditionsList.Add("Importance: $imp")
                }
            } catch {}

            if ($conditionsList.Count -gt 0) {
                Write-Host "      Conditions:" -ForegroundColor Cyan
                foreach ($cond in $conditionsList) {
                    Write-Host "        - $cond" -ForegroundColor Gray
                }
            }

            # Show actions
            $actions = $rule.Actions
            $actionsList = [System.Collections.ArrayList]::new()

            try {
                if ($actions.MoveToFolder.Enabled) {
                    try {
                        $folderPath = $actions.MoveToFolder.Folder.FolderPath
                        [void]$actionsList.Add("Move to: $folderPath")
                    } catch {
                        [void]$actionsList.Add("Move to: (folder unavailable)")
                    }
                }
            } catch {}

            try {
                if ($actions.CopyToFolder.Enabled) {
                    try {
                        $folderPath = $actions.CopyToFolder.Folder.FolderPath
                        [void]$actionsList.Add("Copy to: $folderPath")
                    } catch {
                        [void]$actionsList.Add("Copy to: (folder unavailable)")
                    }
                }
            } catch {}

            try {
                if ($actions.Delete.Enabled) {
                    [void]$actionsList.Add("Delete")
                }
            } catch {}

            try {
                if ($actions.DeletePermanently.Enabled) {
                    [void]$actionsList.Add("Delete permanently")
                }
            } catch {}

            try {
                if ($actions.Forward.Enabled) {
                    $fwdAddresses = [System.Collections.ArrayList]::new()
                    foreach ($recipient in $actions.Forward.Recipients) {
                        [void]$fwdAddresses.Add($recipient.Address)
                    }
                    if ($fwdAddresses.Count -gt 0) {
                        [void]$actionsList.Add("Forward to: $($fwdAddresses -join ', ')")
                    }
                }
            } catch {}

            try {
                if ($actions.MarkAsTask.Enabled) {
                    [void]$actionsList.Add("Mark as task")
                }
            } catch {}

            try {
                if ($actions.AssignToCategory.Enabled) {
                    $categories = $actions.AssignToCategory.Categories -join ", "
                    [void]$actionsList.Add("Assign category: $categories")
                }
            } catch {}

            try {
                if ($actions.MarkRead.Enabled) {
                    [void]$actionsList.Add("Mark as read")
                }
            } catch {}

            try {
                if ($actions.Stop.Enabled) {
                    [void]$actionsList.Add("Stop processing more rules")
                }
            } catch {}

            if ($actionsList.Count -gt 0) {
                Write-Host "      Actions:" -ForegroundColor Cyan
                foreach ($act in $actionsList) {
                    Write-Host "        - $act" -ForegroundColor Gray
                }
            }

            # Execution order
            Write-Host "      Execution Order: $($rule.ExecutionOrder)" -ForegroundColor DarkGray
        }

        Write-Host ""
    }

    if ($displayed -eq 0) {
        if ($Enabled) {
            Write-Host "No enabled rules found." -ForegroundColor Gray
            Write-Host "Use outlook-rules-create.ps1 to create a new rule." -ForegroundColor Gray
        } else {
            Write-Host "No rules found." -ForegroundColor Gray
            Write-Host "Use outlook-rules-create.ps1 to create a new rule." -ForegroundColor Gray
        }
    } else {
        Write-Host "--- $displayed rule(s) shown ($enabledCount enabled of $count total) ---" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Note: Rules access requires Outlook to be fully initialized." -ForegroundColor Gray
}
finally {
    if ($rules) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null } catch {}
    }
    if ($Namespace) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
