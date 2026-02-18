param(
    [Parameter(Mandatory=$true)]
    [string]$Name,

    [string]$FromAddress = "",
    [string]$SubjectContains = "",
    [string]$BodyContains = "",
    [string]$SentTo = "",
    [string]$MoveToFolder = "",
    [string]$CopyToFolder = "",
    [string]$ForwardTo = "",
    [string]$RedirectTo = "",
    [string]$AssignCategory = "",
    [switch]$HasAttachment,
    [switch]$Delete,
    [switch]$DeletePermanently,
    [switch]$StopProcessing,
    [switch]$DesktopAlert,
    [switch]$Disabled,
    [switch]$Force,
    [bool]$StripLinks = $true
)

# Helper: strip URLs from text to reduce context window bloat
function Strip-Links([string]$text) {
    if (-not $text) { return $text }
    return [regex]::Replace($text, 'https?://[^\s<>"''`\)]+', '[URL]')
}

# Outlook Rules Create Script
# Create an email rule with conditions and actions
# Usage: .\outlook-rules-create.ps1 -Name "Archive newsletters" -FromAddress "newsletter@example.com" -MoveToFolder "Archive"
# With category: .\outlook-rules-create.ps1 -Name "Important" -SubjectContains "urgent" -AssignCategory "Important"
# Forward rule: .\outlook-rules-create.ps1 -Name "Forward reports" -SubjectContains "report" -ForwardTo "boss@example.com"
# Attachment rule: .\outlook-rules-create.ps1 -Name "Flag attachments" -HasAttachment -AssignCategory "Has Files"

$Outlook = $null
$Namespace = $null
$rules = $null
$rule = $null

try {
    # Validate: at least one condition required (before COM init)
    $hasCondition = $FromAddress -or $SubjectContains -or $BodyContains -or $SentTo -or $HasAttachment
    if (-not $hasCondition) {
        Write-Host "`nAt least one condition is required." -ForegroundColor Red
        Write-Host "Available conditions: -FromAddress, -SubjectContains, -BodyContains, -SentTo, -HasAttachment" -ForegroundColor Gray
    } else {
        # Validate: at least one action required (before COM init)
        $hasAction = $MoveToFolder -or $CopyToFolder -or $ForwardTo -or $RedirectTo -or $AssignCategory -or $Delete -or $DeletePermanently -or $StopProcessing -or $DesktopAlert
        if (-not $hasAction) {
            Write-Host "`nAt least one action is required." -ForegroundColor Red
            Write-Host "Available actions: -MoveToFolder, -CopyToFolder, -ForwardTo, -RedirectTo, -AssignCategory, -Delete, -DeletePermanently, -StopProcessing, -DesktopAlert" -ForegroundColor Gray
        } else {
            # Connect to Outlook - try active instance first
            try {
                $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            } catch {
                $Outlook = New-Object -ComObject Outlook.Application
                Start-Sleep -Milliseconds 500
            }

            $Namespace = $Outlook.GetNamespace("MAPI")

            # Get rules collection
            $rules = $Namespace.DefaultStore.GetRules()

            # Check if rule name already exists
            $nameExists = $false
            foreach ($existingRule in $rules) {
                if ($existingRule.Name -eq $Name) {
                    $nameExists = $true
                    break
                }
            }

            if ($nameExists -and -not $Force) {
                Write-Host "`nRule '$Name' already exists." -ForegroundColor Yellow
                Write-Host "Note: Outlook allows duplicate rule names." -ForegroundColor Gray
                Write-Host "Use -Force to create anyway, or use a different name." -ForegroundColor Gray
            } else {
                if ($nameExists -and $Force) {
                    Write-Host "`nNote: Creating rule with duplicate name '$Name' (-Force specified)." -ForegroundColor Yellow
                }

                # Create new rule (olRuleReceive = 0 for incoming mail)
                $rule = $rules.Create($Name, 0)

                # Track conditions and actions for display
                $conditionsList = [System.Collections.ArrayList]::new()
                $actionsList = [System.Collections.ArrayList]::new()

                # === SET CONDITIONS ===

                if ($FromAddress) {
                    $rule.Conditions.From.Enabled = $true
                    $rule.Conditions.From.Recipients.Add($FromAddress) | Out-Null
                    if (-not $rule.Conditions.From.Recipients.ResolveAll()) {
                        Write-Host "  Warning: '$FromAddress' could not be resolved in address book. Rule will still match by email address." -ForegroundColor Yellow
                    }
                    [void]$conditionsList.Add("From: $FromAddress")
                }

                if ($SubjectContains) {
                    $rule.Conditions.Subject.Enabled = $true
                    $rule.Conditions.Subject.Text = @($SubjectContains)
                    [void]$conditionsList.Add("Subject contains: $SubjectContains")
                }

                if ($BodyContains) {
                    $rule.Conditions.Body.Enabled = $true
                    $rule.Conditions.Body.Text = @($BodyContains)
                    [void]$conditionsList.Add("Body contains: $BodyContains")
                }

                if ($SentTo) {
                    $rule.Conditions.SentTo.Enabled = $true
                    $rule.Conditions.SentTo.Recipients.Add($SentTo) | Out-Null
                    if (-not $rule.Conditions.SentTo.Recipients.ResolveAll()) {
                        Write-Host "  Warning: '$SentTo' could not be resolved in address book. Rule will still match by email address." -ForegroundColor Yellow
                    }
                    [void]$conditionsList.Add("Sent to: $SentTo")
                }

                if ($HasAttachment) {
                    $rule.Conditions.HasAttachment.Enabled = $true
                    [void]$conditionsList.Add("Has attachment")
                }

                # === SET ACTIONS ===

                # Helper: find folder by name in Inbox subfolders, then store root subfolders
                function Find-OutlookFolder {
                    param([string]$FolderName, $NS)
                    $found = $null
                    # Search Inbox subfolders first
                    $inbox = $NS.GetDefaultFolder(6)
                    try { $found = $inbox.Folders.Item($FolderName) } catch {}
                    if ($found) { return $found }
                    # Search store root subfolders
                    $root = $NS.DefaultStore.GetRootFolder()
                    try { $found = $root.Folders.Item($FolderName) } catch {}
                    if ($found) { return $found }
                    # Walk Inbox subfolders by name comparison
                    foreach ($folder in $inbox.Folders) {
                        if ($folder.Name -eq $FolderName) { return $folder }
                    }
                    return $null
                }

                if ($MoveToFolder) {
                    $targetFolder = Find-OutlookFolder -FolderName $MoveToFolder -NS $Namespace
                    if ($targetFolder) {
                        $rule.Actions.MoveToFolder.Enabled = $true
                        $rule.Actions.MoveToFolder.Folder = $targetFolder
                        [void]$actionsList.Add("Move to: $MoveToFolder")
                    } else {
                        Write-Host "  Warning: Folder '$MoveToFolder' not found. Skipping move action." -ForegroundColor Yellow
                        Write-Host "  Available folders can be listed with outlook-folders.ps1" -ForegroundColor Gray
                    }
                }

                if ($CopyToFolder) {
                    $copyTarget = Find-OutlookFolder -FolderName $CopyToFolder -NS $Namespace
                    if ($copyTarget) {
                        $rule.Actions.CopyToFolder.Enabled = $true
                        $rule.Actions.CopyToFolder.Folder = $copyTarget
                        [void]$actionsList.Add("Copy to: $CopyToFolder")
                    } else {
                        Write-Host "  Warning: Folder '$CopyToFolder' not found. Skipping copy action." -ForegroundColor Yellow
                        Write-Host "  Available folders can be listed with outlook-folders.ps1" -ForegroundColor Gray
                    }
                }

                if ($ForwardTo) {
                    $rule.Actions.Forward.Enabled = $true
                    $rule.Actions.Forward.Recipients.Add($ForwardTo) | Out-Null
                    if (-not $rule.Actions.Forward.Recipients.ResolveAll()) {
                        Write-Host "  Warning: '$ForwardTo' could not be resolved in address book." -ForegroundColor Yellow
                    }
                    [void]$actionsList.Add("Forward to: $ForwardTo")
                }

                if ($RedirectTo) {
                    $rule.Actions.Redirect.Enabled = $true
                    $rule.Actions.Redirect.Recipients.Add($RedirectTo) | Out-Null
                    if (-not $rule.Actions.Redirect.Recipients.ResolveAll()) {
                        Write-Host "  Warning: '$RedirectTo' could not be resolved in address book." -ForegroundColor Yellow
                    }
                    [void]$actionsList.Add("Redirect to: $RedirectTo")
                }

                if ($AssignCategory) {
                    $rule.Actions.AssignToCategory.Enabled = $true
                    $rule.Actions.AssignToCategory.Categories = @($AssignCategory)
                    [void]$actionsList.Add("Assign category: $AssignCategory")
                }

                if ($Delete) {
                    $rule.Actions.Delete.Enabled = $true
                    [void]$actionsList.Add("Delete (move to Deleted Items)")
                }

                if ($DeletePermanently) {
                    $rule.Actions.DeletePermanently.Enabled = $true
                    [void]$actionsList.Add("Delete permanently")
                }

                if ($DesktopAlert) {
                    $rule.Actions.DesktopAlert.Enabled = $true
                    [void]$actionsList.Add("Show desktop alert")
                }

                if ($StopProcessing) {
                    $rule.Actions.Stop.Enabled = $true
                    [void]$actionsList.Add("Stop processing more rules")
                }

                # Check if any action was actually added (folder actions may have been skipped)
                if ($actionsList.Count -eq 0) {
                    Write-Host "`nNo actions could be applied (folder(s) not found)." -ForegroundColor Red
                    Write-Host "Rule was not saved." -ForegroundColor Gray
                } else {
                    # Enable/disable the rule
                    $rule.Enabled = -not $Disabled

                    # Show preview before saving
                    $status = if ($Disabled) { "Disabled" } else { "Enabled" }
                    $statusColor = if ($Disabled) { "Gray" } else { "Green" }

                    Write-Host "`n=== RULE PREVIEW ===" -ForegroundColor Cyan
                    Write-Host "Name: $Name" -ForegroundColor Yellow
                    Write-Host "Status: $status" -ForegroundColor $statusColor

                    Write-Host "`nConditions:" -ForegroundColor Cyan
                    foreach ($cond in $conditionsList) {
                        $condDisplay = if ($StripLinks) { Strip-Links $cond } else { $cond }
                        Write-Host "  - $condDisplay" -ForegroundColor Gray
                    }

                    Write-Host "`nActions:" -ForegroundColor Cyan
                    foreach ($act in $actionsList) {
                        $actDisplay = if ($StripLinks) { Strip-Links $act } else { $act }
                        Write-Host "  - $actDisplay" -ForegroundColor Gray
                    }

                    # Save the rules with error handling
                    try {
                        $rules.Save($true)
                        Write-Host "`n=== RULE CREATED ===" -ForegroundColor Green
                        Write-Host "Rule saved successfully. It will apply to new incoming emails." -ForegroundColor Cyan
                        Write-Host "Use outlook-rules-list.ps1 -Detailed to verify." -ForegroundColor Gray
                    } catch {
                        Write-Host "`nFailed to save rule: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "This may happen if you are offline or disconnected from Exchange." -ForegroundColor Gray
                    }
                }
            }
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nNote: Some rule features may require Exchange or Office 365." -ForegroundColor Gray
}
finally {
    if ($rule) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rule) | Out-Null } catch {}
    }
    if ($rules) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rules) | Out-Null } catch {}
    }
    if ($Namespace) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null } catch {}
    }
    if ($Outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
