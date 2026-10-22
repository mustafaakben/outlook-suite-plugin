param(
    [Parameter(Mandatory = $true)]
    [string]$Name,
    [string]$Mailbox = "",
    [switch]$SendRule,

    # Conditions (several values: separate with ";", any one matches)
    [Alias("From")][string]$FromAddress = "",
    [string]$SubjectContains = "",
    [string]$BodyContains = "",
    [string]$BodyOrSubjectContains = "",
    [string]$HeaderContains = "",
    [string]$SenderAddressContains = "",
    [string]$RecipientAddressContains = "",
    [string]$SentTo = "",
    [switch]$HasAttachment,
    [ValidateSet("", "Low", "Normal", "High")][string]$Importance = "",
    [string]$Category = "",
    [switch]$OnlyToMe,
    [switch]$ToMe,
    [switch]$CcMe,
    [switch]$ToOrCcMe,
    [switch]$NotToMe,
    [switch]$MeetingInviteOrUpdate,
    [string]$Account = "",

    # Exceptions
    [string]$ExceptFrom = "",
    [string]$ExceptSubjectContains = "",
    [string]$ExceptBodyContains = "",
    [string]$ExceptSenderAddressContains = "",
    [switch]$ExceptHasAttachment,
    [switch]$ExceptOnlyToMe,
    [ValidateSet("", "Low", "Normal", "High")][string]$ExceptImportance = "",
    [string]$ExceptCategory = "",
    [switch]$ExceptMeetingInviteOrUpdate,

    # Actions
    [string]$MoveToFolder = "",
    [string]$CopyToFolder = "",
    [string]$ForwardTo = "",
    [string]$ForwardAsAttachmentTo = "",
    [string]$RedirectTo = "",
    [string]$AssignCategory = "",
    [switch]$ClearCategories,
    [switch]$Delete,
    [switch]$DeletePermanently,
    [ValidateSet("", "Today", "Tomorrow", "ThisWeek", "NextWeek", "NoDate")][string]$MarkAsTask = "",
    [string]$FlagText = "Follow up",
    [string]$NewItemAlert = "",
    [switch]$DesktopAlert,
    [string]$PlaySound = "",
    [switch]$StopProcessing,
    # Send-rule actions
    [string]$CcTo = "",
    [switch]$NotifyRead,
    [switch]$NotifyDelivery,

    [switch]$Disabled,
    [switch]$Force,
    [switch]$RunNow,
    [string]$RunFolder = "Inbox",
    [switch]$Confirm,
    [switch]$Json,
    [bool]$StripLinks = $true   # legacy, ignored
)

# Outlook Rules Create: make a rule with any condition, exception and action the Outlook object
# model supports (conditions like size/date range and actions like mark read, print or run script
# exist only in the Rules Wizard; they can be listed and switched on/off, not created).
# Values with ";" mean "any of these". Gated: previews unless -Confirm.
# -Force replaces a rule with the same name. -RunNow runs it on -RunFolder after saving.
# -Disabled creates it switched off. -SendRule makes a rule for messages you send.
#
# Usage: .\outlook-rules-create.ps1 -Name "Newsletters" -SenderAddressContains "news@;digest@" -MoveToFolder "Inbox/News" -Confirm
# Usage: .\outlook-rules-create.ps1 -Name "Boss" -FromAddress "boss@x.com" -AssignCategory "!Urgent" -ExceptSubjectContains "FYI" -Confirm
# Usage: .\outlook-rules-create.ps1 -Name "Cc archive" -SendRule -SubjectContains "[client]" -CcTo "archive@x.com" -Confirm

. (Join-Path $PSScriptRoot "_common.ps1")

# Always returns a string array (the comma stops PowerShell from unrolling a one-item array to
# a plain string; Outlook's Text/Categories properties need an array).
function Split-Values([string]$s) { return , [string[]]@($s -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

try {
    $null = Connect-Outlook
    $store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
    $rules = Get-StoreRules $store

    $existing = $null
    for ($i = 1; $i -le $rules.Count; $i++) { if ($rules.Item($i).Name -ieq $Name) { $existing = $rules.Item($i) } }
    if ($existing -and -not $Force) { Stop-WithError "A rule named '$Name' already exists in $($store.DisplayName)." "Use -Force to replace it, or outlook-rules-manage.ps1 to change it." }

    # Folders for move/copy must be in the rule's mailbox.
    function Resolve-RuleFolder([string]$p) {
        $f = $null
        try { $f = Resolve-OutlookFolder -Namespace $Namespace -Path "$($store.DisplayName)/$p" } catch { $f = Resolve-OutlookFolder -Namespace $Namespace -Path $p }
        if ($f.StoreID -ne $store.StoreID) { throw "Folder '$p' is not in mailbox '$($store.DisplayName)'; rule folders must be in the same mailbox." }
        return $f
    }
    $moveFolder = if ($MoveToFolder) { Resolve-RuleFolder $MoveToFolder } else { $null }
    $copyFolder = if ($CopyToFolder) { Resolve-RuleFolder $CopyToFolder } else { $null }

    $hasCondition = $FromAddress -or $SubjectContains -or $BodyContains -or $BodyOrSubjectContains -or $HeaderContains -or $SenderAddressContains -or $RecipientAddressContains -or $SentTo -or $HasAttachment -or $Importance -or $Category -or $OnlyToMe -or $ToMe -or $CcMe -or $ToOrCcMe -or $NotToMe -or $MeetingInviteOrUpdate -or $Account
    $hasAction = $MoveToFolder -or $CopyToFolder -or $ForwardTo -or $ForwardAsAttachmentTo -or $RedirectTo -or $AssignCategory -or $ClearCategories -or $Delete -or $DeletePermanently -or $MarkAsTask -or $NewItemAlert -or $DesktopAlert -or $PlaySound -or $StopProcessing -or $CcTo -or $NotifyRead -or $NotifyDelivery
    if (-not $hasCondition) { Stop-WithError "Give at least one condition (for example -FromAddress, -SubjectContains, -HasAttachment)." }
    if (-not $hasAction) { Stop-WithError "Give at least one action (for example -MoveToFolder, -AssignCategory, -ForwardTo)." }

    # Plain-language summary for the preview.
    $summary = @()
    foreach ($p in $PSBoundParameters.Keys) {
        if ($p -in "Name", "Mailbox", "Confirm", "Json", "Force", "StripLinks", "RunFolder", "FlagText") { continue }
        $v = $PSBoundParameters[$p]
        $summary += $(if ($v -is [System.Management.Automation.SwitchParameter]) { $p } else { "${p}: $v" })
    }
    if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "rules"))) {
        Write-Host "=== NEW RULE (PREVIEW) ===" -ForegroundColor Yellow
        Write-Host "  Mailbox: $($store.DisplayName)   Name: $Name   $(if ($SendRule) { 'send rule' } else { 'receive rule' })$(if ($Disabled) { ', created OFF' })"
        foreach ($s in $summary) { Write-Host "  $s" -ForegroundColor Cyan }
        if ($existing) { Write-Host "  Replaces the existing rule with this name." -ForegroundColor Yellow }
        Approve-Action -Action "rules" -Confirmed $false -What "creating the rule" | Out-Null
        exit 0
    }

    if ($existing) { $rules.Remove($existing.Name) }
    $rule = $rules.Create($Name, $(if ($SendRule) { 1 } else { 0 }))
    $c = $rule.Conditions; $x = $rule.Exceptions; $a = $rule.Actions

    function Add-Recipients($cond, [string]$list) {
        $cond.Enabled = $true
        foreach ($v in (Split-Values $list)) { [void]$cond.Recipients.Add($v) }
        if (-not $cond.Recipients.ResolveAll()) {
            $bad = @(foreach ($r in $cond.Recipients) { if (-not $r.Resolved) { $r.Name } })
            throw "Could not resolve: $($bad -join '; '). Nothing was saved."
        }
    }
    function Set-Text($cond, [string]$list) { $cond.Enabled = $true; $cond.Text = Split-Values $list }
    $imp = @{ "Low" = 0; "Normal" = 1; "High" = 2 }

    if ($FromAddress) { Add-Recipients $c.From $FromAddress }
    if ($SentTo) { Add-Recipients $c.SentTo $SentTo }
    if ($SubjectContains) { Set-Text $c.Subject $SubjectContains }
    if ($BodyContains) { Set-Text $c.Body $BodyContains }
    if ($BodyOrSubjectContains) { Set-Text $c.BodyOrSubject $BodyOrSubjectContains }
    if ($HeaderContains) { Set-Text $c.MessageHeader $HeaderContains }
    if ($SenderAddressContains) { Set-Text $c.SenderAddress $SenderAddressContains }
    if ($RecipientAddressContains) { Set-Text $c.RecipientAddress $RecipientAddressContains }
    if ($HasAttachment) { $c.HasAttachment.Enabled = $true }
    if ($Importance) { $c.Importance.Enabled = $true; $c.Importance.Importance = $imp[$Importance] }
    if ($Category) { $c.Category.Enabled = $true; $c.Category.Categories = Split-Values $Category }
    if ($OnlyToMe) { $c.OnlyToMe.Enabled = $true }
    if ($ToMe) { $c.ToMe.Enabled = $true }
    if ($CcMe) { $c.CC.Enabled = $true }
    if ($ToOrCcMe) { $c.ToOrCc.Enabled = $true }
    if ($NotToMe) { $c.NotTo.Enabled = $true }
    if ($MeetingInviteOrUpdate) { $c.MeetingInviteOrUpdate.Enabled = $true }
    if ($Account) {
        $acct = $null; foreach ($ac in $Namespace.Accounts) { if ($ac.SmtpAddress -ieq $Account -or $ac.DisplayName -ieq $Account) { $acct = $ac } }
        if (-not $acct) { throw "Account '$Account' not found (outlook-accounts-list.ps1)." }
        $c.Account.Enabled = $true; $c.Account.Account = $acct
    }

    if ($ExceptFrom) { Add-Recipients $x.From $ExceptFrom }
    if ($ExceptSubjectContains) { Set-Text $x.Subject $ExceptSubjectContains }
    if ($ExceptBodyContains) { Set-Text $x.Body $ExceptBodyContains }
    if ($ExceptSenderAddressContains) { Set-Text $x.SenderAddress $ExceptSenderAddressContains }
    if ($ExceptHasAttachment) { $x.HasAttachment.Enabled = $true }
    if ($ExceptOnlyToMe) { $x.OnlyToMe.Enabled = $true }
    if ($ExceptImportance) { $x.Importance.Enabled = $true; $x.Importance.Importance = $imp[$ExceptImportance] }
    if ($ExceptCategory) { $x.Category.Enabled = $true; $x.Category.Categories = Split-Values $ExceptCategory }
    if ($ExceptMeetingInviteOrUpdate) { $x.MeetingInviteOrUpdate.Enabled = $true }

    if ($moveFolder) { $a.MoveToFolder.Enabled = $true; $a.MoveToFolder.Folder = $moveFolder }
    if ($copyFolder) { $a.CopyToFolder.Enabled = $true; $a.CopyToFolder.Folder = $copyFolder }
    if ($ForwardTo) { Add-Recipients $a.Forward $ForwardTo }
    if ($ForwardAsAttachmentTo) { Add-Recipients $a.ForwardAsAttachment $ForwardAsAttachmentTo }
    if ($RedirectTo) { Add-Recipients $a.Redirect $RedirectTo }
    if ($AssignCategory) { $a.AssignToCategory.Enabled = $true; $a.AssignToCategory.Categories = Split-Values $AssignCategory }
    if ($ClearCategories) { $a.ClearCategories.Enabled = $true }
    if ($Delete) { $a.Delete.Enabled = $true }
    if ($DeletePermanently) { $a.DeletePermanently.Enabled = $true }
    if ($MarkAsTask) {
        $a.MarkAsTask.Enabled = $true
        $a.MarkAsTask.MarkInterval = @{ "Today" = 0; "Tomorrow" = 1; "ThisWeek" = 2; "NextWeek" = 3; "NoDate" = 4 }[$MarkAsTask]
        $a.MarkAsTask.FlagTo = $FlagText
    }
    if ($NewItemAlert) { $a.NewItemAlert.Enabled = $true; $a.NewItemAlert.Text = $NewItemAlert }
    if ($DesktopAlert) { $a.DesktopAlert.Enabled = $true }
    if ($PlaySound) {
        if (-not (Test-Path -LiteralPath $PlaySound)) { throw "Sound file not found: $PlaySound" }
        $a.PlaySound.Enabled = $true; $a.PlaySound.FilePath = (Resolve-Path -LiteralPath $PlaySound).Path
    }
    if ($StopProcessing) { $a.Stop.Enabled = $true }
    if ($CcTo) { Add-Recipients $a.CC $CcTo }
    if ($NotifyRead) { $a.NotifyRead.Enabled = $true }
    if ($NotifyDelivery) { $a.NotifyDelivery.Enabled = $true }

    $rule.Enabled = -not $Disabled
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try { $rules.Save($false) }
    catch {
        if ($_.Exception.Message -match "invalid actions or conditions") {
            Stop-WithError "Outlook refused to save the rule: one of its conditions or actions is not allowed in this mailbox. Nothing was saved." "Common causes: moving/copying into a folder inside Deleted Items, a folder in another mailbox, or an action this account type does not support. Try fewer options to find which one."
        }
        throw
    }
    $saveMs = $sw.ElapsedMilliseconds

    $ran = ""
    if ($RunNow) {
        $target = Resolve-OutlookFolder -Namespace $Namespace -Path $RunFolder
        $rule.Execute($false, $target, $false, 0)
        $ran = $target.FolderPath
    }

    $saved = ConvertTo-RuleRecord (Find-Rule -Rules (Get-StoreRules $store) -Name $Name)
    if ($Json) { Write-Json ([ordered]@{ Result = $(if ($existing) { "replaced" } else { "created" }); Mailbox = $store.DisplayName; Rule = $saved; SaveMs = $saveMs; RanOn = $ran }); exit 0 }
    Write-Host "=== RULE $(if ($existing) { 'REPLACED' } else { 'CREATED' }) ===" -ForegroundColor Green
    Write-Host "  Mailbox: $($store.DisplayName)   (saved in $saveMs ms)"
    Write-RuleRecord $saved -Detailed
    if ($ran) { Write-Host "  Ran now on: $ran" -ForegroundColor Cyan }
}
catch {
    Stop-WithError $_
}
