param(
    [ValidateSet("List", "Show", "Create", "Delete", "SetDefault", "Insert")]
    [string]$Action = "List",
    [string]$Name = "",
    [string]$Html = "",
    [string]$HtmlFile = "",
    [string]$Text = "",
    [ValidateSet("New", "Reply", "Both", "None")]
    [string]$For = "Both",
    [string]$Account = "",
    [string]$EntryID = "",
    [switch]$Force,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Signatures (classic Outlook keeps them as files in %APPDATA%\Microsoft\Signatures;
# the object model has no signature API).
#   List                         names, formats, and each account's default new/reply signature
#   Show -Name X                 the signature's text (links as [URL-001] references)
#   Create -Name X -Html|-HtmlFile|-Text   new signature (-Force replaces one: gated)
#   Delete -Name X               remove its files (gated)
#   SetDefault -Name X -For New|Reply|Both [-Account a@x]   default for an account (gated);
#                                -For None clears. With Microsoft 365 "roaming signatures",
#                                Outlook may keep using the cloud setting instead.
#   Insert -Name X -EntryID <draft>  add the signature to the end of your text in a draft
#                                (reply/forward drafts: inside your block, above the quoted thread)
# Usage: .\outlook-signatures.ps1
# Usage: .\outlook-signatures.ps1 -Action Insert -Name "Work" -EntryID "<draft EntryID>"

. (Join-Path $PSScriptRoot "_common.ps1")

$dir = Join-Path $env:APPDATA "Microsoft\Signatures"
$profilesKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

function Get-SignatureNames {
    if (-not (Test-Path $dir)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.Extension -in ".htm", ".txt", ".rtf" } | ForEach-Object { $_.BaseName } | Sort-Object -Unique)
}
function Get-SignatureFile([string]$n, [string]$ext) { return (Join-Path $dir "$n$ext") }
function Convert-RegString($v) {
    if ($null -eq $v) { return "" }
    if ($v -is [byte[]]) { return ([System.Text.Encoding]::Unicode.GetString($v)).TrimEnd([char]0) }
    return [string]$v
}
# Account keys in the default profile that carry signature settings.
function Get-AccountSignatureKeys {
    if (-not (Test-Path $profilesKey)) { return @() }
    $defaultProfile = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name DefaultProfile -ErrorAction SilentlyContinue).DefaultProfile
    $prof = if ($defaultProfile) { Join-Path $profilesKey $defaultProfile } else { (Get-ChildItem $profilesKey | Select-Object -First 1).PSPath }
    return @(Get-ChildItem -Path (Join-Path $prof "9375CFF0413111d3B88A00104B2A6676") -ErrorAction SilentlyContinue | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath
        $acct = Convert-RegString $p."Account Name"
        if ($acct) { [pscustomobject]@{ Key = $_.PSPath; Account = $acct; New = (Convert-RegString $p."New Signature"); Reply = (Convert-RegString $p."Reply-Forward Signature") } }
    })
}

try {
    $names = Get-SignatureNames
    $needName = $Action -in "Show", "Create", "Delete", "SetDefault", "Insert"
    if ($needName -and -not $Name -and -not ($Action -eq "SetDefault" -and $For -eq "None")) { Stop-WithError "Give the signature -Name." }
    if ($Action -in "Show", "Delete", "Insert" -and $names -notcontains $Name) { Stop-WithError "No signature named '$Name'. Signatures: $($names -join ', ')" }

    switch ($Action) {
        "List" {
            $accounts = Get-AccountSignatureKeys
            $list = @(foreach ($n in $names) {
                [ordered]@{ Name = $n; Formats = @(foreach ($e in ".htm", ".txt", ".rtf") { if (Test-Path -LiteralPath (Get-SignatureFile $n $e)) { $e.TrimStart('.') } }) }
            })
            $o = [ordered]@{ Folder = $dir; Signatures = $list; Defaults = @($accounts | ForEach-Object { [ordered]@{ Account = $_.Account; New = $_.New; Reply = $_.Reply } }) }
            if ($Json) { Write-Json $o; exit 0 }
            Write-Host "=== SIGNATURES ($($list.Count)) ===" -ForegroundColor Cyan
            foreach ($s in $list) { Write-Host "  $($s.Name)  [$($s.Formats -join ', ')]" }
            if ($accounts) {
                Write-Host "Defaults per account:" -ForegroundColor Gray
                foreach ($a in $accounts) { Write-Host "  $($a.Account): new = $(if ($a.New) { $a.New } else { '(none)' }); reply/forward = $(if ($a.Reply) { $a.Reply } else { '(none)' })" }
            }
            exit 0
        }
        "Show" {
            $htmPath = Get-SignatureFile $Name ".htm"; $txtPath = Get-SignatureFile $Name ".txt"
            $html = if (Test-Path -LiteralPath $htmPath) { Get-Content -LiteralPath $htmPath -Raw } else { "" }
            $txt = if (Test-Path -LiteralPath $txtPath) { Get-Content -LiteralPath $txtPath -Raw } else { ($html -replace '<[^>]+>', ' ' -replace '\s+', ' ') }
            $res = Format-MailLinks -Text $txt -Html $html
            if ($Json) { Write-Json ([ordered]@{ Name = $Name; Text = $res.Text; Links = $res.Links; HasHtml = [bool]$html }); exit 0 }
            Write-Host "=== SIGNATURE: $Name ===" -ForegroundColor Cyan
            Write-Host $res.Text
            foreach ($l in $res.Links) {
                $label = if ($l.Text) { "  ""$($l.Text)""" } else { "" }
                Write-Host "  $($l.Ref)  $($l.Host)$label" -ForegroundColor DarkGray
            }
            exit 0
        }
        "Create" {
            if (-not ($Html -or $HtmlFile -or $Text)) { Stop-WithError "Give the signature content: -Html, -HtmlFile or -Text." }
            if ($names -contains $Name) {
                if (-not $Force) { Stop-WithError "Signature '$Name' already exists. Add -Force to replace it." }
                if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "signatures"))) {
                    Write-Host "=== REPLACE SIGNATURE '$Name' (PREVIEW) ===" -ForegroundColor Yellow
                    Approve-Action -Action "signatures" -Confirmed $false -What "replacing the signature" | Out-Null
                    exit 0
                }
            }
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            $body = if ($HtmlFile) { Get-Content -LiteralPath $HtmlFile -Raw -Encoding UTF8 } elseif ($Html) { $Html } else { ConvertTo-OutlookHtml $Text }
            $page = if ($body -match '(?i)<html') { $body } else { "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""></head><body>$body</body></html>" }
            $plain = if ($Text) { $Text } else { [System.Net.WebUtility]::HtmlDecode(($body -replace '(?i)<br\s*/?>|</div>|</p>', "`n" -replace '<[^>]+>', '')).Trim() }
            Set-Content -LiteralPath (Get-SignatureFile $Name ".htm") -Value $page -Encoding UTF8
            Set-Content -LiteralPath (Get-SignatureFile $Name ".txt") -Value $plain -Encoding Unicode
            $o = [ordered]@{ Result = $(if ($names -contains $Name) { "replaced" } else { "created" }); Name = $Name; Folder = $dir }
            if ($Json) { Write-Json $o; exit 0 }
            Write-Host "=== SIGNATURE $($o.Result.ToUpper()): $Name ===" -ForegroundColor Green
            exit 0
        }
        "Delete" {
            if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "signatures"))) {
                Write-Host "=== DELETE SIGNATURE '$Name' (PREVIEW) ===" -ForegroundColor Yellow
                Approve-Action -Action "signatures" -Confirmed $false -What "deleting the signature" | Out-Null
                exit 0
            }
            Get-ChildItem -LiteralPath $dir -File | Where-Object { $_.BaseName -eq $Name } | Remove-Item -Force
            $files = Join-Path $dir "${Name}_files"; if (Test-Path -LiteralPath $files) { Remove-Item -LiteralPath $files -Recurse -Force }
            if ($Json) { Write-Json ([ordered]@{ Result = "deleted"; Name = $Name }); exit 0 }
            Write-Host "=== SIGNATURE DELETED: $Name ===" -ForegroundColor Green
            exit 0
        }
        "SetDefault" {
            if ($For -ne "None" -and $names -notcontains $Name) { Stop-WithError "No signature named '$Name'." }
            $accounts = @(Get-AccountSignatureKeys | Where-Object { -not $Account -or $_.Account -ieq $Account })
            if ($accounts.Count -eq 0) { Stop-WithError "No matching account in the Outlook profile$(if ($Account) { " for '$Account'" })." }
            $value = if ($For -eq "None") { "" } else { $Name }
            if (-not ($Confirm -or -not (Test-ConfirmRequired -Action "signatures"))) {
                Write-Host "=== SET DEFAULT SIGNATURE (PREVIEW) ===" -ForegroundColor Yellow
                foreach ($a in $accounts) { Write-Host "  $($a.Account): $(if ($For -in 'New','Both','None') { "new -> $(if ($value) { $value } else { '(none)' })  " })$(if ($For -in 'Reply','Both','None') { "reply/forward -> $(if ($value) { $value } else { '(none)' })" })" }
                Write-Host "  Note: with Microsoft 365 roaming signatures Outlook may keep the cloud setting." -ForegroundColor Gray
                Approve-Action -Action "signatures" -Confirmed $false -What "changing the default signature" | Out-Null
                exit 0
            }
            $bytes = [System.Text.Encoding]::Unicode.GetBytes($value + [char]0)
            foreach ($a in $accounts) {
                if ($For -in "New", "Both", "None") { Set-ItemProperty -Path $a.Key -Name "New Signature" -Value $bytes -Type Binary }
                if ($For -in "Reply", "Both", "None") { Set-ItemProperty -Path $a.Key -Name "Reply-Forward Signature" -Value $bytes -Type Binary }
            }
            if ($Json) { Write-Json ([ordered]@{ Result = "set"; Name = $value; For = $For; Accounts = @($accounts.Account) }); exit 0 }
            Write-Host "=== DEFAULT SIGNATURE SET ($For) for $(@($accounts.Account) -join ', ') ===" -ForegroundColor Green
            Write-Host "  Restart Outlook to apply. With roaming signatures Outlook may keep the cloud setting." -ForegroundColor Gray
            exit 0
        }
        "Insert" {
            if (-not $EntryID) { Stop-WithError "Give the draft's -EntryID." }
            $null = Connect-Outlook
            $mail = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID
            if ($mail.Sent) { Stop-WithError "This item was already sent; signatures can be added to drafts only." }
            $htmPath = Get-SignatureFile $Name ".htm"
            $sigHtml = if (Test-Path -LiteralPath $htmPath) { Get-Content -LiteralPath $htmPath -Raw } else { ConvertTo-OutlookHtml (Get-Content -LiteralPath (Get-SignatureFile $Name ".txt") -Raw) }
            $m = [regex]::Match($sigHtml, '(?is)<body[^>]*>(.*)</body>')
            if ($m.Success) { $sigHtml = $m.Groups[1].Value }
            $block = "<div id=""suite-signature"">$sigHtml</div>"
            if ([int]$mail.BodyFormat -ne 2) {
                $mail.Body = $mail.Body + "`r`n`r`n" + (Get-Content -LiteralPath (Get-SignatureFile $Name ".txt") -Raw)
            } else {
                $h = $mail.HTMLBody
                if ($h.Contains("<!--suite-reply-end-->")) { $h = $h.Replace("</div><!--suite-reply-end-->", "$block</div><!--suite-reply-end-->") }
                else { $b = $h.LastIndexOf("</body>", [StringComparison]::OrdinalIgnoreCase); $h = if ($b -ge 0) { $h.Insert($b, $block) } else { $h + $block } }
                $mail.HTMLBody = $h
            }
            $mail.Save()
            Write-ComposeSummary -Mail $mail -Title "SIGNATURE '$Name' ADDED TO DRAFT" -Json:$Json -Action "signature-insert"
            exit 0
        }
    }
}
catch {
    Stop-WithError $_
}
