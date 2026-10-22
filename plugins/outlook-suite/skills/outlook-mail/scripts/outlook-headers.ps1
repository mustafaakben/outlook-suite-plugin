param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [switch]$Raw,
    [switch]$Json
)

# Outlook Headers: the internet headers of a received email, with a short summary for checking
# where it came from and whether it is genuine: sender vs Return-Path / Reply-To, SPF / DKIM /
# DMARC results, the relay path (Received hops), Message-ID and threading headers, and
# List-Unsubscribe. -Raw prints the full header block.
# Usage: .\outlook-headers.ps1 -EntryID "00000000..."
# Usage: .\outlook-headers.ps1 -EntryID "00000000..." -Raw

. (Join-Path $PSScriptRoot "_common.ps1")
$PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"

function Get-Header([string]$all, [string]$name) {
    $vals = @()
    foreach ($m in [regex]::Matches($all, "(?im)^$([regex]::Escape($name)):[ \t]*(.*(?:\r?\n[ \t]+.*)*)")) {
        $vals += (($m.Groups[1].Value -replace '\r?\n[ \t]+', ' ').Trim())
    }
    return $vals   # callers wrap in @(); a unary comma here would nest the array
}

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    $h = ""
    try { $h = [string]$email.PropertyAccessor.GetProperty($PR_TRANSPORT_MESSAGE_HEADERS) } catch { }
    if (-not $h) { Stop-WithError "This item has no internet headers (drafts, sent items and some internal Exchange messages have none)." }

    if ($Raw) {
        if ($Json) { Write-Json ([ordered]@{ EntryID = $email.EntryID; Subject = $email.Subject; Headers = $h }); exit 0 }
        Write-Host "=== HEADERS: $($email.Subject) ===" -ForegroundColor Cyan
        Write-Host $h
        exit 0
    }

    $auth = (@(Get-Header $h "Authentication-Results") + @(Get-Header $h "ARC-Authentication-Results")) -join " ; "
    $pick = { param($k) if ($auth -match "(?i)\b$k=(\w+)") { $Matches[1] } else { "" } }
    $received = @(Get-Header $h "Received")
    $hops = @(foreach ($r in $received) { if ($r -match '(?i)from\s+(\S+)') { $Matches[1] } })
    [array]::Reverse($hops)
    $o = [ordered]@{
        EntryID       = $email.EntryID
        Subject       = $email.Subject
        From          = (@(Get-Header $h "From") -join "; ")
        ReplyTo       = (@(Get-Header $h "Reply-To") -join "; ")
        ReturnPath    = (@(Get-Header $h "Return-Path") -join "; ")
        Sender        = (@(Get-Header $h "Sender") -join "; ")
        SPF           = (& $pick "spf")
        DKIM          = (& $pick "dkim")
        DMARC         = (& $pick "dmarc")
        CompAuth      = (& $pick "compauth")
        Hops          = $received.Count
        Path          = ($hops -join " -> ")
        MessageID     = (@(Get-Header $h "Message-ID") -join "")
        InReplyTo     = (@(Get-Header $h "In-Reply-To") -join "")
        References    = (@(Get-Header $h "References") -join " ")
        ListUnsubscribe = (@(Get-Header $h "List-Unsubscribe") -join " ")
        Date          = (@(Get-Header $h "Date") -join "")
        XMailer       = (@(Get-Header $h "X-Mailer") -join "")
    }
    $warnings = @()
    $fromDomain = if ($o.From -match '@([\w.-]+)') { $Matches[1].ToLower() } else { "" }
    $retDomain = if ($o.ReturnPath -match '@([\w.-]+)') { $Matches[1].ToLower() } else { "" }
    $replyDomain = if ($o.ReplyTo -match '@([\w.-]+)') { $Matches[1].ToLower() } else { "" }
    if ($o.SPF -and $o.SPF -notmatch '^(pass|none)$') { $warnings += "SPF $($o.SPF)" }
    if ($o.DKIM -and $o.DKIM -notmatch '^(pass|none)$') { $warnings += "DKIM $($o.DKIM)" }
    if ($o.DMARC -and $o.DMARC -notmatch '^(pass|bestguesspass|none)$') { $warnings += "DMARC $($o.DMARC)" }
    if ($replyDomain -and $fromDomain -and $replyDomain -ne $fromDomain) { $warnings += "Reply-To domain ($replyDomain) differs from From ($fromDomain)" }
    if ($retDomain -and $fromDomain -and -not ($retDomain.EndsWith($fromDomain) -or $fromDomain.EndsWith($retDomain))) { $warnings += "Return-Path domain ($retDomain) differs from From ($fromDomain) (normal for mailing services)" }
    $o.Warnings = $warnings

    if ($Json) { Write-Json $o; exit 0 }
    Write-Host "=== HEADERS SUMMARY ===" -ForegroundColor Cyan
    foreach ($k in $o.Keys) {
        if ($k -eq "Warnings") { continue }
        $v = $o[$k]; if ("$v") { Write-Host ("  {0,-16} {1}" -f "${k}:", $v) }
    }
    if ($warnings.Count) { foreach ($w in $warnings) { Write-Host "  WARNING: $w" -ForegroundColor Yellow } }
    else { Write-Host "  No authentication warnings." -ForegroundColor Green }
    Write-Host "Full headers: -Raw" -ForegroundColor DarkGray
}
catch {
    Stop-WithError $_
}
