param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,
    [string]$StoreID = "",
    [string[]]$Ref = @(),
    [switch]$Json
)

# Outlook Links: full URLs behind the [URL-001] references shown by other scripts.
# Numbering is computed from the message itself, so references from an earlier call
# (outlook-read-body, outlook-conversation, outlook-find previews) resolve here.
# Safe Links wrappers are decoded to the real address.
# Usage: .\outlook-links.ps1 -EntryID "00000000..."                 (all links)
# Usage: .\outlook-links.ps1 -EntryID "00000000..." -Ref URL-003     (one link)

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID
    $html = ""; try { $html = $email.HTMLBody } catch { }
    $links = @((Format-MailLinks -Text $email.Body -Html $html -Mode domain).Links)
    if ($Ref.Count -gt 0) {
        $want = @($Ref | ForEach-Object { $_.Trim('[', ']').ToUpperInvariant() })
        $links = @($links | Where-Object { $want -contains $_.Ref })
        $missing = @($want | Where-Object { ($links.Ref) -notcontains $_ })
        if ($missing.Count -gt 0) { Stop-WithError "No link $($missing -join ', ') in this email (it has $(@((Format-MailLinks -Text $email.Body -Mode domain).Links).Count) links)." }
    }
    if ($Json) { Write-Json ([ordered]@{ EntryID = $email.EntryID; Subject = $email.Subject; Links = $links }); exit 0 }

    Write-Host "=== LINKS ===" -ForegroundColor Cyan
    Write-Host "Subject: $($email.Subject)" -ForegroundColor Yellow
    if ($links.Count -eq 0) { Write-Host "No links in this email." -ForegroundColor Gray; exit 0 }
    foreach ($l in $links) {
        $t = if ($l.Text) { "  ""$($l.Text)""" } else { "" }
        Write-Host "$($l.Ref)$t"
        Write-Host "  $($l.Url)"
    }
}
catch {
    Stop-WithError $_
}
