param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [string]$Path = "",
    [int]$Days = 7,
    [int]$AttachmentIndex = 0,
    [switch]$ListOnly,
    [switch]$IncludeInline,
    [switch]$Json
)

# Outlook Attachment Save: list or save an email's attachments.
# Inline images (signature logos, pasted pictures) are marked [inline] and skipped unless
# -IncludeInline. Existing files are never overwritten: a numbered copy is saved instead.
# Usage: .\outlook-attachment-save.ps1 -EntryID "00000000..." -ListOnly
# Usage: .\outlook-attachment-save.ps1 -EntryID "00000000..." -Path "C:\Downloads"
# Usage: .\outlook-attachment-save.ps1 -EntryID "00000000..." -AttachmentIndex 2

. (Join-Path $PSScriptRoot "_common.ps1")
if (-not $Path) { $Path = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads" }

$PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
$PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $html = ""; try { $html = $email.HTMLBody } catch { }

    $list = @(for ($i = 1; $i -le $email.Attachments.Count; $i++) {
        $att = $email.Attachments.Item($i)
        $cid = $null; $hidden = $false
        try { $cid = $att.PropertyAccessor.GetProperty($PR_CONTENT_ID) } catch { }
        try { $hidden = [bool]$att.PropertyAccessor.GetProperty($PR_HIDDEN) } catch { }
        $inline = $hidden -or ($cid -and $html -match [regex]::Escape("cid:$cid"))
        [ordered]@{ Index = $i; FileName = $att.FileName; SizeKB = [math]::Round($att.Size / 1KB, 1); Inline = [bool]$inline; Saved = $null }
    })

    if ($AttachmentIndex -gt $list.Count) {
        Stop-WithError "Attachment index $AttachmentIndex is out of range; the email has $($list.Count) attachment(s)."
    }

    if (-not $ListOnly -and $list.Count -gt 0) {
        if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        foreach ($a in $list) {
            if ($AttachmentIndex -gt 0 -and $a.Index -ne $AttachmentIndex) { continue }
            if ($a.Inline -and -not $IncludeInline -and $AttachmentIndex -eq 0) { continue }
            $name = $a.FileName -replace '[\\/:*?"<>|]', '_'
            $file = Join-Path $Path $name
            $base = [System.IO.Path]::GetFileNameWithoutExtension($name); $ext = [System.IO.Path]::GetExtension($name); $c = 1
            while (Test-Path -LiteralPath $file) { $file = Join-Path $Path "${base}_$c$ext"; $c++ }
            $email.Attachments.Item($a.Index).SaveAsFile($file)
            $a.Saved = $file
        }
    }

    $rec = ConvertTo-MailRecord $email
    if ($Json) { Write-Json ([ordered]@{ EntryID = $email.EntryID; Subject = $email.Subject; Folder = $Path; Attachments = $list }); exit 0 }

    Write-Host "=== ATTACHMENTS ===" -ForegroundColor Cyan
    Write-Host "EntryID: $($rec.EntryID)"
    Write-Host "Subject: $($rec.Subject)" -ForegroundColor Yellow
    Write-Host "From: $($rec.From)"
    Write-Host "Date: $(Format-MailDate $rec.Date)"
    if ($list.Count -eq 0) { Write-Host "This email has no attachments." -ForegroundColor Gray; exit 0 }
    $regular = @($list | Where-Object { -not $_.Inline }).Count
    Write-Host "Attachments: $regular file(s)$(if ($list.Count -gt $regular) { " + $($list.Count - $regular) inline image(s)" })"
    foreach ($a in $list) {
        $mark = if ($a.Inline) { " [inline]" } else { "" }
        $saved = if ($a.Saved) { "  -> saved: $($a.Saved)" } else { "" }
        Write-Host "  [$($a.Index)] $($a.FileName) ($($a.SizeKB) KB)$mark$saved"
    }
    if ($ListOnly) {
        Write-Host "Save one: -AttachmentIndex <n>. Save all: drop -ListOnly (add -IncludeInline for inline images)." -ForegroundColor DarkGray
    } else {
        $n = @($list | Where-Object { $_.Saved }).Count
        Write-Host "=== Saved $n file(s) to $Path ===" -ForegroundColor Cyan
    }
}
catch {
    Stop-WithError $_
}
