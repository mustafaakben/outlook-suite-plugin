param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [string]$Path = "",
    [ValidateSet("msg", "msgunicode", "txt", "html", "mht", "rtf", "doc", "template")]
    [string]$Format = "msgunicode",
    [int]$Days = 7,
    [switch]$Overwrite,
    [switch]$Confirm,
    [switch]$Json
)

# Outlook Save As: save an email to a file (.msg, .txt, .html, .mht, .rtf, .doc, .oft).
# Default format msgunicode (Unicode .msg, keeps non-English text). Default folder: Downloads.
# An existing file is not replaced: a numbered name is used, unless -Overwrite
# (irreversible; previews unless -Confirm).
# Usage: .\outlook-save-as.ps1 -EntryID "00000000..."
# Usage: .\outlook-save-as.ps1 -EntryID "00000000..." -Format txt -Path "C:\Temp\mail.txt"

. (Join-Path $PSScriptRoot "_common.ps1")

$formats = @{
    "txt" = @(0, ".txt"); "rtf" = @(1, ".rtf"); "template" = @(2, ".oft"); "msg" = @(3, ".msg")
    "doc" = @(4, ".doc"); "html" = @(5, ".html"); "msgunicode" = @(9, ".msg"); "mht" = @(10, ".mht")
}

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $ext = $formats[$Format][1]
    if (-not $Path -or (Test-Path -LiteralPath $Path -PathType Container)) {
        $dir = if ($Path) { $Path } else { Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads" }
        $name = (([string]$email.Subject) -replace '[\\/:*?"<>|\r\n\t]', '_').Trim()
        if (-not $name) { $name = "email" }
        if ($name.Length -gt 120) { $name = $name.Substring(0, 120) }
        $Path = Join-Path $dir "$name$ext"
    }
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $Path) {
        if ($Overwrite) {
            if ((Confirm-ItemAction -Action "overwrite" -Title "OVERWRITE FILE" -Item $email -Confirmed $Confirm -Details ([ordered]@{ File = $Path }) -What "the overwrite") -ne $true) { exit 0 }
        } else {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($Path); $e = [System.IO.Path]::GetExtension($Path); $c = 1
            do { $candidate = Join-Path $dir "${base}_$c$e"; $c++ } while (Test-Path -LiteralPath $candidate)
            $Path = $candidate
        }
    }
    $email.SaveAs($Path, $formats[$Format][0])
    $size = [math]::Round((Get-Item -LiteralPath $Path).Length / 1KB, 1)
    Write-ActionResult -Title "EMAIL SAVED" -Item $email -Changes ([ordered]@{ File = $Path; Format = $Format; SizeKB = $size }) -Json:$Json -Action "save-as"
}
catch {
    Stop-WithError $_
}
