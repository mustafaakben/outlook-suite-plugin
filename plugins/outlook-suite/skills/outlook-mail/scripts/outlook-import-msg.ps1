param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path,
    [string]$Folder = "Inbox",
    [switch]$Json
)

# Outlook Import .msg: bring saved .msg files (or a folder of them) into a mail folder.
# Prints each imported email's EntryID.
# Usage: .\outlook-import-msg.ps1 -Path "C:\Mail\note.msg" -Folder "Archive/Imported"
# Usage: .\outlook-import-msg.ps1 -Path "C:\Mail\*.msg"

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $target = Resolve-OutlookFolder -Namespace $Namespace -Path $Folder
    # -LiteralPath when the file exists (names like "[project] note.msg" contain wildcard
    # characters); otherwise treat the path as a wildcard pattern such as C:\Mail\*.msg.
    $files = @(foreach ($p in $Path) {
        $items = if (Test-Path -LiteralPath $p) { Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue } else { Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue }
        $items | Where-Object { $_.Extension -ieq ".msg" }
    })
    if ($files.Count -eq 0) { Stop-WithError "No .msg files found at: $($Path -join ', ')" }
    $done = @(); $failed = @()
    foreach ($f in $files) {
        try {
            $it = $Namespace.OpenSharedItem($f.FullName)
            $m = $it.Move($target)
            $done += [ordered]@{ File = $f.Name; EntryID = $m.EntryID; Subject = $m.Subject }
        } catch { $failed += [ordered]@{ File = $f.Name; Error = $_.Exception.Message } }
    }
    if ($Json) { Write-Json ([ordered]@{ Folder = $target.FolderPath; Imported = $done; Failed = $failed }); if ($failed) { exit 1 } else { exit 0 } }
    Write-Host "=== IMPORTED $($done.Count) of $($files.Count) into $($target.FolderPath) ===" -ForegroundColor $(if ($failed) { "Yellow" } else { "Green" })
    foreach ($d in $done) { Write-Host "  $($d.File) -> EntryID: $($d.EntryID)" }
    foreach ($x in $failed) { Write-Host "  Failed: $($x.File): $($x.Error)" -ForegroundColor Red }
    if ($failed) { exit 1 }
}
catch {
    Stop-WithError $_
}
