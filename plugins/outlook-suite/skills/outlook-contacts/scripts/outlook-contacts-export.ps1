param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$Path = "",
    [string]$Search = "",
    [switch]$Json
)

# Outlook Contact Export: save one contact as a vCard (.vcf). Default folder: Downloads.
# An existing file is never overwritten (name_1.vcf, name_2.vcf, ...).
# Usage: .\outlook-contacts-export.ps1 -EntryID "00000000..."
# Usage: .\outlook-contacts-export.ps1 -EntryID "00000000..." -Path "C:\Contacts\john.vcf" -Json

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

try {
    $null = Connect-Outlook
    $contact = Get-ContactTarget -Namespace $Namespace -EntryID $EntryID -Index $Index -Search $Search
    $name = [string]$contact.FullName

    $isDir = $Path -and ($Path.Trim().EndsWith("\") -or $Path.Trim().EndsWith("/") -or (Test-Path -LiteralPath $Path -PathType Container))
    if (-not $Path -or $isDir) {
        $folder = if ($Path) { $Path } else { Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads" }
        $clean = ($name -replace '[\\/:*?"<>|]', '_').Trim()
        if (-not $clean) { $clean = "contact" }
        $Path = Join-Path $folder ($clean.Substring(0, [Math]::Min(50, $clean.Length)) + ".vcf")
    }
    # Absolute path: Outlook's SaveAs and .NET both resolve relative paths from the process directory.
    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (-not $Path.ToLowerInvariant().EndsWith(".vcf")) { $Path += ".vcf" }
    $dir = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $base = [IO.Path]::GetFileNameWithoutExtension($Path); $n = 1
    while (Test-Path -LiteralPath $Path) { $Path = Join-Path $dir "${base}_$n.vcf"; $n++ }

    $contact.SaveAs($Path, 6)   # olVCard
    $kb = [Math]::Round((Get-Item -LiteralPath $Path).Length / 1KB, 1)
    Write-ContactResult -Title "CONTACT EXPORTED" -Rec (ConvertTo-ContactRecord $contact) -Changes ([ordered]@{ Path = $Path; SizeKB = $kb; Format = "vCard" }) -Action "export" -Json:$Json
}
catch {
    Stop-WithError $_
}
