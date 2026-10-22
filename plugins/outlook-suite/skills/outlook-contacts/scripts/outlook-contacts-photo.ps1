param(
    [string]$EntryID = "",
    [int]$Index = 0,

    [string]$PhotoPath = "",
    [switch]$Remove,
    [string]$Search = "",
    [switch]$Json
)

# Outlook Contact Photo: add or remove a contact's picture (JPG, JPEG, PNG, GIF, BMP).
# Usage: .\outlook-contacts-photo.ps1 -EntryID "00000000..." -PhotoPath "C:\Photos\john.jpg"
# Usage: .\outlook-contacts-photo.ps1 -EntryID "00000000..." -Remove

. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_people.ps1")

if (-not $PhotoPath -and -not $Remove) { Stop-WithError "Give -PhotoPath <file> or -Remove." }
if ($PhotoPath -and $Remove) { Stop-WithError "-PhotoPath and -Remove cannot be used together." }

try {
    $full = ""
    if ($PhotoPath) {
        if (-not (Test-Path -LiteralPath $PhotoPath -PathType Leaf)) { Stop-WithError "Photo not found: $PhotoPath" }
        $full = (Resolve-Path -LiteralPath $PhotoPath).Path   # AddPicture needs an absolute path
        $ext = [IO.Path]::GetExtension($full).ToLowerInvariant()
        if ($ext -notin ".jpg", ".jpeg", ".png", ".gif", ".bmp") { Stop-WithError "Unsupported photo type '$ext'." -Hint "Use JPG, JPEG, PNG, GIF or BMP." }
    }
    $null = Connect-Outlook
    $contact = Get-ContactTarget -Namespace $Namespace -EntryID $EntryID -Index $Index -Search $Search
    if ($Remove) {
        if (-not $contact.HasPicture) {
            Write-ContactResult -Title "NO PHOTO TO REMOVE" -Rec (ConvertTo-ContactRecord $contact) -Changes ([ordered]@{ Photo = "none" }) -Action "photo-remove" -Json:$Json
            exit 0
        }
        $contact.RemovePicture()
        $contact.Save()
        Write-ContactResult -Title "CONTACT PHOTO REMOVED" -Rec (ConvertTo-ContactRecord $contact) -Changes ([ordered]@{ Photo = "removed" }) -Action "photo-remove" -Json:$Json
    } else {
        $contact.AddPicture($full)
        $contact.Save()
        Write-ContactResult -Title "CONTACT PHOTO SET" -Rec (ConvertTo-ContactRecord $contact) -Changes ([ordered]@{ Photo = [IO.Path]::GetFileName($full) }) -Action "photo-add" -Json:$Json
    }
}
catch {
    Stop-WithError $_
}
