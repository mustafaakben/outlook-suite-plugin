<#
.SYNOPSIS
  Moving an email changes its EntryID and WebLink but not its Message-ID, and
  outlook-read-body.ps1 -MessageID finds it again with the current EntryID and WebLink.
  Writes only to a test draft (subject "[suite-test] ...", no recipients, never sent) that ends in
  the hidden folder "Deleted Items/_suite-test". Prints counts and PASS/FAIL only.
#>
param([string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"))
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$ReadBody = Join-Path $ScriptsDir "outlook-read-body.ps1"
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Read([string[]]$a) {
    $out = & pwsh -NoProfile -File $ReadBody @a -Json 2>&1 | Out-String
    $code = $LASTEXITCODE
    $j = $null; try { $j = $out | ConvertFrom-Json } catch { }
    return [pscustomobject]@{ Exit = $code; Out = $out; Json = $j }
}

$app = New-Object -ComObject Outlook.Application
$ns = $app.Session
$deletedRoot = $ns.GetDefaultFolder(3)
$sandbox = $null
foreach ($f in $deletedRoot.Folders) { if ($f.Name -eq "_suite-test") { $sandbox = $f } }
if (-not $sandbox) { $sandbox = $deletedRoot.Folders.Add("_suite-test") }

# A test draft with a known Message-ID (Outlook keeps one set on an unsent item).
$tag = [guid]::NewGuid().ToString("N")
$mid = "<suite-test-$tag@example.invalid>"
$midTag = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
$draft = $app.CreateItem(0)
$draft.Subject = "[suite-test] links $tag"
$draft.Body = "Test item for Test-MailLinks.ps1. It has no recipients and is never sent."
$draft.Save()
try { $draft.PropertyAccessor.SetProperty($midTag, $mid); $draft.Save() } catch { }
$kept = $false; try { $kept = ([string]$draft.PropertyAccessor.GetProperty($midTag)) -eq $mid } catch { }
if (-not $kept) {
    $null = $draft.Move($sandbox)
    Write-Host "SKIP  Outlook did not keep a Message-ID on the test draft; nothing to test" -ForegroundColor Yellow
    exit 0
}

$before = Read @("-EntryID", $draft.EntryID)
Check "draft: read-body shows its Message-ID and a WebLink" ($before.Json.MessageID -eq $mid -and $before.Json.WebLink -match '^https://')

$moved = $draft.Move($sandbox)
Check "moving to another folder changes the EntryID" ($moved.EntryID -ne $before.Json.EntryID)

$after = Read @("-MessageID", $mid)
Check "-MessageID finds the moved email" ($after.Exit -eq 0 -and $after.Json.EntryID -eq $moved.EntryID -and $after.Json.Folder -match '_suite-test$')
Check "and returns a new WebLink" ($after.Json.WebLink -match '^https://' -and $after.Json.WebLink -ne $before.Json.WebLink)

$direct = Read @("-EntryID", $moved.EntryID)
Check "the new WebLink is the moved email's own" ($direct.Json.WebLink -eq $after.Json.WebLink)

$bare = Read @("-MessageID", $mid.Trim('<', '>'))
Check "-MessageID also works without angle brackets" ($bare.Json.EntryID -eq $moved.EntryID)

$none = Read @("-MessageID", "<no-such-$tag@example.invalid>")
Check "an unknown Message-ID exits 1 with a clear error" ($none.Exit -eq 1 -and $none.Out -match "No email with Message-ID")

if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
