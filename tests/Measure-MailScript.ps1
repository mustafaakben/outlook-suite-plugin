<#
.SYNOPSIS
  Runs one mail script and reports ONLY exit code, time, line count and simple pattern counts.

.DESCRIPTION
  Safe for test logs and transcripts: the script's output is never printed.
  -Count takes name=regex pairs; each prints how many times the regex matched.

.EXAMPLE
  ./tests/Measure-MailScript.ps1 outlook-read-body.ps1 -ScriptArgs @('-EntryID', $id) -Count @{ refs = '\[URL-\d{3}\]' }
#>
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Script,
    [object[]]$ScriptArgs = @(),
    [hashtable]$Count = @{},
    [string]$Label = "",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)
$path = Join-Path (Resolve-Path $ScriptsDir).Path $Script
$sw = [Diagnostics.Stopwatch]::StartNew()
$out = & pwsh -NoProfile -File $path @ScriptArgs 2>&1 | Out-String
$code = $LASTEXITCODE
$sw.Stop()
$counts = foreach ($k in ($Count.Keys | Sort-Object)) { "$k=$(([regex]::Matches($out, $Count[$k])).Count)" }
$name = if ($Label) { $Label } else { $Script }
"{0,-36} exit {1}  {2,6} ms  {3,5} lines  {4}" -f $name, $code, [int]$sw.Elapsed.TotalMilliseconds, ($out -split "`n").Count, ($counts -join "  ")
