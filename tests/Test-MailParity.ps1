<#
.SYNOPSIS
  Checks that a candidate copy of the mail scripts returns the same emails as a baseline copy.

.DESCRIPTION
  Runs the same read-only query against two script folders (for example the main
  branch clone and the improvement worktree) and compares the sets of EntryIDs each
  one prints. Only counts and match status are shown; no mail content is printed.

.EXAMPLE
  ./tests/Test-MailParity.ps1 -BaselineDir ../../outlook-suite-plugin/plugins/outlook-suite/skills/outlook-mail/scripts
#>
param(
    [Parameter(Mandatory = $true)][string]$BaselineDir,
    [string]$CandidateDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"),
    [string[]]$Only
)

$ErrorActionPreference = "Stop"
$BaselineDir = (Resolve-Path $BaselineDir).Path
$CandidateDir = (Resolve-Path $CandidateDir).Path

# Read-only cases. Each case may give different argument lists for baseline and candidate
# when a parameter was renamed (CandidateArgs defaults to Args).
$cases = @(
    @{ Name = "find-7d-50";        Script = "outlook-find.ps1";   Args = @("-Days", "7", "-Limit", "50") }
    @{ Name = "find-30d-200";      Script = "outlook-find.ps1";   Args = @("-Days", "30", "-Limit", "200") }
    @{ Name = "find-30d-unread";   Script = "outlook-find.ps1";   Args = @("-Days", "30", "-Limit", "200", "-UnreadOnly") }
    @{ Name = "find-30d-attach";   Script = "outlook-find.ps1";   Args = @("-Days", "30", "-Limit", "200", "-HasAttachment") }
    @{ Name = "find-30d-subject";  Script = "outlook-find.ps1";   Args = @("-Days", "30", "-Limit", "200", "-Subject", "meeting") }
    @{ Name = "find-30d-body";     Script = "outlook-find.ps1";   Args = @("-Days", "30", "-Limit", "200", "-BodyContains", "thank") }
    @{ Name = "find-sent-14d";     Script = "outlook-find.ps1";   Args = @("-Folder", "Sent", "-Days", "14", "-Limit", "100") }
    @{ Name = "read-3d-50";        Script = "outlook-read.ps1";   Args = @("-Days", "3", "-Limit", "50") }
    @{ Name = "search-30d-subject"; Script = "outlook-search.ps1"; Args = @("-Subject", "meeting", "-Days", "30", "-Limit", "100") }
)
if ($Only) { $cases = $cases | Where-Object { $Only -contains $_.Name } }

function Get-EntryIds([string]$dir, [string]$script, [string[]]$scriptArgs) {
    $out = & pwsh -NoProfile -File (Join-Path $dir $script) @scriptArgs 2>&1 | Out-String
    $ids = [regex]::Matches($out, '(?m)EntryID:\s*([0-9A-Fa-f]{20,})') | ForEach-Object { $_.Groups[1].Value }
    [pscustomobject]@{ Ids = @($ids); Exit = $LASTEXITCODE; Error = [bool]($out -match '(?m)^\s*Error:') }
}

$results = foreach ($case in $cases) {
    $candArgs = if ($case.CandidateArgs) { $case.CandidateArgs } else { $case.Args }
    $base = Get-EntryIds $BaselineDir $case.Script $case.Args
    $cand = Get-EntryIds $CandidateDir $case.Script $candArgs
    $onlyBase = @($base.Ids | Where-Object { $cand.Ids -notcontains $_ })
    $onlyCand = @($cand.Ids | Where-Object { $base.Ids -notcontains $_ })
    [pscustomobject]@{
        Case       = $case.Name
        Baseline   = $base.Ids.Count
        Candidate  = $cand.Ids.Count
        OnlyBase   = $onlyBase.Count
        OnlyCand   = $onlyCand.Count
        SameOrder  = (($base.Ids -join ",") -eq ($cand.Ids -join ","))
        Errors     = "$([int]$base.Error)/$([int]$cand.Error)"
        Match      = ($onlyBase.Count -eq 0 -and $onlyCand.Count -eq 0)
    }
}

$results | Format-Table -AutoSize
if ($results | Where-Object { -not $_.Match }) { exit 1 }
