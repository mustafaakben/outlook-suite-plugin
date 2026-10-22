<#
.SYNOPSIS
  Times read-only mail scripts against the live Outlook profile.

.DESCRIPTION
  Runs each benchmark case several times with `pwsh -NoProfile -File` and reports
  elapsed milliseconds and output line counts only. No mail content is printed or
  saved. Results are appended to tests/out/bench.csv (gitignored).

.EXAMPLE
  ./tests/Invoke-MailBench.ps1 -Label baseline
  ./tests/Invoke-MailBench.ps1 -Label branch -ScriptsDir ../other-checkout/plugins/outlook-suite/skills/outlook-mail/scripts
#>
param(
    [string]$Label = "run",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts"),
    [int]$Runs = 3,
    [string[]]$Only,
    [switch]$IncludeSlow
)

$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
$outDir = Join-Path $PSScriptRoot "out"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# Read-only cases only. Name = label in the report; Script + Args = what runs.
$cases = @(
    @{ Name = "accounts-list";   Script = "outlook-accounts-list.ps1";   Args = @() }
    @{ Name = "list-categories"; Script = "outlook-list-categories.ps1"; Args = @() }
    @{ Name = "read-1d-10";      Script = "outlook-read.ps1";            Args = @("-Days", "1", "-Limit", "10") }
    @{ Name = "find-7d-50";      Script = "outlook-find.ps1";            Args = @("-Days", "7", "-Limit", "50") }
    @{ Name = "find-30d-200";    Script = "outlook-find.ps1";            Args = @("-Days", "30", "-Limit", "200") }
    @{ Name = "search-30d";      Script = "outlook-search.ps1";          Args = @("-Subject", "zzqx-no-match", "-Days", "30") }
    @{ Name = "folders";         Script = "outlook-folders.ps1";         Args = @(); Slow = $true }
)

if ($Only) { $cases = $cases | Where-Object { $Only -contains $_.Name } }
if (-not $IncludeSlow -and -not $Only) { $cases = $cases | Where-Object { -not $_.Slow } }

$rows = foreach ($case in $cases) {
    $path = Join-Path $ScriptsDir $case.Script
    if (-not (Test-Path $path)) { Write-Warning "Missing $($case.Script)"; continue }
    foreach ($i in 1..$Runs) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $out = & pwsh -NoProfile -File $path @($case.Args) 2>&1 | Out-String
        $code = $LASTEXITCODE
        $sw.Stop()
        [pscustomobject]@{
            Timestamp = (Get-Date).ToString("s")
            Label     = $Label
            Case      = $case.Name
            Run       = $i
            Ms        = [int]$sw.Elapsed.TotalMilliseconds
            Lines     = ($out -split "`n").Count
            ExitCode  = $code
            HasError  = [bool]($out -match '(?m)^\s*Error:')
        }
    }
}

$rows | Export-Csv -Path (Join-Path $outDir "bench.csv") -Append -NoTypeInformation

$rows | Group-Object Case | ForEach-Object {
    $ms = $_.Group.Ms | Sort-Object
    [pscustomobject]@{
        Case     = $_.Name
        MedianMs = $ms[[int][math]::Floor($ms.Count / 2)]
        MinMs    = $ms[0]
        Lines    = ($_.Group.Lines | Select-Object -First 1)
        Exit     = ($_.Group.ExitCode | Sort-Object -Unique) -join ","
        Errors   = ($_.Group | Where-Object HasError).Count
    }
} | Format-Table -AutoSize
