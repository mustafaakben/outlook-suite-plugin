<#
.SYNOPSIS
  Read-only smoke test of the whole suite through the warm host (outlook-call.ps1, PowerShell 7).

.DESCRIPTION
  Runs one representative read command from each skill through outlook-call.ps1 and checks
  that it exits 0, ran in the host (no "host unavailable" fallback) and returned valid -Json.
  Prints timings only (no names, subjects or addresses). Each command runs -Repeat times
  (default 2): the first call shows the cold cost, the later ones the warm path.
  Never stops the host; it starts it if it is not running (outlook-call does that on first use).
  Usage: pwsh -File tests/Test-SuiteHost.ps1 [-Repeat 3]
#>
param(
    [int]$Repeat = 2,
    [string]$SkillsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills")
)

if ($PSVersionTable.PSVersion.Major -lt 7) { Write-Host "Test-SuiteHost needs PowerShell 7 (the host pipe is pwsh-only)." -ForegroundColor Red; exit 1 }
$SkillsDir = (Resolve-Path $SkillsDir).Path
$call = Join-Path $SkillsDir "outlook-mail/scripts/outlook-call.ps1"

$commands = @(
    @{ Skill = "mail";      Name = "find";          Args = @("find", "-Days", "1", "-Limit", "10", "-Json") },
    @{ Skill = "calendar";  Name = "calendar-list"; Args = @("calendar-list", "-Days", "7", "-Json") },
    @{ Skill = "calendar";  Name = "calendar-free"; Args = @("calendar-free", "-Days", "7", "-Json") },
    @{ Skill = "calendar";  Name = "calendar-week"; Args = @("calendar-week", "-Json") },
    @{ Skill = "contacts";  Name = "people";        Args = @("people", "dan", "-Limit", "5", "-Json") },
    @{ Skill = "contacts";  Name = "contacts-list"; Args = @("contacts-list", "-Limit", "10", "-Json") },
    @{ Skill = "tasks";     Name = "tasks-list";    Args = @("tasks-list", "-Due", "Week", "-Flagged", "-Json") },
    @{ Skill = "tasks";     Name = "agenda";        Args = @("agenda", "-Week", "-Json") },
    @{ Skill = "notes";     Name = "notes-list";    Args = @("notes-list", "-Json") }
)

$fails = 0
$hostUp = (& pwsh -NoProfile -File $call --ping 2>&1 | Out-String) -match "pong"
Write-Host "Host before the run: $(if ($hostUp) { 'running' } else { 'not running (the first call starts it)' })" -ForegroundColor Gray
Write-Host ("{0,-10} {1,-15} {2}" -f "skill", "command", (@(1..$Repeat | ForEach-Object { "run $_".PadLeft(8) }) -join "")) -ForegroundColor Cyan

foreach ($c in $commands) {
    $times = @(); $ok = $true; $why = ""
    for ($i = 1; $i -le $Repeat; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $out = & pwsh -NoProfile -File $call @($c.Args) 2>&1 | Out-String
        $code = $LASTEXITCODE
        $times += $sw.Elapsed.TotalSeconds
        if ($code -ne 0) { $ok = $false; $why = "exit $code" }
        elseif ($out -match "host unavailable") { $ok = $false; $why = "ran outside the host" }
        else { try { $null = $out | ConvertFrom-Json -ErrorAction Stop } catch { $ok = $false; $why = "invalid JSON" } }
    }
    $cells = @($times | ForEach-Object { ("{0:N2}s" -f $_).PadLeft(8) }) -join ""
    if ($ok) { Write-Host ("{0,-10} {1,-15} {2}  PASS" -f $c.Skill, $c.Name, $cells) -ForegroundColor Green }
    else { $fails++; Write-Host ("{0,-10} {1,-15} {2}  FAIL ({3})" -f $c.Skill, $c.Name, $cells, $why) -ForegroundColor Red }
}

Write-Host ""
if ($fails -eq 0) { Write-Host "ALL PASS ($($commands.Count) commands through the host)" -ForegroundColor Green } else { Write-Host "$fails of $($commands.Count) FAILED" -ForegroundColor Red }
exit $fails
