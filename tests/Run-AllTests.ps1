<#
.SYNOPSIS
  Runs every unattended-safe tests/Test-*.ps1 once under PowerShell 7 and once against Windows
  PowerShell 5.1, with a send guard. Prints a summary table and exits 1 on any failure.

.DESCRIPTION
  - Each test runs in its own pwsh process. The 5.1 pass uses the test's own shell switch
    (-Exe powershell or -Shell powershell, read from its param block); tests without one run in
    the pwsh pass only.
  - Tests with a mandatory parameter (they need a choice, e.g. a mailbox) are listed as skipped.
    Benchmarks (Invoke-MailBench, Measure-MailScript) run only with -IncludeBench.
  - Send guard: the Sent Items and Outbox counts of every mailbox are read by this runner before
    and after EACH test. Any increase stops the whole run at once (exit 2).
  - OUTLOOK_SUITE_REQUIRE_CONFIRM=1 is forced for the run and restored afterwards.
  - Output of each test is shown as it runs; the tests themselves print counts only.

.EXAMPLE
  pwsh -File tests/Run-AllTests.ps1
  pwsh -File tests/Run-AllTests.ps1 -Only People,Tasks -SkipPS51
  pwsh -File tests/Run-AllTests.ps1 -List
#>
param(
    [string[]]$Only,          # run only tests whose name contains one of these words
    [switch]$SkipPS51,        # pwsh pass only
    [switch]$IncludeBench,    # also run the benchmarks
    [switch]$List             # show the plan and exit
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Run this with pwsh (PowerShell 7); the 5.1 pass is driven from here: pwsh -File tests/Run-AllTests.ps1" -ForegroundColor Red
    exit 1
}
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Plan
# ---------------------------------------------------------------------------

$bench = @("Invoke-MailBench.ps1", "Measure-MailScript.ps1")
$files = @(Get-ChildItem -Path $PSScriptRoot -Filter "Test-*.ps1" | Sort-Object Name)
if ($IncludeBench) { $files += @($bench | ForEach-Object { Get-Item -LiteralPath (Join-Path $PSScriptRoot $_) }) }
if ($Only) { $files = @($files | Where-Object { $n = $_.Name; @($Only | Where-Object { $n -like "*$_*" }).Count -gt 0 }) }

# Heuristic: a test "writes to the mailbox" when it creates, saves, moves or deletes items,
# adds folders, or passes -Confirm to a script.
$writePattern = 'CreateItem\(|\.Save\(\)|\.Delete\(\)|\.Move\(|Folders\.Add\(|"-Confirm"|-Confirm\b'

$plan = New-Object System.Collections.Generic.List[object]
foreach ($f in $files) {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
    $shellSwitch = ""; $mandatory = @()
    foreach ($p in @($ast.ParamBlock.Parameters)) {
        $name = $p.Name.VariablePath.UserPath
        $isMandatory = @($p.Attributes | ForEach-Object { $_.NamedArguments } | Where-Object { $_.ArgumentName -eq "Mandatory" }).Count -gt 0
        if ($isMandatory) { $mandatory += $name }
        $set = @($p.Attributes | Where-Object { $_.TypeName.Name -eq "ValidateSet" } | ForEach-Object { $_.PositionalArguments.Value })
        if ($name -in "Exe", "Shell" -and $set -contains "powershell") { $shellSwitch = $name }
    }
    $writes = [bool](Select-String -LiteralPath $f.FullName -Pattern $writePattern -Quiet)
    $plan.Add([pscustomobject]@{ Name = $f.BaseName; Path = $f.FullName; Switch = $shellSwitch; Mandatory = $mandatory; Writes = $writes })
}

$runnable = @($plan | Where-Object { $_.Mandatory.Count -eq 0 })
$skipped = @($plan | Where-Object { $_.Mandatory.Count -gt 0 })
Write-Host "=== Outlook suite tests ===" -ForegroundColor Cyan
Write-Host "Write to the mailbox (disposable [suite-test] items only): $(@($runnable | Where-Object Writes | ForEach-Object Name) -join ', ')" -ForegroundColor Yellow
Write-Host "Read-only: $(@($runnable | Where-Object { -not $_.Writes } | ForEach-Object Name) -join ', ')" -ForegroundColor Gray
Write-Host "5.1 pass: $(if ($SkipPS51) { 'skipped (-SkipPS51)' } else { @($runnable | Where-Object Switch | ForEach-Object { "$($_.Name) (-$($_.Switch))" }) -join ', ' })" -ForegroundColor Gray
foreach ($s in $skipped) { Write-Host "Skipped (needs -$($s.Mandatory -join ', -')): $($s.Name)" -ForegroundColor DarkGray }
if (-not $IncludeBench) { Write-Host "Skipped (benchmarks; add -IncludeBench): $($bench -join ', ')" -ForegroundColor DarkGray }
if ($List) { exit 0 }

# ---------------------------------------------------------------------------
# Send guard
# ---------------------------------------------------------------------------

$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")

# Item counts of Sent Items (5) and Outbox (4) in every mailbox, keyed "<store>|<folder>".
function Get-SendCounts {
    $c = @{}
    foreach ($store in $ns.Stores) {
        try { if ($store.ExchangeStoreType -eq 2) { continue } } catch { }   # public folders
        foreach ($id in 5, 4) {
            try { $f = $store.GetDefaultFolder($id); $c["$($store.StoreID)|$id"] = [int]$f.Items.Count } catch { }
        }
    }
    return $c
}

# Keys whose count went up, as "Sent Items +n" / "Outbox +n" (no store names: counts only).
function Compare-SendCounts($before, $after) {
    $up = @()
    foreach ($k in $after.Keys) {
        $b = if ($before.ContainsKey($k)) { $before[$k] } else { 0 }
        if ($after[$k] -gt $b) { $up += "$(if ($k.EndsWith('|5')) { 'Sent Items' } else { 'Outbox' }) +$($after[$k] - $b)" }
    }
    return $up
}

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

$runs = New-Object System.Collections.Generic.List[object]
foreach ($t in $runnable) { $runs.Add([pscustomobject]@{ Test = $t; Shell = "pwsh"; Args = @(if ($t.Switch) { "-$($t.Switch)"; "pwsh" }) }) }
if (-not $SkipPS51) {
    foreach ($t in @($runnable | Where-Object Switch)) { $runs.Add([pscustomobject]@{ Test = $t; Shell = "powershell"; Args = @("-$($t.Switch)", "powershell") }) }
}

$oldConfirm = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"
$results = New-Object System.Collections.Generic.List[object]
$stopped = ""
$total = [Diagnostics.Stopwatch]::StartNew()
try {
    $i = 0
    foreach ($r in $runs) {
        $i++
        Write-Host ""
        Write-Host "===== [$i/$($runs.Count)] $($r.Test.Name) ($($r.Shell)) =====" -ForegroundColor Cyan
        $before = Get-SendCounts
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $lines = New-Object System.Collections.Generic.List[string]
        & pwsh -NoProfile -File $r.Test.Path @($r.Args) 2>&1 | ForEach-Object { $s = [string]$_; $lines.Add($s); Write-Host $s }
        $exit = $LASTEXITCODE
        $secs = [Math]::Round($sw.Elapsed.TotalSeconds, 1)
        # "PASS  name" lines, or table rows ending in PASS/FAIL (Test-SuiteHost)
        $pass = @($lines | Where-Object { $_ -match '^\s*PASS\b|\sPASS\s*$' }).Count
        $fail = @($lines | Where-Object { $_ -match '^\s*FAIL\b|\sFAIL\s*$' }).Count
        if ($exit -ne 0 -and $fail -eq 0) { $fail = 1 }   # failed without a FAIL line (crash, parse error)
        $results.Add([pscustomobject]@{ Test = $r.Test.Name; Shell = $r.Shell; Pass = $pass; Fail = $fail; Seconds = $secs; Exit = $exit })
        $up = Compare-SendCounts $before (Get-SendCounts)
        if ($up.Count -gt 0) {
            $stopped = "$($r.Test.Name) ($($r.Shell)): $($up -join ', ')"
            Write-Host ""
            Write-Host "STOPPED: the send guard saw $($up -join ', ') during $($r.Test.Name) ($($r.Shell)). Check Sent Items and Outbox now." -ForegroundColor Red
            break
        }
    }
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $oldConfirm
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== SUMMARY ($([Math]::Round($total.Elapsed.TotalMinutes, 1)) min) ===" -ForegroundColor Cyan
Write-Host ("{0,-24} {1,-11} {2,5} {3,5} {4,8}" -f "Test", "Shell", "Pass", "Fail", "Seconds")
foreach ($x in $results) {
    $color = if ($x.Fail -gt 0) { "Red" } else { "Green" }
    Write-Host ("{0,-24} {1,-11} {2,5} {3,5} {4,8}" -f $x.Test, $x.Shell, $x.Pass, $x.Fail, $x.Seconds) -ForegroundColor $color
}
$failed = @($results | Where-Object { $_.Fail -gt 0 })
$sumPass = ($results | Measure-Object Pass -Sum).Sum; $sumFail = ($results | Measure-Object Fail -Sum).Sum
Write-Host ("{0,-24} {1,-11} {2,5} {3,5}" -f "TOTAL", "$($results.Count) runs", $sumPass, $sumFail)
if ($stopped) { Write-Host "Run STOPPED by the send guard: $stopped" -ForegroundColor Red; exit 2 }
if ($failed.Count -gt 0) { Write-Host "$($failed.Count) run(s) failed: $(@($failed | ForEach-Object { "$($_.Test) ($($_.Shell))" }) -join ', ')" -ForegroundColor Red; exit 1 }
Write-Host "All runs passed; Sent Items and Outbox unchanged." -ForegroundColor Green
exit 0
