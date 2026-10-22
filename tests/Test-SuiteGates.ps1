<#
.SYNOPSIS
  Guards the confirm gate in every outlook-suite skill.

.DESCRIPTION
  A confirm helper that writes anything to the pipeline turns its "$false" into a truthy array and
  opens the gate (incident 2026-10-22). This test checks, for every script under skills\:

  1. Every file parses without errors (run it under pwsh AND powershell to cover 5.1).
  2. Every gate helper outputs only a boolean. Gate helpers are functions named Confirm-* or
     Approve-*, Test-ConfirmRequired, and any function that calls Approve-Action or
     Test-ConfirmRequired. Each statement must be silent, following the helpers it calls, and each
     return must be $true, $false, a [bool] cast, or a call to another gate helper.
  3. Every call of a Confirm-*/Approve-* helper is fail-closed. Accepted forms:
       (Confirm-X ...) -ne $true      (Confirm-X ...) -eq $true
       $null = Approve-X ...          return (Approve-X ...)   (inside another gate helper)
  4. Live: Approve-Action and Confirm-ItemAction return exactly one [bool] $false when confirmation
     is required, on a throwaway draft the test creates and deletes.

  Silent statements inside gate helpers are: assignments, [void](...), "| Out-Null", control flow,
  Write-Host and the other host-only commands (without stream merging such as 6>&1), and calls to
  helpers that are silent themselves. Anything else, including bare method calls or $i++, is
  treated as possible output; wrap it in [void](...) or assign it to $null.

  Prints PASS/FAIL and counts only.

.EXAMPLE
  pwsh -NoProfile -File tests\Test-SuiteGates.ps1
  powershell -NoProfile -File tests\Test-SuiteGates.ps1
#>
param([string]$SkillsDir = (Join-Path $PSScriptRoot "..\plugins\outlook-suite\skills"))

$ErrorActionPreference = "Stop"
$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
$L = "System.Management.Automation.Language"

# ---------------------------------------------------------------------------
# Parse every script
# ---------------------------------------------------------------------------
$files = @(Get-ChildItem -Path $SkillsDir -Recurse -Filter *.ps1 -File)
$funcs = @{}   # name -> list of FunctionDefinitionAst (every definition is checked)
$asts = @{}
$parseErrors = 0
foreach ($f in $files) {
    $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errs)
    if ($errs.Count -gt 0) { $parseErrors++; Write-Host "      parse error: $($f.Name):$($errs[0].Extent.StartLineNumber)" -ForegroundColor DarkGray }
    $asts[$f.FullName] = $ast
    foreach ($fd in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if (-not $funcs.ContainsKey($fd.Name)) { $funcs[$fd.Name] = New-Object System.Collections.ArrayList }
        [void]$funcs[$fd.Name].Add($fd)
    }
}
Check "scripts found ($($files.Count))" ($files.Count -ge 50)
Check "no parse errors under PowerShell $($PSVersionTable.PSVersion.Major) ($parseErrors files)" ($parseErrors -eq 0)

function Get-CalledNames($fd) {
    return @($fd.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
        ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
}

# Gate helpers are found by name. A helper with another name that wraps the gate is still covered:
# its own Approve-*/Confirm-* call must pass the fail-closed call-site check below.
$gateNames = New-Object System.Collections.Generic.HashSet[string]
foreach ($n in $funcs.Keys) {
    if ($n -match '^(Confirm|Approve)-' -or $n -eq "Test-ConfirmRequired") { [void]$gateNames.Add($n) }
}
$required = @("Approve-Action", "Test-ConfirmRequired", "Confirm-ItemAction")
Check "core gate helpers found ($($gateNames.Count) gate helpers)" ($gateNames.Count -ge 5 -and @($required | Where-Object { -not $gateNames.Contains($_) }).Count -eq 0)

# ---------------------------------------------------------------------------
# 2. Gate helpers output only a boolean
# ---------------------------------------------------------------------------
$hostOnly = @("Write-Host", "Write-Warning", "Write-Verbose", "Write-Debug", "Write-Information", "Out-Null", "Out-Host",
    "Set-Variable", "Remove-Variable", "Start-Sleep", "Set-StrictMode", "Set-Location", "Remove-Item")
$silentMemo = @{}

function Test-BoolReturn($pipeline) {
    if (-not $pipeline) { return $true }
    $t = $pipeline.Extent.Text.Trim()
    if ($t -in '$true', '$false') { return $true }
    if ($t -match '^\[bool\]') { return $true }
    $cmd = $pipeline.Find({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    if ($cmd -and $t -match '^\(?\s*&?\s*[A-Za-z-]+' -and $gateNames.Contains([string]$cmd.GetCommandName()) -and
        $pipeline.Extent.Text.Trim().TrimStart('(').Trim().StartsWith($cmd.GetCommandName())) { return $true }
    return $false
}

function Test-SilentStatement($s, [bool]$isGate) {
    if ($s -is [System.Management.Automation.Language.AssignmentStatementAst]) { return $true }
    if ($s -is [System.Management.Automation.Language.ReturnStatementAst]) {
        if (-not $s.Pipeline) { return $true }
        return ($isGate -and (Test-BoolReturn $s.Pipeline))
    }
    if ($s -is [System.Management.Automation.Language.ExitStatementAst] -or
        $s -is [System.Management.Automation.Language.ThrowStatementAst] -or
        $s -is [System.Management.Automation.Language.BreakStatementAst] -or
        $s -is [System.Management.Automation.Language.ContinueStatementAst] -or
        $s -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $true }
    if ($s -is [System.Management.Automation.Language.IfStatementAst]) {
        foreach ($c in $s.Clauses) { if (-not (Test-SilentBlock $c.Item2 $isGate)) { return $false } }
        if ($s.ElseClause) { return (Test-SilentBlock $s.ElseClause $isGate) }
        return $true
    }
    if ($s -is [System.Management.Automation.Language.SwitchStatementAst]) {
        foreach ($c in $s.Clauses) { if (-not (Test-SilentBlock $c.Item2 $isGate)) { return $false } }
        if ($s.Default) { return (Test-SilentBlock $s.Default $isGate) }
        return $true
    }
    if ($s -is [System.Management.Automation.Language.LoopStatementAst] -or $s -is [System.Management.Automation.Language.ForEachStatementAst]) {
        return (Test-SilentBlock $s.Body $isGate)
    }
    if ($s -is [System.Management.Automation.Language.TryStatementAst]) {
        if (-not (Test-SilentBlock $s.Body $isGate)) { return $false }
        foreach ($c in $s.CatchClauses) { if (-not (Test-SilentBlock $c.Body $isGate)) { return $false } }
        if ($s.Finally) { return (Test-SilentBlock $s.Finally $isGate) }
        return $true
    }
    if ($s -is [System.Management.Automation.Language.PipelineAst]) {
        $els = $s.PipelineElements
        $last = $els[$els.Count - 1]
        if ($last -is [System.Management.Automation.Language.CommandAst] -and $last.GetCommandName() -eq "Out-Null" -and $last.Redirections.Count -eq 0) { return $true }
        if ($els.Count -eq 1 -and $els[0] -is [System.Management.Automation.Language.CommandExpressionAst]) {
            $e = $els[0].Expression
            return ($e -is [System.Management.Automation.Language.ConvertExpressionAst] -and $e.Type.TypeName.Name -eq "void")
        }
        if ($els.Count -eq 1 -and $els[0] -is [System.Management.Automation.Language.CommandAst]) {
            $c = $els[0]
            # 6>&1, *>&1 and similar merge host output into the success stream.
            if (@($c.Redirections | Where-Object { $_ -is [System.Management.Automation.Language.MergingRedirectionAst] }).Count -gt 0) { return $false }
            $name = $c.GetCommandName()
            if ($name -in $hostOnly) { return $true }
            if ($name -and $funcs.ContainsKey($name)) { return (Test-SilentFunction $name) }
        }
        return $false
    }
    return $false
}

function Test-SilentBlock($block, [bool]$isGate) {
    if (-not $block) { return $true }
    foreach ($s in $block.Statements) { if (-not (Test-SilentStatement $s $isGate)) { return $false } }
    return $true
}

function Test-FunctionBody($fd, [bool]$isGate) {
    $b = $fd.Body
    return ((Test-SilentBlock $b.BeginBlock $isGate) -and (Test-SilentBlock $b.ProcessBlock $isGate) -and (Test-SilentBlock $b.EndBlock $isGate))
}

# A non-gate helper called from a gate helper must write nothing at all (every definition).
function Test-SilentFunction([string]$name) {
    if ($gateNames.Contains($name)) { return $true }   # checked on its own, as a gate helper
    if ($silentMemo.ContainsKey($name)) { return $silentMemo[$name] }
    $silentMemo[$name] = $true   # cycles count as silent
    $ok = $true
    foreach ($fd in $funcs[$name]) { if (-not (Test-FunctionBody $fd $false)) { $ok = $false } }
    $silentMemo[$name] = $ok
    return $ok
}

$bad = @()
foreach ($n in $gateNames) {
    foreach ($fd in $funcs[$n]) { if (-not (Test-FunctionBody $fd $true)) { $bad += "$n ($([IO.Path]::GetFileName($fd.Extent.File)):$($fd.Extent.StartLineNumber))" } }
}
Check "gate helpers output only a boolean" ($bad.Count -eq 0) ("leaky: " + ($bad -join ", "))

# ---------------------------------------------------------------------------
# 3. Every Confirm-*/Approve-* call is fail-closed
# ---------------------------------------------------------------------------
function Get-EnclosingFunction($node) {
    $p = $node.Parent
    while ($p) { if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $p }; $p = $p.Parent }
    return $null
}

function Test-ClosedCall($cmd) {
    $pipe = $cmd.Parent
    if (-not ($pipe -is [System.Management.Automation.Language.PipelineAst])) { return $false }
    # Approve-X ... | Out-Null : preview text only, the result is discarded (the decision was made before).
    if ($pipe.PipelineElements.Count -eq 2 -and $pipe.PipelineElements[0] -eq $cmd -and
        $pipe.PipelineElements[1] -is [System.Management.Automation.Language.CommandAst] -and
        $pipe.PipelineElements[1].GetCommandName() -eq "Out-Null") { return $true }
    if ($pipe.PipelineElements.Count -ne 1) { return $false }
    $up = $pipe.Parent
    # $null = Approve-X ...
    if ($up -is [System.Management.Automation.Language.AssignmentStatementAst]) { return ($up.Left.Extent.Text -eq '$null') }
    # return Approve-X ... / return (Approve-X ...) inside another gate helper
    $r = $up
    if ($r -is [System.Management.Automation.Language.ParenExpressionAst]) { $r = $r.Parent.Parent.Parent }   # paren -> command expression -> pipeline -> return
    if ($r -is [System.Management.Automation.Language.ReturnStatementAst]) {
        $fn = Get-EnclosingFunction $cmd
        return ($fn -and $gateNames.Contains($fn.Name))
    }
    # (Confirm-X ...) -ne $true / -eq $true
    if ($up -is [System.Management.Automation.Language.ParenExpressionAst]) {
        $bin = $up.Parent
        if ($bin -is [System.Management.Automation.Language.CommandExpressionAst]) { $bin = $bin.Parent }
        if ($bin -is [System.Management.Automation.Language.BinaryExpressionAst] -and $bin.Left -eq $up -and
            $bin.Operator -in "Ine", "Ieq", "Cne", "Ceq" -and $bin.Right.Extent.Text -eq '$true') { return $true }
    }
    return $false
}

$open = 0
foreach ($k in $asts.Keys) {
    $calls = $asts[$k].FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and [string]$n.GetCommandName() -match '^(Confirm|Approve)-' }, $true)
    foreach ($c in $calls) {
        if (-not (Test-ClosedCall $c)) { $open++; Write-Host "      not fail-closed: $([IO.Path]::GetFileName($k)):$($c.Extent.StartLineNumber)" -ForegroundColor DarkGray }
    }
}
Check "every Confirm-*/Approve-* call is fail-closed ($open open)" ($open -eq 0)

# ---------------------------------------------------------------------------
# 4. Live: the shared gate returns exactly one [bool] $false when confirmation is required
# ---------------------------------------------------------------------------
$prevGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"
$draft = $null
try {
    . (Join-Path $SkillsDir "outlook-mail\scripts\_common.ps1")
    $r = @(Approve-Action -Action "send" -Confirmed $false -What "gate test" 6>$null)
    Check "Approve-Action returns one [bool] `$false" ($r.Count -eq 1 -and $r[0] -is [bool] -and -not $r[0])
    $null = Connect-Outlook
    $draft = $Outlook.CreateItem(0)
    $draft.Subject = "[suite-test] gate test"
    $draft.Save()
    $r = @(Confirm-ItemAction -Action "delete" -Title "GATE TEST" -Item $draft -Confirmed $false -What "nothing" 6>$null)
    Check "Confirm-ItemAction returns one [bool] `$false" ($r.Count -eq 1 -and $r[0] -is [bool] -and -not $r[0])
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $prevGate
    if ($draft) {
        try {
            $id = $draft.EntryID; $draft.Delete()
            $q = "[Subject] = '[suite-test] gate test'"
            foreach ($i in @($Namespace.GetDefaultFolder(3).Items.Restrict($q))) { $i.Delete() }
        } catch { }
    }
}

Write-Host ""
if ($script:fails -eq 0) { Write-Host "All gate checks passed." -ForegroundColor Green; exit 0 }
Write-Host "$($script:fails) gate check(s) failed." -ForegroundColor Red; exit 1
