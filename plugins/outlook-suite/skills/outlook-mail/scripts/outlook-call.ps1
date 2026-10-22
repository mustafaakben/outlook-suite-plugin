# Run an Outlook mail script through the warm host (fast path).
#
# Usage:
#   outlook-call.ps1 <script> [script arguments...]
#     <script> can be "find", "outlook-find", "outlook-find.ps1" or a full path.
#     Arguments are exactly what you would pass to the script itself.
#   outlook-call.ps1 --batch <requests.json>   # [{ "script": "find", "args": "-Days 1" }, ...]
#   outlook-call.ps1 --ping | --stop
#
# Examples:
#   & ./outlook-call.ps1 find -Days 7 -UnreadOnly
#   & ./outlook-call.ps1 read-body -EntryID "00000000..."
#
# Output and exit code are the same as running the script directly. The host starts
# on first use and exits after 30 idle minutes. If the host cannot be reached, the
# script runs directly in this shell instead. Under Windows PowerShell 5.1, which cannot
# open the current-user-only pipe, scripts always run directly.

$ErrorActionPreference = "Stop"
$pipeName = "outlook-suite-$($env:USERNAME)"
$here = $PSScriptRoot
$canUseHost = $PSVersionTable.PSVersion.Major -ge 7

function ConvertTo-ArgToken($t) {
    if ($t -is [bool]) { if ($t) { return '$true' } else { return '$false' } }
    if ($t -is [int] -or $t -is [long] -or $t -is [double] -or $t -is [decimal]) { return [string]$t }
    if ($t -is [array]) { return (@($t | ForEach-Object { "'" + ([string]$_).Replace("'", "''") + "'" }) -join ',') }
    return "'" + ([string]$t).Replace("'", "''") + "'"
}

# Rebuild a PowerShell argument line from $args, keeping parameter names, switches,
# "-Name:value" pairs, booleans, numbers and arrays exactly as typed.
function ConvertTo-ArgLine([object[]]$tokens) {
    $parts = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $t = $tokens[$i]
        if ($t -is [string] -and $t -match '^-[A-Za-z][A-Za-z0-9_]*:$' -and $i + 1 -lt $tokens.Count) {
            $parts.Add($t + (ConvertTo-ArgToken $tokens[$i + 1])); $i++; continue
        }
        if ($t -is [string] -and $t -match '^-[A-Za-z][A-Za-z0-9_]*$') { $parts.Add($t); continue }
        $parts.Add((ConvertTo-ArgToken $t))
    }
    return ($parts -join ' ')
}

function Resolve-ScriptPath([string]$name) {
    if (Test-Path -LiteralPath $name) { return (Resolve-Path -LiteralPath $name).Path }
    $n = $name
    if ($n -notlike "*.ps1") { $n = "$n.ps1" }
    if ($n -notlike "outlook-*") { $n = "outlook-$n" }
    $p = Join-Path $here $n
    if (Test-Path -LiteralPath $p) { return $p }
    throw "Unknown script '$name'. Scripts live in $here."
}

function Send-HostRequest($request, [int]$timeoutMs) {
    $client = New-Object System.IO.Pipes.NamedPipeClientStream(".", $pipeName,
        [System.IO.Pipes.PipeDirection]::InOut, [System.IO.Pipes.PipeOptions]::CurrentUserOnly)
    try {
        $client.Connect($timeoutMs)
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($client, $utf8, 65536, $true)
        $reader = New-Object System.IO.StreamReader($client, $utf8, $false, 65536, $true)
        $writer.WriteLine(($request | ConvertTo-Json -Depth 6 -Compress))
        $writer.Flush()
        $line = $reader.ReadLine()
        if (-not $line) { throw "empty response from host" }
        return ($line | ConvertFrom-Json)
    } finally { $client.Dispose() }
}

function Test-HostPipe { return [System.IO.Directory]::GetFiles("\\.\pipe\") -contains "\\.\pipe\$pipeName" }

function Start-OutlookHost {
    $hostScript = Join-Path $here "outlook-host.ps1"
    Start-Process -FilePath (Get-Process -Id $PID).Path -WindowStyle Hidden `
        -ArgumentList @("-NoProfile", "-NonInteractive", "-File", $hostScript) | Out-Null
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        if (Test-HostPipe) { return $true }
        Start-Sleep -Milliseconds 100
    }
    return $false
}

function Invoke-ViaHost($request) {
    if (-not $canUseHost) { throw "the host needs PowerShell 7 (pwsh)" }
    if (-not (Test-HostPipe)) {
        if (-not (Start-OutlookHost)) { throw "host did not start" }
    }
    # Connect waits while the host is busy with another request (queue).
    return Send-HostRequest $request 180000
}

$first = if ($args.Count -gt 0) { [string]$args[0] } else { "" }
if (-not $first) {
    Write-Host "Usage: outlook-call.ps1 <script> [arguments]   e.g. outlook-call.ps1 find -Days 7"
    exit 1
}

if ($first -in "--ping", "--stop", "--batch" -and -not $canUseHost) {
    Write-Output "Error: $first needs PowerShell 7 (pwsh). Run: pwsh -File outlook-call.ps1 $first ..."; exit 1
}
if ($first -eq "--ping") {
    if (-not (Test-HostPipe)) { Write-Output "host not running"; exit 1 }
    $r = Send-HostRequest @{ op = "ping" } 5000; Write-Output $r.output; exit 0
}
if ($first -eq "--stop") {
    if (-not (Test-HostPipe)) { Write-Output "host not running"; exit 0 }
    $r = Send-HostRequest @{ op = "stop" } 30000; Write-Output $r.output; exit 0
}
if ($first -eq "--batch") {
    $file = [string]$args[1]
    $items = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    $batch = @(foreach ($i in $items) {
        @{ path = (Resolve-ScriptPath $i.script); argline = [string]$i.args; cwd = (Get-Location).Path }
    })
    $resp = Invoke-ViaHost @{ batch = $batch }
    $n = 0
    foreach ($r in $resp.results) {
        $n++
        Write-Output "===== [$n] $($items[$n - 1].script) $($items[$n - 1].args)  (exit $($r.exitCode)) ====="
        Write-Output $r.output
    }
    exit ([int]$resp.exitCode)
}

$path = Resolve-ScriptPath $first
$argLine = ConvertTo-ArgLine @($args | Select-Object -Skip 1)
$request = @{ path = $path; argline = $argLine; cwd = (Get-Location).Path }

try {
    $resp = Invoke-ViaHost $request
} catch {
    # Fallback: run directly in this shell (silently on 5.1, where that is the normal path).
    if ($canUseHost) { Write-Host "(host unavailable: $($_.Exception.Message); running directly)" -ForegroundColor DarkGray }
    $global:LASTEXITCODE = 0
    Invoke-Expression "& '$($path.Replace("'", "''"))' $argLine"
    exit ([int]$global:LASTEXITCODE)
}
if ($resp.output) { Write-Output $resp.output }
exit ([int]$resp.exitCode)
