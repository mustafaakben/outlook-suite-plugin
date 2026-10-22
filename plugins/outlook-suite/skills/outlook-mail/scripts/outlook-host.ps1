# Outlook Suite warm host
#
# Keeps one PowerShell runspace and one Outlook COM session open and runs mail scripts
# on request, so each call skips PowerShell startup, the Outlook connection and folder
# lookups. Requests come over a local named pipe that only the current Windows user can
# open. Requests are handled one at a time: Outlook's object model serves calls on a
# single thread, so queuing here is the safe way to let many agents share it.
#
# You normally never start this yourself: outlook-call.ps1 starts it on first use.
# It exits after -IdleMinutes without requests.
#
# Usage: pwsh -NoProfile -File outlook-host.ps1 [-IdleMinutes 30]
# Stop:  pwsh -NoProfile -File outlook-call.ps1 --stop

param(
    [int]$IdleMinutes = 30
)

$ErrorActionPreference = "Stop"
# The pipe's current-user-only protection exists in PowerShell 7 (.NET 5+) only.
if ($PSVersionTable.PSVersion.Major -lt 7) { Write-Output "Error: the host needs PowerShell 7 (pwsh)."; exit 1 }
$pipeName = "outlook-suite-$($env:USERNAME)"
$logDir = Join-Path $env:LOCALAPPDATA "outlook-suite"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir "host.log"

function Write-Log([string]$msg) {
    try { Add-Content -Path $logPath -Value "$((Get-Date).ToString('s')) [$PID] $msg" } catch { }
}

# One host per user.
$mutex = New-Object System.Threading.Mutex($false, "Local\outlook-suite-host-$($env:USERNAME)")
if (-not $mutex.WaitOne(0)) { Write-Log "another host is running; exiting"; exit 0 }

# Persistent runspace that keeps the Outlook session in $global:OutlookSuite.
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
$runspace = [runspacefactory]::CreateRunspace($iss)
$runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
$runspace.Open()
$init = [powershell]::Create(); $init.Runspace = $runspace
[void]$init.AddScript('$global:OutlookSuite = @{ InHost = $true; App = $null; Namespace = $null; FolderCache = @{} }').Invoke()
$init.Dispose()

# Runs one request and returns @{ exitCode; output }. The script runs exactly as
# `& <path> <argline>` would, and every stream is rendered to text in order.
$runner = @'
param($Path, $ArgLine, $Cwd)
if ($Cwd -and (Test-Path -LiteralPath $Cwd)) { Set-Location -LiteralPath $Cwd }
$global:LASTEXITCODE = 0
$lines = New-Object System.Collections.Generic.List[string]
$code = 0
$partial = ""   # Write-Host -NoNewline text, joined with what follows on the same line
try {
    $sb = [scriptblock]::Create("& '$($Path.Replace("'", "''"))' $ArgLine *>&1")
    foreach ($rec in (& $sb)) {
        if ($rec -is [System.Management.Automation.InformationRecord]) {
            $m = $rec.MessageData
            if ($m -is [System.Management.Automation.HostInformationMessage] -and $m.NoNewLine) { $partial += [string]$m.Message; continue }
            $lines.Add($partial + [string]$m); $partial = ""; continue
        }
        if ($partial) { $lines.Add($partial); $partial = "" }
        if ($rec -is [System.Management.Automation.ErrorRecord]) { $lines.Add("Error: $($rec.Exception.Message)") }
        elseif ($rec -is [System.Management.Automation.WarningRecord]) { $lines.Add("WARNING: $($rec.Message)") }
        elseif ($rec -is [string]) { $lines.Add($rec) }
        elseif ($null -ne $rec) { foreach ($l in ($rec | Out-String -Stream)) { $lines.Add($l) } }
    }
    if ($partial) { $lines.Add($partial) }
    $code = [int]$global:LASTEXITCODE
} catch {
    $lines.Add("Error: $($_.Exception.Message)")
    $code = 1
}
@{ exitCode = $code; output = ($lines -join "`n") }
'@

function Invoke-Request($req) {
    $path = [string]$req.path
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        return @{ exitCode = 1; output = "Error: script not found: $path" }
    }
    $ps = [powershell]::Create(); $ps.Runspace = $runspace
    [void]$ps.AddScript($runner).AddArgument($path).AddArgument([string]$req.argline).AddArgument([string]$req.cwd)
    try {
        $res = $ps.Invoke()
        if ($res.Count -gt 0) { return $res[$res.Count - 1] }
        return @{ exitCode = 1; output = "Error: no result" }
    } catch {
        return @{ exitCode = 1; output = "Error: $($_.Exception.Message)" }
    } finally { $ps.Dispose() }
}

Write-Log "host started (pipe $pipeName, idle $IdleMinutes min)"
$idleMs = [int]([TimeSpan]::FromMinutes($IdleMinutes).TotalMilliseconds)
$utf8 = New-Object System.Text.UTF8Encoding($false)

try {
    while ($true) {
        $server = New-Object System.IO.Pipes.NamedPipeServerStream(
            $pipeName,
            [System.IO.Pipes.PipeDirection]::InOut,
            1,
            [System.IO.Pipes.PipeTransmissionMode]::Byte,
            ([System.IO.Pipes.PipeOptions]::Asynchronous -bor [System.IO.Pipes.PipeOptions]::CurrentUserOnly))
        $wait = $server.WaitForConnectionAsync()
        if (-not $wait.Wait($idleMs)) {
            Write-Log "idle for $IdleMinutes min; exiting"
            $server.Dispose()
            break
        }
        $stop = $false
        try {
            $reader = New-Object System.IO.StreamReader($server, $utf8, $false, 65536, $true)
            $writer = New-Object System.IO.StreamWriter($server, $utf8, 65536, $true)
            $line = $reader.ReadLine()
            $req = $line | ConvertFrom-Json
            $sw = [Diagnostics.Stopwatch]::StartNew()
            if ($req.op -eq "stop") {
                $resp = @{ exitCode = 0; output = "Outlook Suite host stopped." }
                $stop = $true
            } elseif ($req.op -eq "ping") {
                $resp = @{ exitCode = 0; output = "pong $PID" }
            } elseif ($req.batch) {
                $results = @(foreach ($r in $req.batch) { Invoke-Request $r })
                $resp = @{ exitCode = (($results | ForEach-Object { $_.exitCode } | Measure-Object -Maximum).Maximum); results = $results }
            } else {
                $resp = Invoke-Request $req
            }
            Write-Log "$(if ($req.path) { Split-Path $req.path -Leaf } else { $req.op }) exit=$($resp.exitCode) $($sw.ElapsedMilliseconds)ms"
            $writer.WriteLine(($resp | ConvertTo-Json -Depth 6 -Compress))
            $writer.Flush()
        } catch {
            Write-Log "request failed: $($_.Exception.Message)"
        } finally {
            $server.Dispose()
        }
        if ($stop) { break }
    }
} finally {
    $runspace.Dispose()
    $mutex.ReleaseMutex(); $mutex.Dispose()
    Write-Log "host exited"
}
