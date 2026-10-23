<#
.SYNOPSIS
  Live test for outlook-everything.ps1 (Windows Search across mail, calendar and files) and
  outlook-attachment-read.ps1 (attachment and file text).

.DESCRIPTION
  1. Builds fixture files with Word, Excel and PowerPoint (docx, doc, pdf, xlsx, xls, pptx, ppt)
     plus txt, html and msg, all carrying known marker text, and reads each with -FilePath.
  2. Attaches them to a "[suite-test]" draft in Deleted Items/_suite-test (never addressed,
     never sent) and reads them through -EntryID.
  3. Read-only search of the real index: groups, EntryIDs that open, dates, filters.
  Prints PASS/FAIL and counts only. Deletes everything it creates.
#>
param(
    [ValidateSet("pwsh", "powershell")]
    [string]$Exe = "pwsh",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-mail/scripts")
)
$ErrorActionPreference = "Stop"
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
$Application = Connect-Outlook
Write-Host "Running under $Exe"

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green } else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
function Run([string]$script, [object[]]$a) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $out = & $Exe -NoProfile -File (Join-Path $ScriptsDir $script) @a 2>&1 | Out-String
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out; Ms = $sw.ElapsedMilliseconds }
}
function JsonOf($r) { try { return ($r.Out | ConvertFrom-Json) } catch { return $null } }
function Get-OfficeCount { return @(Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue).Count }

$MarkA = "suitefixturealpha"; $MarkB = "suitefixturebeta"; $MarkN = "suitefixturenotes"
$tmp = Join-Path $env:TEMP ("suite-read-test-" + [guid]::NewGuid().ToString("N").Substring(0, 8)); New-Item -ItemType Directory -Path $tmp | Out-Null
$sandbox = $null; foreach ($f in $Namespace.GetDefaultFolder(3).Folders) { if ($f.Name -eq "_suite-test") { $sandbox = $f } }
if (-not $sandbox) { $sandbox = $Namespace.GetDefaultFolder(3).Folders.Add("_suite-test") }
$sentBefore = $Namespace.GetDefaultFolder(5).Items.Count + $Namespace.GetDefaultFolder(4).Items.Count
$prevGate = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM; $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"
$draft = $null

try {
    # Office fixtures need Word, Excel and PowerPoint; without them only the search part runs.
    $hasOffice = @("Word.Application", "Excel.Application", "PowerPoint.Application" | Where-Object { -not (Test-Path "Registry::HKEY_CLASSES_ROOT\$_\CLSID") }).Count -eq 0
    if (-not $hasOffice) { Write-Host "SKIP  attachment-read fixtures: Word, Excel or PowerPoint is not installed" -ForegroundColor Yellow }
    else {
        # --- 1. Fixtures ---
        $officePids = @(Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
        try {
            $d = $w.Documents.Add()
            $d.Content.Text = "Heading $MarkA`rSecond paragraph $MarkB"; $d.Content.InsertParagraphAfter()
            $t = $d.Tables.Add($d.Paragraphs.Last.Range, 2, 2); $t.Cell(1, 1).Range.Text = "cellone"; $t.Cell(2, 2).Range.Text = "celltwo"
            $d.SaveAs2((Join-Path $tmp "fx.docx"), 16); $d.SaveAs2((Join-Path $tmp "fx.doc"), 0); $d.SaveAs2((Join-Path $tmp "fx.pdf"), 17); $d.Close(0)
        }
        finally { $w.Quit(0) }
        $x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
        try {
            $wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = "Budget"
            $ws.Cells.Item(1, 1).Value2 = "Item"; $ws.Cells.Item(1, 3).Value2 = $MarkA; $ws.Cells.Item(2, 1).Value2 = $MarkB; $ws.Cells.Item(2, 2).Value2 = 42
            for ($i = 3; $i -le 30; $i++) { $ws.Cells.Item($i, 1).Value2 = "row$i" }
            $wb.SaveAs((Join-Path $tmp "fx.xlsx"), 51); $wb.SaveAs((Join-Path $tmp "fx.xls"), 56); $wb.Close($false)
        }
        finally { $x.Quit() }
        $p = New-Object -ComObject PowerPoint.Application
        try {
            $pr = $p.Presentations.Add(0); $s = $pr.Slides.Add(1, 2)   # ppLayoutText
            $s.Shapes.Item(1).TextFrame.TextRange.Text = "Slide $MarkA"; $s.Shapes.Item(2).TextFrame.TextRange.Text = "Bullet $MarkB"
            $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Speaker $MarkN"
            $pr.SaveAs((Join-Path $tmp "fx.pptx"), 24); $pr.SaveAs((Join-Path $tmp "fx.ppt"), 1); $pr.Close()
        }
        finally { $p.Quit() }
        [IO.File]::WriteAllText((Join-Path $tmp "fx.txt"), "plain $MarkA")
        [IO.File]::WriteAllText((Join-Path $tmp "fx.html"), "<html><style>.x{}</style><body><p>web $MarkA</p><p>&amp; $MarkB</p></body></html>")
        [IO.File]::WriteAllBytes((Join-Path $tmp "fx.bin"), [byte[]](1..64))
        $m = $Application.CreateItem(0); $m.Subject = "[suite-test] inner $MarkA"; $m.Body = "inner body $MarkB"; $m.SaveAs((Join-Path $tmp "fx.msg"), 9)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
        # Let go of the fixture builders' COM objects so their Office processes exit; stop any
        # that this test started and that still linger.
        $d = $t = $wb = $ws = $pr = $s = $w = $x = $p = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        for ($i = 0; $i -lt 10 -and @(Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | Where-Object { $officePids -notcontains $_.Id }).Count; $i++) { Start-Sleep -Seconds 1 }
        Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue | Where-Object { $officePids -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
        $officeBefore = Get-OfficeCount
        Check "fixtures built (11 files)" (@(Get-ChildItem $tmp).Count -eq 11) "files=$(@(Get-ChildItem $tmp).Count)"

        # --- 2. -FilePath on each type ---
        $expect = [ordered]@{ "fx.docx" = "OOXML"; "fx.xlsx" = "OOXML"; "fx.pptx" = "OOXML"; "fx.doc" = "Word"; "fx.pdf" = "Word"; "fx.xls" = "Excel"; "fx.ppt" = "PowerPoint"; "fx.txt" = "Text"; "fx.html" = "Text"; "fx.msg" = "Outlook" }
        foreach ($name in $expect.Keys) {
            $r = Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp $name), "-Json"); $j = JsonOf $r
            $a = if ($j) { @($j.Attachments)[0] } else { $null }
            $ok = $r.Exit -eq 0 -and $a -and -not $a.Error -and $a.Method -eq $expect[$name] -and $a.Text -match $MarkA
            if ($name -notin "fx.txt") { $ok = $ok -and $a.Text -match $MarkB }
            Check "-FilePath $name via $($expect[$name]) has the marker text ($($r.Ms) ms)" $ok "exit=$($r.Exit) method=$($a.Method) err=$($a.Error)"
        }
        $j = JsonOf (Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.docx"), "-Json")); $t = @($j.Attachments)[0].Text
        Check "docx table rows come out as tab-separated cells" ($t -match "(?m)^cellone`t*$" -and $t -match "(?m)^`tcelltwo$")
        $j = JsonOf (Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.pptx"), "-Json")); $t = @($j.Attachments)[0].Text
        Check "pptx has slide header and speaker notes" ($t -match "--- Slide 1 ---" -and $t -match "\[Notes\].*$MarkN")
        $j = JsonOf (Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.xlsx"), "-MaxRows", "5", "-Json")); $t = @($j.Attachments)[0].Text
        Check "xlsx keeps sheet name, column gaps and the row cap" ($t -match "--- Sheet: Budget \(30 rows\) ---" -and $t -match "Item`t`t$MarkA" -and $t -match "$MarkB`t42" -and $t -match "25 more rows")
        $j = JsonOf (Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.html"), "-Json")); $t = @($j.Attachments)[0].Text
        Check "html: tags and styles stripped, entities decoded" ($t -notmatch "<|\.x\{" -and $t -match "& $MarkB")
        $j = JsonOf (Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.docx"), "-MaxChars", "10", "-Json")); $a = @($j.Attachments)[0]
        Check "-MaxChars cuts and reports the full length" ($a.Truncated -and $a.Text.Length -eq 10 -and $a.Chars -gt 10)
        $r = Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.bin"), "-Json"); $a = @((JsonOf $r).Attachments)[0]
        Check "binary file: reported as not readable, exit 0" ($r.Exit -eq 0 -and $a.Error -match "not readable")
        $r = Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "nope.docx"))
        Check "missing file: clear error, exit 1" ($r.Exit -eq 1 -and $r.Out -match "Error")
        $r = Run "outlook-attachment-read.ps1" @("-FilePath", (Join-Path $tmp "fx.docx"))
        Check "text mode prints the header and the text" ($r.Exit -eq 0 -and $r.Out -match "ATTACHMENT TEXT" -and $r.Out -match $MarkA)

        # --- 3. Attachments on a test draft in the sandbox (never addressed, never sent) ---
        $draft = $Application.CreateItem(0); $draft.Subject = "[suite-test] attachment read"; $draft.HTMLBody = "<p>test</p>"
        foreach ($name in "fx.docx", "fx.xlsx", "fx.pptx", "fx.pdf", "fx.msg", "fx.bin") { [void]$draft.Attachments.Add((Join-Path $tmp $name)) }
        $draft.Save(); $draft = $draft.Move($sandbox); $did = $draft.EntryID
        $tempBefore = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Directory -Filter "outlook-suite-read-*" -ErrorAction SilentlyContinue).Count
        $r = Run "outlook-attachment-read.ps1" @("-EntryID", $did, "-Json"); $j = JsonOf $r
        $att = @($j.Attachments)
        $good = @($att | Where-Object { -not $_.Error -and $_.Text -match $MarkA }).Count
        Check "email: 6 attachments, 5 read with marker text, the binary flagged ($($r.Ms) ms)" ($r.Exit -eq 0 -and $att.Count -eq 6 -and $good -eq 5 -and @($att | Where-Object { $_.Error -match "not readable" }).Count -eq 1) "count=$($att.Count) good=$good"
        $r = Run "outlook-attachment-read.ps1" @("-EntryID", $did, "-AttachmentIndex", "3", "-Json"); $j = JsonOf $r
        Check "-AttachmentIndex 3 reads just the pptx" ($r.Exit -eq 0 -and @($j.Attachments).Count -eq 1 -and @($j.Attachments)[0].FileName -eq "fx.pptx")
        $r = Run "outlook-attachment-read.ps1" @("-EntryID", $did, "-AttachmentIndex", "9")
        Check "-AttachmentIndex out of range: clear error" ($r.Exit -eq 1 -and $r.Out -match "out of range")
        $tempAfter = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Directory -Filter "outlook-suite-read-*" -ErrorAction SilentlyContinue).Count
        Check "temp copies are deleted after reading" ($tempAfter -le $tempBefore) "before=$tempBefore after=$tempAfter"
        Start-Sleep -Seconds 2
        Check "no Word, Excel or PowerPoint process left behind" ((Get-OfficeCount) -le $officeBefore) "before=$officeBefore after=$(Get-OfficeCount)"

    }

    # --- 4. Everything search (read-only on the real index) ---
    $r = Run "outlook-everything.ps1" @("meeting", "-Json"); $j = JsonOf $r
    $kinds = @($j.groups.PSObject.Properties.Name)
    Check "everything: all six kinds searched ($($r.Ms) ms, index $($j.elapsedMs) ms)" ($r.Exit -eq 0 -and ($kinds -join ",") -eq "Mail,Calendar,Contact,Task,Note,File") "kinds=$($kinds -join ',')"
    $mail = @($j.groups.Mail); $opened = 0; $dated = 0
    foreach ($h in ($mail | Where-Object { $_.EntryID } | Select-Object -First 5)) {
        $it = $Namespace.GetItemFromID($h.EntryID); $opened++
        if ([math]::Abs(($it.ReceivedTime - ([datetimeoffset]::Parse($h.Date)).LocalDateTime).TotalMinutes) -lt 2) { $dated++ }
    }
    Write-Host "      mail hits: $($mail.Count), opened: $opened, date agrees: $dated"
    Check "mail hits carry EntryIDs that open, with Outlook's received time" ($mail.Count -eq 0 -or ($opened -gt 0 -and $dated -eq $opened))
    $cal = @($j.groups.Calendar | Where-Object { $_.EntryID } | Select-Object -First 3)
    $calOk = @($cal | Where-Object { $Namespace.GetItemFromID($_.EntryID).MessageClass -like "IPM.Appointment*" -or $Namespace.GetItemFromID($_.EntryID).MessageClass -like "IPM.Schedule*" }).Count
    Check "calendar hits open as appointments" ($cal.Count -eq $calOk) "hits=$($cal.Count) ok=$calOk"
    $r = Run "outlook-everything.ps1" @("meeting", "-Kind", "Mail", "-Days", "30", "-Json"); $j = JsonOf $r
    $old = @($j.groups.Mail | Where-Object { ([datetimeoffset]::Parse($_.Date)).LocalDateTime -lt (Get-Date).AddDays(-30) }).Count
    Check "-Kind Mail -Days 30: only mail, nothing older" ($r.Exit -eq 0 -and @($j.groups.PSObject.Properties).Count -eq 1 -and $old -eq 0)
    $docs = [Environment]::GetFolderPath("MyDocuments")
    $r = Run "outlook-everything.ps1" @("meeting", "-Kind", "File", "-Path", $docs, "-Json"); $j = JsonOf $r
    $files = @($j.groups.File)
    if ($files.Count -eq 0) { Write-Host "      (no indexed file under Documents matched 'meeting'; the -Path check below is vacuous)" -ForegroundColor Yellow }
    Check "-Path keeps files inside that folder ($($files.Count) files)" ($r.Exit -eq 0 -and @($files | Where-Object { $_.Path -notlike "$docs\*" }).Count -eq 0)
    $r = Run "outlook-everything.ps1" @("meeting", "-Kind", "Mail", "-NoResolve", "-Json"); $j = JsonOf $r
    Check "-NoResolve skips Outlook lookups" ($r.Exit -eq 0 -and @($j.groups.Mail | Where-Object { $_.EntryID }).Count -eq 0)
    $r = Run "outlook-everything.ps1" @("it's", "-Kind", "Mail", "-Top", "2")
    Check "apostrophe in the query is escaped" ($r.Exit -eq 0)
    $r = Run "outlook-everything.ps1" @()
    Check "empty query: clear error" ($r.Exit -eq 1 -and $r.Out -match "words to search")
    $r = Run "outlook-everything.ps1" @("meeting", "-Top", "3")
    Check "text mode groups by kind" ($r.Exit -eq 0 -and $r.Out -match "=== EVERYTHING: meeting ===" -and $r.Out -match "Mail \(" -and $r.Out -match "File \(")
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $prevGate
    if ($draft) { try { $draft.Delete() } catch { } }
    $q = "@SQL=""urn:schemas:httpmail:subject"" LIKE '[suite-test]%'"
    foreach ($f in @($sandbox, $Namespace.GetDefaultFolder(16), $Namespace.GetDefaultFolder(3))) { try { foreach ($i in @($f.Items.Restrict($q))) { try { $i.Delete() } catch { } } } catch { } }
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}
$sentAfter = $Namespace.GetDefaultFolder(5).Items.Count + $Namespace.GetDefaultFolder(4).Items.Count
Check "nothing was sent: Sent Items + Outbox unchanged" ($sentAfter -le $sentBefore) "before=$sentBefore after=$sentAfter"
Write-Host ""
if ($script:fails -eq 0) { Write-Host "ALL PASSED" -ForegroundColor Green; exit 0 } else { Write-Host "$script:fails FAILED" -ForegroundColor Red; exit 1 }
