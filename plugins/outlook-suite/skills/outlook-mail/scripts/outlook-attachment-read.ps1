param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [int]$Days = 7,
    [int]$AttachmentIndex = 0,
    [string]$FilePath = "",
    [int]$MaxChars = 8000,
    [int]$MaxRows = 200,
    [int]$TimeoutSec = 90,
    [switch]$IncludeInline,
    [switch]$Json
)

# Outlook Attachment Read: the TEXT of an email's attachments (or of a local file), so an agent
# can triage, summarize or prep a meeting without saving and opening files by hand. Read-only.
#
#   .docx .xlsx .pptx     read straight from the file's XML: no Office app is started (fast)
#   .pdf .doc .rtf .odt   Word, hidden, read-only, macros forced off
#   .xls  / .ppt          Excel / PowerPoint, the same way
#   .msg                  the attached email's headers and body, through Outlook
#   .txt .csv .md .html … read directly (HTML tags stripped)
# Images, archives and other binaries are listed as not readable. Each attachment is copied to
# a private temp folder that is deleted afterwards. Text is cut at -MaxChars (0 = no limit);
# spreadsheets at -MaxRows rows per sheet. An Office read that hangs is ended after -TimeoutSec.
#
# Usage: .\outlook-attachment-read.ps1 -EntryID "00000000..."                 # every attachment
#        .\outlook-attachment-read.ps1 -EntryID "00000000..." -AttachmentIndex 2 -MaxChars 0
#        .\outlook-attachment-read.ps1 -FilePath "C:\Users\me\Dropbox\proposal.docx"
#        .\outlook-attachment-read.ps1 -Index 1 -Json                          # newest Inbox email

. (Join-Path $PSScriptRoot "_common.ps1")
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
$PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
$TextExt = @(".txt", ".csv", ".tsv", ".md", ".json", ".xml", ".log", ".ics", ".vcf", ".htm", ".html", ".yaml", ".yml", ".py", ".r", ".ps1", ".sql", ".tex", ".bib")
$WordExt = @(".pdf", ".doc", ".docm", ".rtf", ".odt", ".dot", ".dotx", ".wpd")

# ---------------------------------------------------------------------------
# Office Open XML (no Office needed)
# ---------------------------------------------------------------------------

function Read-ZipEntry($Zip, [string]$Name) {
    $e = $Zip.GetEntry($Name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open())
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

# Text of a WordprocessingML or DrawingML part: <t> runs, a tab per <tab>/<tc> end, a line per <p>/<br>.
function ConvertFrom-OoxmlText([string]$Xml) {
    if (-not $Xml) { return "" }
    $sb = New-Object System.Text.StringBuilder
    $settings = New-Object System.Xml.XmlReaderSettings; $settings.DtdProcessing = "Prohibit"
    $r = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($Xml)), $settings)
    $inText = $false
    $cellStart = New-Object System.Collections.Generic.Stack[int]
    try {
        while ($r.Read()) {
            switch ($r.NodeType) {
                "Element" {
                    switch ($r.LocalName) {
                        "t" { $inText = -not $r.IsEmptyElement }
                        "tc" { if (-not $r.IsEmptyElement) { $cellStart.Push($sb.Length) } else { [void]$sb.Append("`t") } }
                        "tab" { [void]$sb.Append("`t") }
                        "br" { [void]$sb.Append("`n") }
                        "cr" { [void]$sb.Append("`n") }
                    }
                }
                "Text" { if ($inText) { [void]$sb.Append($r.Value) } }
                "SignificantWhitespace" { if ($inText) { [void]$sb.Append($r.Value) } }
                "Whitespace" { if ($inText) { [void]$sb.Append($r.Value) } }
                "EndElement" {
                    # A table cell ends its last paragraph with a tab instead of a line;
                    # a row ends its last cell with a line.
                    switch ($r.LocalName) {
                        "t" { $inText = $false }
                        "p" { [void]$sb.Append("`n") }
                        "tc" {
                            $start = if ($cellStart.Count) { $cellStart.Pop() } else { 0 }
                            while ($sb.Length -gt $start -and $sb[$sb.Length - 1] -eq "`n") { [void]$sb.Remove($sb.Length - 1, 1) }
                            [void]$sb.Append("`t")
                        }
                        "tr" { while ($sb.Length -and $sb[$sb.Length - 1] -eq "`t") { [void]$sb.Remove($sb.Length - 1, 1) }; [void]$sb.Append("`n") }
                    }
                }
            }
        }
    }
    finally { $r.Dispose() }
    return ($sb.ToString() -replace "[ \t]+`n", "`n" -replace "`n{3,}", "`n`n").Trim()
}

function Read-Docx([string]$File) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($File)
    try {
        $parts = @("word/document.xml") + @($z.Entries | Where-Object { $_.FullName -match '^word/(footnotes|endnotes)\.xml$' } | ForEach-Object { $_.FullName })
        return (@($parts | ForEach-Object { ConvertFrom-OoxmlText (Read-ZipEntry $z $_) } | Where-Object { $_ }) -join "`n`n")
    }
    finally { $z.Dispose() }
}

function Read-Pptx([string]$File) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($File)
    try {
        $num = { [int]([regex]::Match($_.FullName, '(\d+)\.xml$').Groups[1].Value) }
        $slides = @($z.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object $num)
        $out = foreach ($s in $slides) {
            $n = [int]([regex]::Match($s.FullName, '(\d+)\.xml$').Groups[1].Value)
            $text = ConvertFrom-OoxmlText (Read-ZipEntry $z $s.FullName)
            $notes = ConvertFrom-OoxmlText (Read-ZipEntry $z "ppt/notesSlides/notesSlide$n.xml")
            # Notes pages carry the slide number as a text box; drop number-only lines.
            $notes = ($notes -replace '(?m)^\s*\d+\s*$', '').Trim()
            "--- Slide $n ---`n$text$(if ($notes) { "`n[Notes] $notes" })"
        }
        return ($out -join "`n`n")
    }
    finally { $z.Dispose() }
}

# Column letters (A, AB) of a cell reference -> 0-based index.
function Get-ColumnIndex([string]$Ref) {
    $letters = [regex]::Match($Ref, '^[A-Z]+').Value; $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return [Math]::Max($n - 1, 0)
}

function Read-Xlsx([string]$File, [int]$RowLimit) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($File)
    try {
        $shared = New-Object System.Collections.Generic.List[string]
        $ssXml = Read-ZipEntry $z "xl/sharedStrings.xml"
        if ($ssXml) { foreach ($si in ([xml]$ssXml).sst.si) { $shared.Add((@($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join "")) } }
        $wb = [xml](Read-ZipEntry $z "xl/workbook.xml")
        $rels = [xml](Read-ZipEntry $z "xl/_rels/workbook.xml.rels")
        $target = @{}; foreach ($rel in $rels.Relationships.Relationship) { $target[$rel.Id] = $rel.Target }
        $rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        $out = foreach ($sheet in $wb.workbook.sheets.sheet) {
            $t = $target[$sheet.GetAttribute("id", $rNs)]
            if (-not $t) { continue }
            $path = if ($t.StartsWith("/")) { $t.TrimStart("/") } else { "xl/" + $t }
            $sx = Read-ZipEntry $z $path
            if (-not $sx) { continue }
            $rowsText = New-Object System.Collections.Generic.List[string]; $total = 0
            foreach ($row in @(([xml]$sx).worksheet.sheetData.row)) {
                if (-not $row) { continue }
                $total++
                if ($rowsText.Count -ge $RowLimit) { continue }
                $cells = @{}; $max = -1
                foreach ($c in @($row.c)) {
                    if (-not $c) { continue }
                    $i = if ($c.r) { Get-ColumnIndex $c.r } else { $max + 1 }
                    $v = switch ($c.t) {
                        "s" { $shared[[int]$c.v] }
                        "inlineStr" { (@($c.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join "") }
                        "b" { if ($c.v -eq "1") { "TRUE" } else { "FALSE" } }
                        default { [string]$c.v }
                    }
                    $cells[$i] = $v; if ($i -gt $max) { $max = $i }
                }
                if ($max -lt 0) { continue }
                $rowsText.Add((@(0..$max | ForEach-Object { [string]$cells[$_] }) -join "`t").TrimEnd("`t"))
            }
            $more = if ($total -gt $rowsText.Count) { "`n[... $($total - $rowsText.Count) more rows; raise -MaxRows]" } else { "" }
            "--- Sheet: $($sheet.name) ($total rows) ---`n$($rowsText -join "`n")$more"
        }
        return ($out -join "`n`n")
    }
    finally { $z.Dispose() }
}

# ---------------------------------------------------------------------------
# Office apps (legacy formats and PDF): hidden, read-only, macros off. Each read gets a
# watchdog that ends the Office process this script started if it hangs (a hidden dialog
# would otherwise block forever). An Office app that was already running (PowerPoint is
# single-instance) is never quit; only the opened file is closed.
# ---------------------------------------------------------------------------

function Invoke-OfficeApp([string]$ProgId, [string]$ProcessName, [scriptblock]$Body) {
    $before = @(Get-Process $ProcessName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $app = New-Object -ComObject $ProgId
    $mine = @(Get-Process $ProcessName -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    $dog = $null
    if ($mine.Count -eq 1) {
        $dog = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList "-NoProfile", "-Command", "Start-Sleep $TimeoutSec; Stop-Process -Id $($mine[0].Id) -Force -ErrorAction SilentlyContinue"
    }
    # Macros off for this read (msoAutomationSecurityForceDisable); the old value goes back after,
    # which matters when the app was already running (PowerPoint is single-instance).
    $oldSecurity = $null; try { $oldSecurity = $app.AutomationSecurity } catch { }
    try { $app.AutomationSecurity = 3; return (& $Body $app) }
    catch {
        if ($mine.Count -and -not (Get-Process -Id $mine[0].Id -ErrorAction SilentlyContinue)) { throw "$ProcessName did not finish within $TimeoutSec s (raise -TimeoutSec)." }
        throw
    }
    finally {
        if ($null -ne $oldSecurity) { try { $app.AutomationSecurity = $oldSecurity } catch { } }
        if ($mine.Count) { try { $app.Quit() } catch { } }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        if ($dog) { Stop-Process -Id $dog.Id -Force -ErrorAction SilentlyContinue }
        # An app this script started must not linger (Excel can, while COM references drain).
        if ($mine.Count) {
            $p = Get-Process -Id $mine[0].Id -ErrorAction SilentlyContinue
            if ($p -and -not $p.WaitForExit(2000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Read-WithWord([string]$File) {
    Invoke-OfficeApp "Word.Application" "WINWORD" {
        param($w)
        $w.Visible = $false; $w.DisplayAlerts = 0
        # Word asks before converting a PDF even with alerts off (tested: without this a PDF
        # hangs until the watchdog). Turn that one prompt off for this open only and put the
        # user's setting back. If this process is killed mid-open the value stays 1, which only
        # silences that prompt in Word.
        $key = "HKCU:\Software\Microsoft\Office\$($w.Version)\Word\Options"; $name = "DisableConvertPdfWarning"
        $old = (Get-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue).$name
        $isPdf = [System.IO.Path]::GetExtension($File) -eq ".pdf"
        if ($isPdf) { if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }; Set-ItemProperty -Path $key -Name $name -Value 1 -Type DWord }
        try {
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, ..., Visible = 12th).
            # A dummy password makes a protected file fail at once instead of waiting on a prompt.
            $m = [Type]::Missing
            $doc = $w.Documents.Open($File, $false, $true, $false, "x", $m, $m, $m, $m, $m, $m, $false)
        }
        finally {
            if ($isPdf) { if ($null -eq $old) { Remove-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue } else { Set-ItemProperty -Path $key -Name $name -Value $old -Type DWord } }
        }
        try { [string]$doc.Content.Text } finally { $doc.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    }
}

function Read-WithExcel([string]$File, [int]$RowLimit) {
    Invoke-OfficeApp "Excel.Application" "EXCEL" {
        param($x)
        $x.Visible = $false; $x.DisplayAlerts = $false; $x.AskToUpdateLinks = $false
        $wb = $x.Workbooks.Open($File, 0, $true, [Type]::Missing, "x")   # UpdateLinks 0, ReadOnly, dummy password (no prompt)
        try {
            $out = foreach ($ws in $wb.Worksheets) {
                $ur = $ws.UsedRange; $rows = [Math]::Min($ur.Rows.Count, $RowLimit); $cols = $ur.Columns.Count
                $lines = for ($r = 1; $r -le $rows; $r++) { (@(for ($c = 1; $c -le $cols; $c++) { [string]$ur.Cells.Item($r, $c).Text }) -join "`t").TrimEnd("`t") }
                $more = if ($ur.Rows.Count -gt $rows) { "`n[... $($ur.Rows.Count - $rows) more rows; raise -MaxRows]" } else { "" }
                "--- Sheet: $($ws.Name) ($($ur.Rows.Count) rows) ---`n$($lines -join "`n")$more"
            }
            $out -join "`n`n"
        }
        finally { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    }
}

function Read-WithPowerPoint([string]$File) {
    Invoke-OfficeApp "PowerPoint.Application" "POWERPNT" {
        param($p)
        $pres = $p.Presentations.Open($File, -1, 0, 0)   # ReadOnly, not Untitled, no window
        try {
            $out = foreach ($s in $pres.Slides) {
                $texts = foreach ($sh in $s.Shapes) { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $sh.TextFrame.TextRange.Text } }
                "--- Slide $($s.SlideIndex) ---`n$(($texts -join "`n").Trim())"
            }
            $out -join "`n`n"
        }
        finally { $pres.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
    }
}

function Read-Msg([string]$File) {
    if (-not $Namespace) { $null = Connect-Outlook }
    $m = $Namespace.OpenSharedItem($File)
    try {
        $head = @("Subject: $($m.Subject)")
        try { $head += "From: $($m.SenderName) <$($m.SenderEmailAddress)>"; $head += "Date: $(Format-MailDate $m.SentOn)"; $head += "To: $($m.To)" } catch { }
        return ($head -join "`n") + "`n`n" + [string]$m.Body
    }
    finally {
        # olDiscard, then let go so Outlook releases its lock on the .msg file.
        try { $m.Close(1) } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m) } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function ConvertFrom-HtmlText([string]$Html) {
    $t = $Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', '' -replace '(?i)<br\s*/?>|</(p|div|li|tr|h\d)>', "`n" -replace '<[^>]+>', ''
    return ([System.Net.WebUtility]::HtmlDecode($t) -replace "[ \t]+`n", "`n" -replace "`n{3,}", "`n`n").Trim()
}

# One file -> @{ Text; Method } or @{ Error }.
function Read-DocumentText([string]$File) {
    $ext = [System.IO.Path]::GetExtension($File).ToLowerInvariant()
    try {
        switch -Regex ($ext) {
            '^\.(docx|docm|dotx)$' { return @{ Text = Read-Docx $File; Method = "OOXML" } }
            '^\.(pptx|pptm)$' { return @{ Text = Read-Pptx $File; Method = "OOXML" } }
            '^\.(xlsx|xlsm)$' { return @{ Text = Read-Xlsx $File $MaxRows; Method = "OOXML" } }
            '^\.msg$' { return @{ Text = Read-Msg $File; Method = "Outlook" } }
            '^\.(xls|xlsb)$' { return @{ Text = Read-WithExcel $File $MaxRows; Method = "Excel" } }
            '^\.(ppt|pps|ppsx)$' { return @{ Text = Read-WithPowerPoint $File; Method = "PowerPoint" } }
        }
        if ($WordExt -contains $ext) { return @{ Text = Read-WithWord $File; Method = "Word" } }
        if ($TextExt -contains $ext) {
            $raw = [System.IO.File]::ReadAllText($File)
            if ($ext -in ".htm", ".html") { $raw = ConvertFrom-HtmlText $raw }
            return @{ Text = $raw; Method = "Text" }
        }
        return @{ Error = "not readable as text ($ext)" }
    }
    catch { return @{ Error = "could not read: $($_.Exception.Message.Split("`n")[0])" } }
}

function Limit-Text([string]$Text) {
    $t = ($Text -replace "`r`n?", "`n" -replace [char]7, "" -replace [char]12, "`n").Trim()
    if ($MaxChars -gt 0 -and $t.Length -gt $MaxChars) { return @{ Text = $t.Substring(0, $MaxChars); Truncated = $true; Chars = $t.Length } }
    return @{ Text = $t; Truncated = $false; Chars = $t.Length }
}

# ---------------------------------------------------------------------------

try {
    $results = New-Object System.Collections.Generic.List[object]
    $header = $null
    # Every file is read from its own copy under a fresh temp path: the original is never locked,
    # and Word's "it caused a serious error last time" prompt (keyed by path) cannot block a read.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("outlook-suite-read-" + [guid]::NewGuid().ToString("N"))
    if ($FilePath) {
        $full = (Resolve-Path -LiteralPath $FilePath -ErrorAction Stop).Path
        New-Item -ItemType Directory -Path $tmp | Out-Null
        try {
            $copy = Join-Path $tmp ("1-" + [System.IO.Path]::GetFileName($full))
            Copy-Item -LiteralPath $full -Destination $copy
            $res = Read-DocumentText $copy
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
        $rec = [ordered]@{ Index = 1; FileName = [System.IO.Path]::GetFileName($full); SizeKB = [math]::Round((Get-Item -LiteralPath $full).Length / 1KB, 1); Method = $res.Method; Error = $res.Error; Chars = 0; Truncated = $false; Text = "" }
        if (-not $res.Error) { $l = Limit-Text $res.Text; $rec.Text = $l.Text; $rec.Chars = $l.Chars; $rec.Truncated = $l.Truncated }
        $results.Add($rec)
        $header = [ordered]@{ File = $full }
    }
    else {
        $null = Connect-Outlook
        $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
        $html = ""; try { $html = $email.HTMLBody } catch { }
        $count = $email.Attachments.Count
        if ($AttachmentIndex -gt $count) { Stop-WithError "Attachment index $AttachmentIndex is out of range; the email has $count attachment(s)." }
        New-Item -ItemType Directory -Path $tmp | Out-Null
        try {
            for ($i = 1; $i -le $count; $i++) {
                if ($AttachmentIndex -gt 0 -and $i -ne $AttachmentIndex) { continue }
                $att = $email.Attachments.Item($i)
                $cid = $null; $hidden = $false
                try { $cid = $att.PropertyAccessor.GetProperty($PR_CONTENT_ID) } catch { }
                try { $hidden = [bool]$att.PropertyAccessor.GetProperty($PR_HIDDEN) } catch { }
                $inline = $hidden -or ($cid -and $html -match [regex]::Escape("cid:$cid"))
                if ($inline -and -not $IncludeInline -and $AttachmentIndex -eq 0) { continue }
                $rec = [ordered]@{ Index = $i; FileName = [string]$att.FileName; SizeKB = [math]::Round($att.Size / 1KB, 1); Method = $null; Error = $null; Chars = 0; Truncated = $false; Text = "" }
                if ($att.Type -ne 1) {
                    # 5 = embedded Outlook item, 6 = OLE object; only by-value files can be saved.
                    if ($att.Type -eq 5) { if ($rec.FileName -notmatch '\.msg$') { $rec.FileName = "$($rec.FileName).msg" } } else { $rec.Error = "not a file attachment (type $($att.Type))"; $results.Add($rec); continue }
                }
                $safe = ($rec.FileName -replace '[\\/:*?"<>|]', '_'); if (-not $safe) { $safe = "attachment$i" }
                $file = Join-Path $tmp ("$i-" + $safe)
                $att.SaveAsFile($file)
                $res = Read-DocumentText $file
                $rec.Method = $res.Method; $rec.Error = $res.Error
                if (-not $res.Error) { $l = Limit-Text $res.Text; $rec.Text = $l.Text; $rec.Chars = $l.Chars; $rec.Truncated = $l.Truncated }
                $results.Add($rec)
            }
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
        $header = [ordered]@{ EntryID = $email.EntryID; Subject = $email.Subject; From = (Get-SenderSmtp $email); Date = Format-MailDateIso $email.ReceivedTime }
    }

    if ($Json) {
        $o = $header; $o.MaxChars = $MaxChars; $o.Attachments = $results.ToArray()
        Write-Json $o
        exit 0
    }

    Write-Host "=== ATTACHMENT TEXT ===" -ForegroundColor Cyan
    if ($FilePath) { Write-Host "File: $($header.File)" }
    else {
        Write-Host "EntryID: $($header.EntryID)"
        Write-Host "Subject: $($header.Subject)" -ForegroundColor Yellow
        Write-Host "From: $($header.From)"
        Write-Host "Date: $(Format-MailDate $email.ReceivedTime)"
    }
    if ($results.Count -eq 0) { Write-Host "No file attachments$(if (-not $IncludeInline) { ' (inline images skipped; add -IncludeInline)' })." -ForegroundColor Gray; exit 0 }
    foreach ($r in $results) {
        Write-Host ""
        Write-Host "--- [$($r.Index)] $($r.FileName) ($($r.SizeKB) KB)$(if ($r.Method) { ", via $($r.Method)" }) ---" -ForegroundColor Cyan
        if ($r.Error) { Write-Host "  ($($r.Error))" -ForegroundColor DarkGray; continue }
        if (-not $r.Text) { Write-Host "  (no text found; it may be a scan or images only)" -ForegroundColor DarkGray; continue }
        Write-Output $r.Text
        if ($r.Truncated) { Write-Host "[... cut at $MaxChars of $($r.Chars) characters; -MaxChars 0 for all]" -ForegroundColor DarkGray }
    }
}
catch {
    Stop-WithError $_
}
