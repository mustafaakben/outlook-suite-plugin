<#
.SYNOPSIS
  Live test for the outlook-contacts scripts and outlook-people.ps1.

.DESCRIPTION
  Creates one disposable contact ("[suite-test] ..."), runs every contact script against it,
  checks the real Outlook state after each step and prints PASS/FAIL lines and timings only
  (no names, addresses or EntryIDs). Read-only checks of outlook-people.ps1 use the signed-in
  user's own address. Cleans up: the test contact is deleted and purged from Deleted Items.

  Run under PowerShell 7 (default) or Windows PowerShell 5.1:
    pwsh -File tests/Test-PeopleContacts.ps1
    pwsh -File tests/Test-PeopleContacts.ps1 -Shell powershell
#>
param(
    [ValidateSet("pwsh", "powershell")][string]$Shell = "pwsh",
    [string]$ScriptsDir = (Join-Path $PSScriptRoot "../plugins/outlook-suite/skills/outlook-contacts/scripts")
)

$ErrorActionPreference = "Stop"
if ($PSVersionTable.PSVersion.Major -lt 7) {
    # Windows PowerShell 5.1 drops empty native arguments (update -Phone ""), so the harness itself
    # must run under pwsh; -Shell picks which PowerShell runs the scripts under test.
    Write-Host "Run this test with pwsh (PowerShell 7). To test the scripts under Windows PowerShell 5.1: pwsh -File tests/Test-PeopleContacts.ps1 -Shell powershell" -ForegroundColor Red
    exit 1
}
$ScriptsDir = (Resolve-Path $ScriptsDir).Path
. (Join-Path $ScriptsDir "_common.ps1")
. (Join-Path $ScriptsDir "_people.ps1")
$null = Connect-Outlook

$outDir = Join-Path $PSScriptRoot "out"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$script:fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = "") {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { $script:fails++; Write-Host "FAIL  $name $detail" -ForegroundColor Red }
}
# Runs one script in a child process with a 60 s limit: a hang fails the check (exit 124)
# instead of blocking the run. ArgumentList passes each argument exactly (empty strings too).
function Run([string]$script, [object[]]$a, [string]$Exe = $Shell, [int]$TimeoutSec = 60) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $psi = [Diagnostics.ProcessStartInfo]::new($Exe)
    foreach ($x in @("-NoProfile", "-File", (Join-Path $ScriptsDir $script)) + @($a)) { $psi.ArgumentList.Add([string]$x) }
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.UseShellExecute = $false
    $p = [Diagnostics.Process]::Start($psi)
    $stdout = $p.StandardOutput.ReadToEndAsync(); $stderr = $p.StandardError.ReadToEndAsync()
    if (-not $p.WaitForExit($TimeoutSec * 1000)) {
        try { $p.Kill($true) } catch { }
        return [pscustomobject]@{ Exit = 124; Out = "TIMEOUT after $TimeoutSec s: $script"; Ms = $sw.ElapsedMilliseconds }
    }
    $p.WaitForExit()
    return [pscustomobject]@{ Exit = $p.ExitCode; Out = ($stdout.Result + $stderr.Result); Ms = $sw.ElapsedMilliseconds }
}
# How many of these EntryIDs are still in Contacts, read by a SEPARATE process with a fresh table
# (this session can keep seeing deleted items). Retries up to 5 x 500 ms until it reaches $Expect.
function Get-FreshContactsCount([string[]]$Ids, [int]$Expect = 0) {
    $probe = @'
$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
$t = $ns.GetDefaultFolder(10).GetTable(); $t.Columns.RemoveAll(); [void]$t.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x66700102")
$want = $env:SUITE_TEST_IDS.Split(","); $n = 0
while (-not $t.EndOfTable) { $v = $t.GetNextRow().Item(1); if ($v -is [byte[]] -and $want -contains [BitConverter]::ToString($v).Replace("-", "")) { $n++ } }
$n
'@
    $env:SUITE_TEST_IDS = ($Ids -join ",")
    $count = -1
    for ($try = 0; $try -lt 5; $try++) {
        $count = [int](& pwsh -NoProfile -Command $probe | Select-Object -Last 1)
        if ($count -eq $Expect) { break }
        Start-Sleep -Milliseconds 500
    }
    $env:SUITE_TEST_IDS = $null
    return $count
}

function JsonOf($r) { return ($r.Out.Trim() -split "`n")[-1] | ConvertFrom-Json }

$tag = "[suite-test] People $(Get-Random -Maximum 99999)"
$id = $null
Write-Host "Shell: $Shell" -ForegroundColor Gray

$oldConfirm = $env:OUTLOOK_SUITE_REQUIRE_CONFIRM
$env:OUTLOOK_SUITE_REQUIRE_CONFIRM = "1"   # force the gate on for the whole run, whatever the user's config says

try {
    # create
    $r = Run "outlook-contacts-create.ps1" @("-FullName", $tag, "-Email", "suite-test@suite-test.invalid", "-Phone", "555-0100", "-Company", "Suite Test Co", "-JobTitle", "Tester", "-Json")
    $j = JsonOf $r; $id = $j.EntryID
    Check "create -Json ($($r.Ms) ms)" ($r.Exit -eq 0 -and $id -and $j.Action -eq "create")
    $c = $Namespace.GetItemFromID($id)
    Check "create stored fields" ($c.FullName -eq $tag -and $c.Email1Address -eq "suite-test@suite-test.invalid" -and $c.CompanyName -eq "Suite Test Co")

    # list / search
    $r = Run "outlook-contacts-list.ps1" @("-Search", $tag, "-Json")
    $j = JsonOf $r
    Check "list -Search -Json ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Total -eq 1 -and $j.Contacts[0].EntryID -eq $id)
    $r = Run "outlook-contacts-list.ps1" @("-Limit", "5")
    Check "list text ($($r.Ms) ms)" ($r.Exit -eq 0 -and $r.Out -match "=== CONTACTS")
    $r = Run "outlook-contacts-list.ps1" @("-AllStores", "-Json", "-Limit", "1")
    $j = JsonOf $r
    Check "list -AllStores ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Total -ge 1)
    $r = Run "outlook-contacts-search.ps1" @("-Name", $tag.Substring(1), "-Company", "suite test", "-Json")
    $j = JsonOf $r
    Check "search -Name -Company AND ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Total -eq 1)
    $r = Run "outlook-contacts-search.ps1" @("-Name", $tag.Substring(1), "-Company", "nomatch-zz", "-Json")
    $j = JsonOf $r
    Check "search AND excludes" ($r.Exit -eq 0 -and $j.Total -eq 0)
    $r = Run "outlook-contacts-search.ps1" @()
    Check "search without criteria fails" ($r.Exit -eq 1)

    # update by EntryID and by Index
    $r = Run "outlook-contacts-update.ps1" @("-EntryID", $id, "-JobTitle", "Lead Tester", "-Mobile", "555-0101", "-Department", "QA", "-Json")
    $c = $Namespace.GetItemFromID($id)
    Check "update -EntryID ($($r.Ms) ms)" ($r.Exit -eq 0 -and $c.JobTitle -eq "Lead Tester" -and $c.MobileTelephoneNumber -eq "555-0101" -and $c.Department -eq "QA")
    $r = Run "outlook-contacts-update.ps1" @("-Index", "1", "-Search", $tag, "-Phone", "")
    $c = $Namespace.GetItemFromID($id)
    Check "update -Index -Search clears a field" ($r.Exit -eq 0 -and -not $c.BusinessTelephoneNumber)
    $r = Run "outlook-contacts-update.ps1" @("-EntryID", $id)
    Check "update without fields fails" ($r.Exit -eq 1)

    # export
    $r = Run "outlook-contacts-export.ps1" @("-EntryID", $id, "-Path", "$outDir\", "-Json")
    $j = JsonOf $r
    Check "export vCard ($($r.Ms) ms)" ($r.Exit -eq 0 -and (Test-Path -LiteralPath $j.Path) -and ((Get-Content -LiteralPath $j.Path -Raw) -match "BEGIN:VCARD"))
    if ($j.Path) { Remove-Item -LiteralPath $j.Path -ErrorAction SilentlyContinue }

    # photo (1x1 PNG)
    $png = Join-Path $outDir "suite-test-photo.png"
    [IO.File]::WriteAllBytes($png, [Convert]::FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="))
    $r = Run "outlook-contacts-photo.ps1" @("-EntryID", $id, "-PhotoPath", $png)
    $c = $Namespace.GetItemFromID($id)
    Check "photo add ($($r.Ms) ms)" ($r.Exit -eq 0 -and $c.HasPicture)
    $r = Run "outlook-contacts-photo.ps1" @("-EntryID", $id, "-Remove")
    $c = $Namespace.GetItemFromID($id)
    Check "photo remove" ($r.Exit -eq 0 -and -not $c.HasPicture)
    Remove-Item -LiteralPath $png -ErrorAction SilentlyContinue

    # people (read-only)
    $me = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    $r = Run "outlook-people.ps1" @($tag, "-Json")
    $j = JsonOf $r
    Check "people finds the test contact ($($r.Ms) ms)" ($r.Exit -eq 0 -and @($j.People | Where-Object { $_.Email -eq "suite-test@suite-test.invalid" -and $_.Sources -contains "Contacts" }).Count -eq 1)
    $r = Run "outlook-people.ps1" @($me, "-Json")
    $j = JsonOf $r
    Check "people by email: directory hit with contact points ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.People[0].Email -eq $me -and $j.People[0].Sources -contains "Directory" -and $j.People[0].Title)
    $r = Run "outlook-people.ps1" @($me, "-Profile", "-Json")
    $j = JsonOf $r
    Check "people -Profile ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Person.Email -eq $me -and $null -ne $j.Mail.Received -and $null -ne $j.Meetings.Past)
    $r = Run "outlook-people.ps1" @("zzqx-nobody-zzqx", "-Json")
    $j = JsonOf $r
    Check "people no match" ($r.Exit -eq 0 -and $j.Count -eq 0)
    $r = Run "outlook-people.ps1" @("zzqx-nobody-zzqx", "-Profile")
    Check "people -Profile no match fails" ($r.Exit -eq 1)

    # Clear-person rule (synthetic records: Match is the pre-bonus match, Score includes bonuses)
    function P($m, $s, $ns = 1) { [ordered]@{ Name = "p$m-$s"; Email = "p$m$s@suite-test.invalid"; Match = $m; Score = $s; Namesakes = $ns } }
    Check "clear: exact 100 over 80" ((Select-ClearPerson @((P 100 100), (P 80 95))).Reason -eq "clear")
    Check "clear: 80 leading 55 by 20+" ((Select-ClearPerson @((P 80 80), (P 55 60))).Reason -eq "clear")
    Check "not clear: substring 60 lifted by bonuses to 95" ((Select-ClearPerson @((P 60 95))).Reason -eq "weak")
    Check "not clear: 80 vs 80 (bonus order ignored)" ((Select-ClearPerson @((P 80 100), (P 80 80))).Reason -eq "ambiguous")
    Check "not clear: 90 vs 80 lead under 20" ((Select-ClearPerson @((P 90 90), (P 80 99))).Reason -eq "ambiguous")
    Check "not clear: exact name with directory namesakes" ((Select-ClearPerson @((P 100 100 2))).Reason -eq "namesakes")

    # Directory namesakes: a display name shared by several entries never resolves to one address
    $names = Get-DirectoryNames -Namespace $Namespace
    $dup = $names | Group-Object | Where-Object { $_.Count -gt 1 -and $_.Name -match '^\S+ \S+$' } | Select-Object -First 1
    if ($dup) {
        $all = Find-DirectoryNames -Names $names -Query $dup.Name -Limit 5   # assign first: comma-wrapped result
        $hit = @($all | Where-Object { $_.Name -eq $dup.Name })
        Check "namesakes counted in the name index ($($dup.Count) entries)" ($hit.Count -eq 1 -and $hit[0].Count -eq $dup.Count)
        $threw = $false; try { $null = Resolve-PersonSmtp -Namespace $Namespace -Query $dup.Name } catch { $threw = $_.Exception.Message -match "more than one person" }
        Check "Resolve-PersonSmtp throws for namesakes" $threw
    } else { Write-Host "SKIP  namesakes (no shared display name in the directory)" -ForegroundColor DarkGray }

    # One-off SMTP addresses are not directory hits
    $found = Resolve-Person -Namespace $Namespace -Query "nobody-zz-suite@suite-test.invalid" -Sources Directory
    Check "external address is not a Directory hit" ($found.Count -eq 0)
    $sw = [Diagnostics.Stopwatch]::StartNew(); $smtp = Resolve-PersonSmtp -Namespace $Namespace -Query $me
    Check "Resolve-PersonSmtp returns one string for an address ($($sw.ElapsedMilliseconds) ms)" ($smtp -is [string] -and $smtp -eq $me)

    # Received folders: Inbox and rule folders, never Sent / Deleted / Drafts / Junk
    $sw = [Diagnostics.Stopwatch]::StartNew(); $rf = Get-ReceivedFolders -Namespace $Namespace
    $ids = @($rf | ForEach-Object { $_.EntryID })
    $bad = @(foreach ($n in 3, 5, 16, 23) { $Namespace.GetDefaultFolder($n).EntryID }) | Where-Object { $ids -contains $_ }
    Check "received folders: $($rf.Count) incl. Inbox, no Sent/Deleted/Drafts/Junk ($($sw.ElapsedMilliseconds) ms)" ($ids -contains $Namespace.GetDefaultFolder(6).EntryID -and -not $bad)

    # -Members: a local contact group, then an Exchange distribution list
    $grpName = "$tag group"
    $g = $Outlook.CreateItem(7)   # olDistributionListItem
    $g.DLName = $grpName
    foreach ($a in "suite-test-1@suite-test.invalid", "suite-test-2@suite-test.invalid") { $rc = $Namespace.CreateRecipient($a); $null = $rc.Resolve(); $g.AddMember($rc) }
    $g.Save()
    $r = Run "outlook-people.ps1" @("-Members", $grpName, "-Json")
    $j = JsonOf $r
    Check "-Members contact group ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Source -eq "Contact group" -and $j.Total -eq 2 -and @($j.Members | Where-Object { $_.Email -like "suite-test-*@suite-test.invalid" }).Count -eq 2)
    $memberOf = $Namespace.CurrentUser.AddressEntry.GetExchangeUser().GetMemberOfList()
    $dlAddr = $null
    if ($memberOf) { for ($i = 1; $i -le [Math]::Min($memberOf.Count, 10); $i++) { $e = $memberOf.Item($i); if ($e.AddressEntryUserType -eq 1) { $dlAddr = $e.GetExchangeDistributionList().PrimarySmtpAddress; break } } }
    if ($dlAddr) {
        $r = Run "outlook-people.ps1" @("-Members", $dlAddr, "-Limit", "5", "-Json")
        $j = JsonOf $r
        Check "-Members distribution list, first 5 of $($j.Total) ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Source -eq "Directory" -and $j.Total -ge 1 -and $j.Count -le 5 -and $j.Count -ge 1)
    } else { Write-Host "SKIP  -Members distribution list (you are in none)" -ForegroundColor DarkGray }
    $r = Run "outlook-people.ps1" @("-Members", "zzqx-no-such-group", "-Json")
    Check "-Members unknown group fails" ($r.Exit -eq 1)

    # delete: preview, then confirm
    # Org view and directory search by attribute (read-only)
    $meUser = $Namespace.CurrentUser.AddressEntry.GetExchangeUser()
    $r = Run "outlook-people.ps1" @("-Org", $me, "-Json")
    $j = JsonOf $r
    Check "-Org: person, $(@($j.Managers).Count) managers, $($j.DirectReports.Total) reports, $($j.Peers.Total) peers ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Person.Email -eq $me -and $null -ne $j.Peers.Total -and @($j.Peers.People).Count -le 10)
    $r = Run "outlook-people.ps1" @("-Org", "zzqx-nobody-zzqx")
    Check "-Org unknown person fails" ($r.Exit -eq 1)
    if ($meUser.Department) {
        $last = ($meUser.Name -split ' ')[-1]
        $r = Run "outlook-people.ps1" @($last, "-Department", $meUser.Department, "-Json")
        $j = JsonOf $r
        Check "-Department finds yourself by your own department ($($r.Ms) ms)" ($r.Exit -eq 0 -and @($j.People | Where-Object { $_.Name -eq $meUser.Name }).Count -ge 1)
    } else { Write-Host "SKIP  -Department (your directory entry has no department)" -ForegroundColor DarkGray }
    $r = Run "outlook-people.ps1" @("-Title", "zzqx-no-such-title", "-Json")
    $j = JsonOf $r
    Check "-Title with no match ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Total -eq 0)
    $af = Join-Path $env:LOCALAPPDATA "outlook-suite\directory-attributes.tsv"
    $galCount = (Get-DirectoryList -Namespace $Namespace).AddressEntries.Count
    $afLines = 0; if (Test-Path $af) { $afLines = ([IO.File]::ReadAllLines($af)).Count }
    Check "attribute index complete ($afLines of $galCount lines, $([Math]::Round((Get-Item $af).Length / 1KB)) KB)" ($afLines -ge 0.95 * $galCount)

    # Dedupe. A and C share an email (mergeable; keeper = most recently modified), A and B share only
    # a name (report only), the two role contacts share an address but no name word (report only).
    # C also carries an unresolvable Exchange (EX) address that must never be copied.
    $mk = { param($n, $e, $co, $mob, $note) $x = $Outlook.CreateItem(2); $x.FullName = $n; if ($e) { $x.Email1Address = $e }; if ($co) { $x.CompanyName = $co }; if ($mob) { $x.MobileTelephoneNumber = $mob }; if ($note) { $x.Body = $note }; $x.Save(); Start-Sleep -Milliseconds 1100; $x.EntryID }
    $dA = & $mk "[suite-test] People Dup One" "dup1@suite-test.invalid" "Old Co" "" "note A"
    $dB = & $mk "[suite-test] people dup, one" "" "Name Co" "" ""
    $dC = & $mk "[suite-test] People Other" "DUP1@suite-test.invalid" "Dup Co" "555-0199" "note C"
    $x = $Namespace.GetItemFromID($dC); $x.Email2Address = "/O=SUITE-TEST/OU=NONE/CN=RECIPIENTS/CN=NOBODY"; try { $x.Email2AddressType = "EX" } catch { }; $x.Save()
    $roles = @(foreach ($n in "Alphazq Onezq", "Betazq Twozq") { $x = $Outlook.CreateItem(2); $x.FullName = $n; $x.Email1Address = "role@suite-test.invalid"; $x.CompanyName = "[suite-test] People Role"; $x.Categories = "[suite-test]"; $x.Save(); $x.EntryID })
    $inContactsCount = { param($ids) @(foreach ($x in $ids) { try { if ($Namespace.GetItemFromID($x).Parent.EntryID -eq $Namespace.GetDefaultFolder(10).EntryID) { 1 } } catch { } }).Count }
    $r = Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Json")
    $j = JsonOf $r
    $g = $j.DuplicateGroups | Select-Object -First 1
    Check "dedupe report: 1 mergeable group, keeper = most recently modified, plan hash ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Groups -eq 1 -and $j.Duplicates -eq 1 -and $g.Keep -eq $dC -and $j.Plan -match '^[0-9a-f]{16}$')
    Check "dedupe: role address with different names is report-only" (@($j.ReportOnlyGroups | Where-Object { $_.Reason -match 'no shared name word' }).Count -eq 1)
    Check "dedupe: same name only is report-only" ($j.NameOnlyGroups -eq 1)
    Check "dedupe: conflict reported (Company kept vs dropped)" (@($g.Detail.Conflicts | Where-Object { $_.Field -eq "Company" -and $_.Kept -eq "Dup Co" -and $_.Dropped -eq "Old Co" }).Count -eq 1 -and $g.Detail.NotesAppended -eq 1)
    $plan1 = $j.Plan
    $r = Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Merge", "-Json")
    $j = JsonOf $r
    Check "dedupe -Merge without -Confirm only previews" ($r.Exit -eq 0 -and $j.Result -eq "preview" -and (& $inContactsCount @($dA, $dB, $dC)) -eq 3)
    $r = Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Merge", "-Confirm", "-Json")
    Check "dedupe -Merge -Confirm without -Plan is refused" ($r.Exit -eq 1 -and (& $inContactsCount @($dA, $dC)) -eq 2)
    $dD = & $mk "[suite-test] People Dup Three" "dup1@suite-test.invalid" "" "" ""   # changes the plan
    $r = Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Merge", "-Confirm", "-Plan", $plan1, "-Json")
    Check "dedupe stale -Plan is refused, nothing deleted" ($r.Exit -eq 1 -and $r.Out -match "plan changed" -and (& $inContactsCount @($dA, $dC, $dD)) -eq 3)
    $j = JsonOf (Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Json"))
    $keepId = ($j.DuplicateGroups | Select-Object -First 1).Keep
    $r = Run "outlook-contacts-dedupe.ps1" @("-Search", "[suite-test] People", "-Merge", "-Confirm", "-Plan", $j.Plan, "-Json")
    $m = JsonOf $r
    $kept = $Namespace.GetItemFromID($keepId)
    $leftFresh = Get-FreshContactsCount @(@($dA, $dC, $dD) | Where-Object { $_ -ne $keepId }) 0
    $slots = @($kept.Email1Address, $kept.Email2Address, $kept.Email3Address)
    Check "dedupe -Merge -Confirm -Plan: 2 deleted, notes appended, no legacy DN copied ($($r.Ms) ms)" ($r.Exit -eq 0 -and $m.Deleted -eq 2 -and $m.DeletedTo -eq "Deleted Items" -and ($kept.Body -match "note A") -and ($kept.Body -match "note C") -and @($slots | Where-Object { $_ -match '^/O=' -and $_ -match 'SUITE-TEST' }).Count -le 1 -and $leftFresh -eq 0) "exit=$($r.Exit) deleted=$($m.Deleted) failed=$($m.Failed) noteA=$($kept.Body -match 'note A') noteC=$($kept.Body -match 'note C') dn=$(@($slots | Where-Object { $_ -match '^/O=' }).Count) left=$leftFresh keeper=$(@($dA,$dC,$dD).IndexOf($keepId))"
    Check "dedupe left the same-name and role contacts alone" ((& $inContactsCount @($dB)) -eq 1 -and (& $inContactsCount $roles) -eq 2)
    $dnInKeeper = @($slots | Where-Object { $_ -match '^/O=SUITE-TEST' }).Count
    $dnWasKeeper = ($keepId -eq $dC)
    Check "dedupe never writes an unresolved EX address into another contact" ($dnWasKeeper -or $dnInKeeper -eq 0)

    # Import: CSV with a header map, then -Update, then vCard
    $csv = Join-Path $outDir "suite-test-import.csv"
    @(
        "Name,Work Email,Cell,Org"
        "[suite-test] People Imp A,impa@suite-test.invalid,555-0001,Imp Co"
        "[suite-test] People Imp B,impb@suite-test.invalid,,"
        "[suite-test] People Existing,suite-test@suite-test.invalid,,"
        ",,555-0002,"
    ) | Set-Content -LiteralPath $csv -Encoding UTF8
    $map = "Work Email=Email; Cell=Mobile; Org=Company"
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-Json")
    $j = JsonOf $r
    $imp = (Get-ContactRecords -Namespace $Namespace -Search "[suite-test] People Imp").Count
    Check "import preview: 2 create, 1 skip, 1 invalid, nothing created ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Result -eq "preview" -and $j.Counts.create -eq 2 -and $j.Counts.'skip (email exists)' -eq 1 -and $j.Counts.invalid -eq 1 -and $imp -eq 0)
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-Confirm", "-Json")
    $j = JsonOf $r
    $all = Get-ContactRecords -Namespace $Namespace -Search "impa@suite-test.invalid"; $a = $all.ToArray()
    Check "import -Confirm creates 2 with mapped fields ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Created -eq 2 -and $a.Count -eq 1 -and $a[0].MobilePhone -eq "555-0001" -and $a[0].Company -eq "Imp Co")
    @(
        "Name,Work Email,Org,Title,Notes"
        "[suite-test] People Imp A,impa@suite-test.invalid,Imp Co 2,Lead,imp note"
        "[suite-test] People Imp A,,Imp Co,,"
        "[suite-test] People Role Row,role@suite-test.invalid,,,"
    ) | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-Update", "-Json")
    $j = JsonOf $r
    $diff = @(($j.Rows | Where-Object { $_.Action -eq "update" } | Select-Object -First 1).Changes)
    Check "import -Update preview: field diff, name+company skip, ambiguous email" ($r.Exit -eq 0 -and $j.Result -eq "preview" -and @($diff | Where-Object { $_.Field -eq "JobTitle" -and $_.Mode -eq "fill" }).Count -eq 1 -and @($diff | Where-Object { $_.Field -eq "Company" }).Count -eq 0 -and $j.Counts.'skip (name and company exist)' -eq 1 -and $j.Counts.ambiguous -eq 1)
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-Update", "-Confirm", "-Json")
    $j = JsonOf $r
    $all = Get-ContactRecords -Namespace $Namespace -Search "impa@suite-test.invalid"; $a = $all.ToArray()
    $it = $Namespace.GetItemFromID($a[0].EntryID)
    Check "import -Update fills empty fields only, appends notes" ($r.Exit -eq 0 -and $j.Updated -eq 1 -and $j.Created -eq 0 -and $a[0].Company -eq "Imp Co" -and $a[0].JobTitle -eq "Lead" -and $it.Body -match "imp note")
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-Update", "-Overwrite", "-Confirm", "-Json")
    $all = Get-ContactRecords -Namespace $Namespace -Search "impa@suite-test.invalid"; $a = $all.ToArray()
    Check "import -Update -Overwrite replaces differing values" ($r.Exit -eq 0 -and $a[0].Company -eq "Imp Co 2" -and (([regex]::Matches($Namespace.GetItemFromID($a[0].EntryID).Body, "imp note")).Count -eq 1))
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-Map", $map, "-MaxRows", "1")
    Check "import over -MaxRows is refused" ($r.Exit -eq 1 -and $r.Out -match "MaxRows")
    $r = Run "outlook-contacts-import.ps1" @("-Path", $csv, "-MaxFileMB", "0.00001")
    Check "import over -MaxFileMB is refused" ($r.Exit -eq 1 -and $r.Out -match "MB limit")

    # UTF-8 with BOM and non-ASCII names, imported under Windows PowerShell 5.1
    $bom = Join-Path $outDir "suite-test-bom.csv"
    $uName = "[suite-test] People Zo" + [char]0xEB + " " + [char]0x15E + "ahin " + [char]0xC7 + "elik"
    [IO.File]::WriteAllText($bom, "Name,Email`r`n$uName,zoe@suite-test.invalid`r`n", (New-Object Text.UTF8Encoding($true)))
    $r = Run "outlook-contacts-import.ps1" @("-Path", $bom, "-Confirm", "-Json") -Exe "powershell"
    $all = Get-ContactRecords -Namespace $Namespace -Search "zoe@suite-test.invalid"; $z = $all.ToArray()
    Check "import BOM CSV with non-ASCII name under powershell 5.1" ($r.Exit -eq 0 -and $z.Count -eq 1 -and $z[0].FullName -eq $uName)

    $vcf = Join-Path $outDir "suite-test-import.vcf"
    @(
        "BEGIN:VCARD", "VERSION:3.0", "FN:[suite-test] People Vcf", "N:Vcf;People;;;", "EMAIL;TYPE=INTERNET:vcf@suite-test.invalid",
        "TEL;VALUE=uri;TYPE=cell:tel:+1-555-0003", "ORG:Smith\; Sons;Research", "TITLE:Tester",
        "NOTE;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:caf=C3=A9 au lait", "END:VCARD",
        "BEGIN:VCARD", "VERSION:3.0", "FN:[suite-test] People Imp B again", "EMAIL:impb@suite-test.invalid", "END:VCARD",
        "BEGIN:VCARD", "VERSION:3.0", "FN:[suite-test] People Broken", "EMAIL:broken@suite-test.invalid"
    ) | Set-Content -LiteralPath $vcf -Encoding UTF8
    $r = Run "outlook-contacts-import.ps1" @("-Path", $vcf, "-Confirm", "-Json")
    $j = JsonOf $r
    $all = Get-ContactRecords -Namespace $Namespace -Search "vcf@suite-test.invalid"; $v = $all.ToArray()
    $vb = if ($v.Count) { [string]$Namespace.GetItemFromID($v[0].EntryID).Body } else { "" }
    Check "import vCard: tel: URI, ORG with \;, quoted-printable note, 1 skipped, 1 malformed ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Created -eq 1 -and $j.Counts.'skip (email exists)' -eq 1 -and $j.Counts.invalid -eq 1 -and $v.Count -eq 1 -and ($v[0].MobilePhone -replace '\D', '') -eq "15550003" -and $v[0].Company -eq "Smith; Sons" -and $v[0].Department -eq "Research" -and $vb -eq ("caf" + [char]0xE9 + " au lait") -and $v[0].FullName -eq "[suite-test] People Vcf")
    foreach ($tmpFile in $csv, $vcf, $bom) { [IO.File]::Delete($tmpFile) }

    # Confirm wrapper: exactly one [bool] $false when confirmation is missing, in both modes
    $out = @(Approve-ContactAction -Action "delete" -Confirmed $false -What "x" -Json 6>$null)
    Check "Approve-ContactAction -Json returns exactly one [bool] false" ($out.Count -eq 1 -and $out[0] -is [bool] -and -not $out[0])
    $out = @(Approve-ContactAction -Action "delete" -Confirmed $false -What "x" 6>$null)
    Check "Approve-ContactAction text returns exactly one [bool] false" ($out.Count -eq 1 -and $out[0] -is [bool] -and -not $out[0])
    $r = Run "outlook-contacts-delete.ps1" @("-EntryID", $id, "-Json")
    $j = JsonOf $r
    $still = $null; try { $still = $Namespace.GetItemFromID($id) } catch { }
    Check "delete -Json previews and does not delete" ($r.Exit -eq 0 -and $j.Result -eq "preview" -and $still -and $still.Parent.EntryID -eq $Namespace.GetDefaultFolder(10).EntryID)
    $r = Run "outlook-contacts-delete.ps1" @("-EntryID", $id)
    $still = $null; try { $still = $Namespace.GetItemFromID($id) } catch { }
    Check "delete previews without -Confirm" ($r.Exit -eq 0 -and $r.Out -match "PREVIEW" -and $still -and $still.Parent.EntryID -eq $Namespace.GetDefaultFolder(10).EntryID)
    $r = Run "outlook-contacts-delete.ps1" @("-Name", $tag, "-Confirm", "-Json")
    $j = JsonOf $r
    $inContacts = (Get-ContactRecords -Namespace $Namespace -Search $tag).Count
    Check "delete -Name -Confirm ($($r.Ms) ms)" ($r.Exit -eq 0 -and $j.Action -eq "delete" -and $inContacts -eq 0)
    $t = $Namespace.GetDefaultFolder(3).GetTable("[MessageClass] = 'IPM.Contact'")
    $t.Columns.RemoveAll(); [void]$t.Columns.Add("Subject")
    $inDeleted = 0; while (-not $t.EndOfTable) { if ([string]$t.GetNextRow().Item(1) -eq $tag) { $inDeleted++ } }
    Check "deleted contact is in Deleted Items, found by Subject for the purge" ($inDeleted -eq 1)
}
finally {
    $env:OUTLOOK_SUITE_REQUIRE_CONFIRM = $oldConfirm
    # Remove anything this run left behind: Contacts first (Delete moves to Deleted Items),
    # then Deleted Items (Delete purges). Only "[suite-test]" items, which only tests create.
    $del = $Namespace.GetDefaultFolder(3)
    foreach ($folder in $Namespace.GetDefaultFolder(10), $del) {
        # contacts and contact groups this run created; in Deleted Items, Delete() purges them
        $t = $folder.GetTable("[MessageClass] = 'IPM.Contact' OR [MessageClass] = 'IPM.DistList'")
        $t.Columns.RemoveAll(); [void]$t.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x66700102"); [void]$t.Columns.Add("Subject"); [void]$t.Columns.Add("Categories")
        # Collect first, then delete: deleting while reading the same table skips rows.
        $ids = New-Object System.Collections.Generic.List[string]
        while (-not $t.EndOfTable) {
            $row = $t.GetNextRow()
            # Anywhere in the subject: Outlook reorders "Last, First" names ("x, one" -> "one x").
            if ([string]$row.Item(2) -like "*[[]suite-test] *" -or [string]$row.Item(3) -like "*[[]suite-test]*") {   # role contacts carry the tag as a category
                $eid = $row.Item(1); if ($eid -is [byte[]]) { $eid = ConvertTo-HexId $eid }
                $ids.Add([string]$eid)
            }
        }
        foreach ($eid in $ids) { try { $Namespace.GetItemFromID($eid).Delete() } catch { } }
    }
}

Write-Host ""
if ($script:fails -eq 0) { Write-Host "All checks passed." -ForegroundColor Green; exit 0 }
Write-Host "$script:fails check(s) failed." -ForegroundColor Red; exit 1
