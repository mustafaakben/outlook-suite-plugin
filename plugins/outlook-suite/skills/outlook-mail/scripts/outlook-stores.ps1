param(
    [string]$Mailbox = "",
    [int]$TopFolders = 0,
    [switch]$Json
)

# Outlook Stores: the mailboxes and data files in the profile with type, cache mode, data file,
# size and quota (when the server reports one), and whether Instant Search / conversations work.
# -TopFolders N  the N biggest mail folders of the default mailbox (or -Mailbox), for clean-up.
# Usage: .\outlook-stores.ps1
# Usage: .\outlook-stores.ps1 -TopFolders 15

. (Join-Path $PSScriptRoot "_common.ps1")

$P = @{
    Size = "http://schemas.microsoft.com/mapi/proptag/0x0E080014"
    SendQuota = "http://schemas.microsoft.com/mapi/proptag/0x666E0003"
    ReceiveQuota = "http://schemas.microsoft.com/mapi/proptag/0x666A0003"
    StorageLimit = "http://schemas.microsoft.com/mapi/proptag/0x3FF50003"
}
$types = @{ 0 = "Exchange mailbox"; 1 = "Exchange delegate"; 2 = "Exchange public folders"; 3 = "Not Exchange (PST/IMAP/other)"; 4 = "Exchange additional mailbox" }
function Get-Prop($acc, [string]$tag) { try { return $acc.PropertyAccessor.GetProperty($tag) } catch { return $null } }
function Format-Size($bytes) { if ($null -eq $bytes) { return "" }; $b = [double]$bytes; if ($b -ge 1GB) { "{0:N1} GB" -f ($b / 1GB) } elseif ($b -ge 1MB) { "{0:N1} MB" -f ($b / 1MB) } else { "{0:N0} KB" -f ($b / 1KB) } }

try {
    $null = Connect-Outlook
    $default = $Namespace.DefaultStore.StoreID
    $list = @(foreach ($s in $Namespace.Stores) {
        $size = Get-Prop $s $P.Size
        $file = ""; try { $file = $s.FilePath } catch { }
        $o = [ordered]@{
            Name = $s.DisplayName; Type = $types[[int]$s.ExchangeStoreType]; Default = ($s.StoreID -eq $default)
            Cached = $(try { [bool]$s.IsCachedExchange } catch { $false }); DataFile = $file
            DataFileSize = $(if ($file -and (Test-Path -LiteralPath $file)) { Format-Size (Get-Item -LiteralPath $file).Length } else { "" })
            Size = Format-Size $size; SizeBytes = $size
            InstantSearch = $(try { [bool]$s.IsInstantSearchEnabled } catch { $false })
            Conversations = $(try { [bool]$s.IsConversationEnabled } catch { $false })
        }
        $sq = Get-Prop $s $P.SendQuota; $rq = Get-Prop $s $P.ReceiveQuota; $lim = Get-Prop $s $P.StorageLimit
        if ($sq) { $o.SendQuota = Format-Size ([double]$sq * 1KB) }
        if ($rq) { $o.ReceiveQuota = Format-Size ([double]$rq * 1KB) }
        if ($lim) { $o.StorageLimit = Format-Size ([double]$lim * 1KB); if ($size) { $o.PercentUsed = [math]::Round(100 * [double]$size / ([double]$lim * 1KB), 1) } }
        $o
    })

    $top = @()
    if ($TopFolders -gt 0) {
        $store = Get-RuleStore -Namespace $Namespace -Mailbox $Mailbox
        $top = @(foreach ($f in (Get-MailSubfolders $store.GetRootFolder())) {
            $sz = Get-Prop $f $P.Size
            [ordered]@{ Folder = (($f.FolderPath -replace '^\\\\[^\\]+\\', '') -replace '\\', '/'); Items = $(try { $f.Items.Count } catch { 0 }); SizeBytes = $sz; Size = Format-Size $sz }
        }) | Sort-Object { [double]$_.SizeBytes } -Descending | Select-Object -First $TopFolders
    }

    if ($Json) { Write-Json ([ordered]@{ stores = $list; topFolders = $top }); exit 0 }
    Write-Host "=== MAILBOXES AND DATA FILES ===" -ForegroundColor Cyan
    foreach ($o in $list) {
        Write-Host "  $($o.Name)$(if ($o.Default) { '  (default)' })" -ForegroundColor Yellow
        $bits = @($o.Type)
        if ($o.Size) { $bits += "size $($o.Size)" }
        if ($o.StorageLimit) { $bits += "limit $($o.StorageLimit) ($($o.PercentUsed)% used)" }
        if ($o.Cached) { $bits += "cached" }
        if ($o.DataFile) { $bits += "file $($o.DataFileSize)" }
        if ($o.InstantSearch) { $bits += "Instant Search" }
        Write-Host "    $($bits -join ' | ')"
    }
    if ($top.Count) {
        Write-Host "Biggest folders:" -ForegroundColor Cyan
        foreach ($t in $top) { Write-Host ("    {0,10}  {1,7} items  {2}" -f $t.Size, $t.Items, $t.Folder) }
    }
}
catch {
    Stop-WithError $_
}
