param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Low", "Normal", "High")]
    [string]$Importance,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Set Importance: Low, Normal or High.
# Usage: .\outlook-set-importance.ps1 -EntryID "00000000..." -Importance High

. (Join-Path $PSScriptRoot "_common.ps1")

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $before = $script:ImportanceNames[[int]$email.Importance]
    $email.Importance = @{ "Low" = 0; "Normal" = 1; "High" = 2 }[$Importance]
    $email.Save()
    Write-ActionResult -Title "IMPORTANCE SET" -Item $email -Changes ([ordered]@{ Before = $before; Importance = $Importance }) -Json:$Json -Action "set-importance"
}
catch {
    Stop-WithError $_
}
