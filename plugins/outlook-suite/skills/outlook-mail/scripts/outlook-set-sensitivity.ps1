param(
    [string]$EntryID = "",
    [string]$StoreID = "",
    [int]$Index = 0,
    [Parameter(Mandatory = $true)]
    [ValidateSet("Normal", "Personal", "Private", "Confidential")]
    [string]$Sensitivity,
    [int]$Days = 7,
    [switch]$Json
)

# Outlook Set Sensitivity: Normal, Personal, Private or Confidential.
# Usage: .\outlook-set-sensitivity.ps1 -EntryID "00000000..." -Sensitivity Confidential

. (Join-Path $PSScriptRoot "_common.ps1")

$map = @{ "Normal" = 0; "Personal" = 1; "Private" = 2; "Confidential" = 3 }
$names = @{ 0 = "Normal"; 1 = "Personal"; 2 = "Private"; 3 = "Confidential" }

try {
    $null = Connect-Outlook
    $email = Get-OutlookItem -Namespace $Namespace -EntryID $EntryID -StoreID $StoreID -Index $Index -Days $Days
    $before = $names[[int]$email.Sensitivity]
    $email.Sensitivity = $map[$Sensitivity]
    $email.Save()
    Write-ActionResult -Title "SENSITIVITY SET" -Item $email -Changes ([ordered]@{ Before = $before; Sensitivity = $Sensitivity }) -Json:$Json -Action "set-sensitivity"
}
catch {
    Stop-WithError $_
}
