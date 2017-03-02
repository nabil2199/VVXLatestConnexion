<#
VVX IP address
Verion 0.1
OCWS
#>

param([string]$userCsv = "C:\Sources\users.csv",[string]$IPfilePath = "C:\Sources\ProvisioningPIN.csv")

#IP file initialise
Add-Content "upn,FirstName,LastName,EmailAddress,Desk,Extension,PIN" -Path $IPfilePath

# SQL Server for historical connection information
$SqlServer = "dsi-w4173bddv0.groupe.generali.fr\SQLT110_LYNK2_PA"

#User CSV loading
$usersList = $null
$usersList = Import-Csv $userCsv
$count = $usersList.count
Write-Host "User count within CSV file=" $count
Write-Host ""

foreach ($user in $usersList)
{
    $UserSIP = $user.SipAddress
    Write-Host " "
    Write-Host "Retrieving Recent Lync Client Connections for " -foregroundcolor green -NoNewline; Write-Host $UserSIP -foregroundcolor blue
    
