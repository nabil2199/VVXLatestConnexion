<#
VVX IP address
Verion 0.1
OCWS
#>

param([string]$userCsv = "C:\Sources\users.csv",[string]$IPfilePath = "C:\Sources\IPfile.csv")

#IP file initialise
Add-Content "upn,SipAddress,IP_address" -Path $IPfilePath

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

    $SQLQuery = "SELECT TOP 1 [SessionIdTime]
      ,[RegisterTime]
      ,[UserUri]
      ,[ClientVersion]
      ,[IpAddress]
      ,[DiagnosticId]
      ,[Registrar]
      ,[Pool]
      ,[IsUserServiceAvailable]
      ,[DeviceMacAddress]
      ,[DeviceManufacturer]
      ,[DeviceHardwareVersion]
      ,[DeRegisterTime]
      FROM [LcsCDR].[dbo].[RegistrationView] where [UserUri] LIKE '%$UserSIP%' AND [ClientVersion] LIKE '%Polycom%' ORDER BY SessionIdTime DESC"
      
    $Connection = new-object system.data.sqlclient.sqlconnection
    $Connection.connectionString="Data Source=$SqlServer;Initial Catalog=LcsCDR;Integrated Security=SSPI"
    $Connection.open()
    $Command = $Connection.CreateCommand()
    $Command.Commandtext = $SqlQuery
    $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $Command
    $Dataset = New-Object System.Data.Dataset
    $DataAdapter.Fill($Dataset)
    # $Dataset.Tables[0] | Export-CSV TempList.csv -notype
    $Connection.close()
    $Connection = $null

    $Results = $Dataset.Tables[0].rows
    $line = $user.upn + "," + $UserSIP + "," + $Results.IpAddress
    Add-Content $line -Path $IPfilePath

}