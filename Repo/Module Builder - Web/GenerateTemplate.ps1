$ErrorActionPreference = "stop"

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition
Set-Location $RunDir

Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

#Load function library from modules folder
Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "..\Modules\Functions" -DisableNameChecking

#Load config from JSON in project root dir
Log "Loading Configuration..." #test comment
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -config $Config

#Load Modules in PSModules section of config.json
Log "Loading Modules..."
Prereqs -config $Config

$TempDir = "$RunDir\Temp"
if((Test-Path $TempDir) -eq $false){New-Item -Path $TempDir -ItemType Directory}

Log "Deleting Old Templates..."
Get-ChildItem $RunDir -Filter "*.xlsx" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

Log "Creating New Template..."
$Template = [PSCustomObject]@{
    Update                   = "<? default 1, 0 ignore>"
    Partition_ID             = "<? Top Level Partition>"
    Parent_Path              = "<? Tech Space\PDU01>"
    Partition_Types           = "<? Module, Above Floor\PDU>"
    Device_Type              = "<? Power Distribution Unit>"
    Device_Type_Short        = "<? PDU>"
    Device_Name              = "<? PDU01>"
    Sensor_Type              = "<? Total KW>"
    Sensor_Type_Short        = "<? TotKW>"
    L                        = "=LEN(G2)"
    Reading_Type             = "<? Number>"
    Measurement_Type         = "<? KW>"
    Monitor_Type             = "<? Evaluation>"
    Monitor_URL              = "<? opc://10.0.7.35/>"
    Monitor_Server_ID        = "<? Monitor-Item?ServerId=ABB.AC800MC_OpcDaServer.3&ItemId="
    Monitor_Item_ID          = "<? Applications.D24.PDUA_Sigs.BCMS1.Amps_B.Value>"
    SP                       = "<? TRUE if SP>"
    SPMin                    = "<? Set Point Min>"
    SPMax                    = "<? Set Point Max>"
    BQ                       = "<? QualityID>"
    Last_Updated_By          = 616
    Polling                  = "<? 300>"
    Is_Deleted               = "FALSE"
    #Is_Critical              = "FALSE"
    Rounding                 = 0
    SP_Timeout               = "NULL"
    Description              = "NULL"
    CatType                  = "NULL"
    BQ_Pri                   = 0
    CF_Pri                   = 0
    Convert                  = "NULL"
    Thresh1                  = $Null
    Thresh2                  = $Null
    Thresh3                  = $Null
    Thresh4                  = $Null
    Thresh5                  = $Null
    TheshIgnore              = $Null
    Device_Type_CD           = $Null
    Device_ID                = $Null
    Sensor_Type_CD           = $Null
    Device_Sensor_ID         = $Null
    Path                     = $Null
}

#$Template | Export-Csv $RunDir\Temp\Template.csv -NoTypeInformation
$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Template"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Template"
    ClearSheet = $true
}

$Template | Export-Excel @ExportParams

Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10){
    Log "Retrieved Login Ticket"
}else {
    Write-Error "Error Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

Log "Getting Device Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "SELECT_1127"
}
$Device_Types = GetSelectMethod @params
#$Device_Types.Result | Export-Csv $RunDir\Temp\Device_Types.csv -NoTypeInformation

$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Device_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Device_Types"
    ClearSheet = $true
}

$Device_Types.Result | Export-Excel @ExportParams


Log "Getting Sensor Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.DeviceMethod
    method  = "Device_Sensor_Types_Get"
}
$SelectProperties = "Device_Type_CD",
"Sensor_Type_CD",
"Sensor_Name",
"Sensor_Short_name",
"Measurement_Type_CD",
"Last_Updated_By",
"Is_Deleted",
"Is_Critical",
"Rounding_Precision",
@{N = "Deadband_Percentage"; E = {if ([string]::IsNullOrEmpty($_.Deadband_Percentage)) {"NULL"}else {$_.Deadband_Percentage}}},
"Disable_History",
@{N = "External_Tag_Format"; E = {if ([string]::IsNullOrEmpty($_.External_Tag_Format)) {"NULL"}else {$_.External_Tag_Format}}},
@{N = "Setpoint_Timeout_Seconds"; E = {if ([string]::IsNullOrEmpty($_.Setpoint_Timeout_Seconds)) {"NULL"}else {$_.Setpoint_Timeout_Seconds}}},
@{N = "Description"; E = {if ([string]::IsNullOrEmpty($_.Description)) {"NULL"}else {$_.Description}}},
@{N = "Sensor_Category_Type_CD"; E = {if ([string]::IsNullOrEmpty($_.Sensor_Category_Type_CD)) {"NULL"}else {$_.Sensor_Category_Type_CD}}};

$Sensor_Types = (GetSelectMethod @params).Result |
    Select-Object -Property $SelectProperties |Sort-Object Sensor_Type_CD -Descending


#$Sensor_Types | Export-Csv $RunDir\Temp\Sensor_Types.csv -NoTypeInformation
$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Sensor_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Sensor_Types"
    ClearSheet = $true
}

$Sensor_Types | Export-Excel @ExportParams

Log "Getting Measurement Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "SELECT_897"
}
$Measurement_Types = GetSelectMethod @params
#$Measurement_Types.Result | Export-Csv $RunDir\Temp\Measurement_Types.csv -NoTypeInformation

$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Measurement_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Measurement_Types"
    ClearSheet = $true
}

$Measurement_Types.Result | Export-Excel @ExportParams


Log "Getting Sensor Category Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "Sensor_Category_Types_Get"
}
$Category_Types = GetSelectMethod @params
#$Category_Types.Result | Export-Csv $RunDir\Temp\Category_Types.csv -NoTypeInformation

$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Category_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Category_Types"
    ClearSheet = $true
}

$Category_Types.Result | Export-Excel @ExportParams

Log "Getting Monitor Types Data..."
$params = @{
    config = $Config
    user = $User
    service = $Config.DeviceMethod
    method = "SELECT_1049"
}

$Monitor_Types = GetSelectMethod @params
#$Monitor_Types.Result | Export-Csv $RunDir\Temp\Monitor_Types.csv -NoTypeInformation

$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Monitor_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Monitor_Types"
    ClearSheet = $true
}

$Monitor_Types.Result | Export-Excel @ExportParams


Log "Getting Reading Types Data..."
$params = @{
    config = $Config
    user = $User
    service = $Config.DeviceMethod
    method = "SELECT_1048"
}

$Reading_Types = GetSelectMethod @params
#$Reading_Types.Result | Export-Csv $RunDir\Temp\Reading_Types.csv -NoTypeInformation

$ExportParams = @{
    Path = "$RunDir\Templates\Integration_Template.xlsx"
    WorkSheetName = "Reading_Types"
    AutoSize = $true
    FreezeTopRow = $true
    BoldTopRow = $true
    TableStyle = "Medium21"
    TableName = "Reading_Types"
    ClearSheet = $true
}

$Reading_Types.Result | Export-Excel @ExportParams


#Kill Orphaned Excel Procs
[array]$Procs = Get-WmiObject Win32_Process -Filter "name = 'Excel.exe'"
$embeddedXL = $Procs | ?{$_.CommandLine -like "*/automation -Embedding*"}
$embeddedXL | ForEach-Object { $_.Terminate() | Out-Null }

Log "Merging Data into Template..."
#Merge-CSVFiles -CSVPath $RunDir\Temp -XLOutput $RunDir\Templates\Integration_Template.xlsx

Log "Cleaning up temp files..."
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

Log "Done."

Invoke-Item $RunDir\Templates\Integration_Template.xlsx
