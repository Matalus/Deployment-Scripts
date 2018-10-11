$ErrorActionPreference = "stop"

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition
Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

Set-Location $RunDir

#Load function library from modules folder
Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "..\Modules\Functions" -DisableNameChecking

#Load config from JSON in project root dir
Log "Loading Configuration..."
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -config $Config

#Load Modules in PSModules section of config.json
Log "Loading Modules..."
Prereqs -config $Config

Log "Deleting Old Templates..."
Get-ChildItem $RunDir -Filter "*.xlsx" | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue


Log "Creating New Template..."
$Template = [PSCustomObject]@{
    Device_Type_CD           = "NULL"
    Sensor_Type_CD           = "NULL"
    Sensor_Name              = "<Sensor Name Here>"
    Sensor_Short_name        = "<Sensor Short Name Here>"
    Measurement_Type_CD      = 0
    Last_Updated_By          = 616
    Is_Deleted               = "FALSE"
    Is_Critical              = "FALSE"
    Rounding_Precision       = 0
    Deadband_Percentage      = "NULL"
    Disable_History          = "FALSE"
    External_Tag_Format      = "NULL"
    Setpoint_Timeout_Seconds = "NULL"
    Description              = "NULL"
    Sensor_Category_Type_CD  = "NULL"
}

$Template | Export-Csv $RunDir\Temp\Template.csv -NoTypeInformation

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
$Device_Types = (GetSelectMethod @params).Result | Sort-Object Device_Type_CD -Descending
$Device_Types | Export-Csv $RunDir\Temp\Device_Types.csv -NoTypeInformation

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


$Sensor_Types | Export-Csv $RunDir\Temp\Sensor_Types.csv -NoTypeInformation

Log "Getting Measurement Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "SELECT_897"
}
$Measurement_Types = GetSelectMethod @params
$Measurement_Types.Result | Export-Csv $RunDir\Temp\Measurement_Types.csv -NoTypeInformation

Log "Getting Sensor Category Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "Sensor_Category_Types_Get"
}
$Category_Types = GetSelectMethod @params
$Category_Types.Result | Export-Csv $RunDir\Temp\Senor_Category_Types.csv -NoTypeInformation

Log "Merging Data into Template..."
Merge-CSVFiles -CSVPath $RunDir\Temp -XLOutput $RunDir\Sensor_Type_Template.xlsx

Log "Cleaning up temp files..."
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

Log "Done."

Invoke-Item $RunDir\Sensor_Type_Template.xlsx

