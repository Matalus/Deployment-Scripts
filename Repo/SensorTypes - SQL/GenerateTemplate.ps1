$ErrorActionPreference = "stop"

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition

Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "$RunDir\Functions.psm1" -DisableNameChecking -ErrorAction SilentlyContinue
Import-Module "$RunDir\Modules\Invoke-SqlCmd2" -DisableNameChecking -ErrorAction SilentlyContinue
Prereqs
Log "Functions Loaded..."


Log "Loading Configuration..."
$Config = (Get-Content "$RunDir\config.json") -join "`n" | ConvertFrom-Json

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

Log "Getting Device Types Data..."
$Device_Types = SQLQuery -query "SELECT * FROM dbo.Device_Types WHERE Is_Deleted = 0"
$Device_Types | Export-Csv $RunDir\Temp\Device_Types.csv -NoTypeInformation

Log "Getting Sensor Types Data..."
$Sensor_Types = SQLQuery -query @"
SELECT
Device_Type_CD
,Sensor_Type_CD
,Sensor_Name
,Sensor_Short_name
,Measurement_Type_CD
,Last_Updated_By
,Is_Deleted
,Is_Critical
,Rounding_Precision
,ISNULL(CONVERT(varchar(10),Deadband_Percentage),'NULL') as [Deadband_Percentage] 
,Disable_History
,ISNULL(CONVERT(varchar(10),External_Tag_Format), 'NULL') AS [External_Tag_Name]
,ISNULL(CONVERT(varchar(10),Setpoint_Timeout_Seconds), 'NULL') AS [Setpoint_Timeout_Seconds] 
,ISNULL(CONVERT(varchar(10),Description), 'NULL') AS [Description]
,ISNULL(CONVERT(varchar(10),Sensor_Category_Type_CD), 'NULL') AS [Sensor_Category_Type_CD]
FROM dbo.Device_Sensor_Types
ORDER BY Sensor_Type_CD DESC
"@
$Sensor_Types | Export-Csv $RunDir\Temp\Sensor_Types.csv -NoTypeInformation

Log "Getting Measurement Types Data..."
$Measurement_Types = SQLQuery -query "SELECT * FROM dbo.Device_Measurement_Types WHERE Is_Deleted = 0"
$Measurement_Types | Export-Csv $RunDir\Temp\Measurement_Types.csv -NoTypeInformation

Log "Getting Sensor Category Types Data..."
$Category_Types = SQLQuery -query "SELECT * FROM dbo.Device_Sensor_Category_Types WHERE Is_Deleted = 0"
$Category_Types | Export-Csv $RunDir\Temp\Senor_Category_Types.csv

Log "Merging Data into Template..."
Merge-CSVFiles -CSVPath $RunDir\Temp -XLOutput $RunDir\Sensor_Type_Template.xlsx

Log "Cleaning up temp files..."
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

Log "Done."

Invoke-Item $RunDir\Sensor_Type_Template.xlsx

