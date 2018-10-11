<#

.SYNOPSIS

Script will consume an Excel template file to mass deploy sensor types

#>

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

Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

$importList = Get-ChildItem $RunDir -Filter "Sensor_Type_Template.xlsx"

Log "Extracting Template Data..."
ForEach ($xl in $importList) {
    Write-Host -ForegroundColor Magenta "Importing Data from Excel : $($xl.name)..."
    ExportWSToCSV -excelFile $xl -csvLoc $RunDir\Temp
}

$Template = Import-Csv $RunDir\Temp\Template.csv
$Measurement_Types = Import-Csv $RunDir\Temp\Measurement_Types.csv


Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10){
    Log "Retrieved Login Ticket"
}else {
    Write-Error "Error Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

$Count = 0
ForEach($Type in $Template){
    $Count++
    IF ($Type.Sensor_Type_CD -eq $Null -or $Type.Sensor_Type_CD -eq "NULL"){
        ""
        "";Log "$Count : Inserting Sensor Type: $($Type.Sensor_Name)"
    }ELSE{
        "";Log "$Count : Updating Sensor Type: $($Type.Sensor_Name)"
    }

    $SensorSetParams = @{
        data = $Type
        config = $Config
        user = $User 
    }
    Try{
    $SensorSet = $Null
    $SensorSet = SetSensorType332 @SensorSetParams
    }Catch{$WebException = $_}

    if($SensorSet.ResultCode -eq 0 -and 
        $SensorSet.ResultMessage -eq $Null -and 
        $Type.Sensor_Type_CD -eq 0){
            Log "Success! Inserted: $($SensorSet.Result.Sensor_Type_CD)"
    }
    elseif($SensorSet.ResultCode -eq 0 -and
        $SensorSet.ResultMessage -eq $Null -and
        $Type.Sensor_Type_CD -ne 0){
            Log "Success! Updated: $($SensorSet.Result.Sensor_Type_CD)"
    }
    else{
        Write-Error "Error : ResultCode: $($SensorSet.ResultCode) `n ResultMessage: $($SensorSet.ResultMessage) `n InvocationInfo: $($WebException.InvocationInfo.Line) `n $($WebException.Exception)"
    }

    IF($SensorSet.ResultCode -eq 0 -and $Type.Sensor_Type_CD -eq 0){
        $Type.Sensor_Type_CD = $SensorSet.Result.Sensor_Type_CD
    }
}
""
$Template | Export-Csv $RunDir\Temp\Template.csv -NoTypeInformation -Force

Log "Merging Updated Data to Template..."
#Get-ChildItem $RunDir -Filter "*.xlsx" | Remove-Item -Force -ErrorAction SilentlyContinue
Merge-CSVFiles -CSVPath $RunDir\Temp -XLOutput $RunDir\Sensor_Type_Template.xlsx

Log "Done, Template Updated"





