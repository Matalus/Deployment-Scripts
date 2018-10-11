<#

.SYNOPSIS

Script will consume an Excel template file to mass deploy sensor types

#>

$ErrorActionPreference = "stop"

Log "Determining Root Path..."
$RunDir = split-path -parent $MyInvocation.MyCommand.Definition

Log "Loading Functions..."
Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "$RunDir\Functions.psm1" -DisableNameChecking -ErrorAction SilentlyContinue
Import-Module "$RunDir\Modules\Invoke-SqlCmd2"

Log "Loading Configuration..."
$Config = (Get-Content "$RunDir\config.json") -join "`n" | ConvertFrom-Json
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

$importList = Get-ChildItem $RunDir -Filter "Sensor_Type_Template.xlsx"

Log "Extracting Template Data..."
ForEach ($xl in $importList) {
    Write-Host -ForegroundColor Magenta "Importing Data from Excel : $($xl.name)..."
    ExportWSToCSV -excelFile $xl -csvLoc $RunDir\Temp
}

$Template = Import-Csv $RunDir\Temp\Template.csv

Log "Starting Sensor Type Insert / Update..."

$Count = 0
ForEach($Type in $Template){
    $Count++
    IF ($Type.Sensor_Type_CD -eq $Null -or $Type.Sensor_Type_CD -eq "NULL"){
        ""
        Log "$Count : Inserting Sensor Type: $($Type.Sensor_Name)"
    }ELSE{
        Log "$Count : Updating Sensor Type: $($Type.Sensor_Name)"
    }
        $ExecSQL = SQLQuery -query $(SensorTypeSet $Type)
    IF($ExecSQL.Sensor_Type_CD -ne $Null){
        Log "Success!"
    }

    IF($ExecSQL -and $Type.Sensor_Type_CD -eq "NULL"){
        $Type.Sensor_Type_CD = $ExecSQL.Sensor_Type_CD
    }
}

$Template | Export-Csv $RunDir\Temp\Template.csv -NoTypeInformation -Force

Log "Merging Updated Data to Template..."
#Get-ChildItem $RunDir -Filter "*.xlsx" | Remove-Item -Force -ErrorAction SilentlyContinue
Merge-CSVFiles -CSVPath $RunDir\Temp -XLOutput $RunDir\Sensor_Type_Template.xlsx

Log "Done, Template Updated"





