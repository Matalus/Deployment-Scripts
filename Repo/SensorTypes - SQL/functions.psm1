#Created by Matt Hamende BASELAYER TECHNOLOGY 2017
#Library of functions to drive Deployment Tools

Function Logo() {
    Write-Host @"
  *,,**    __       __  ____              ___  __  
,/(%###(. |__)  /\ |__ |__  |     /\ \ / |__  |__) 
/(###(%#/ |__) /  \___||___ |___ /  \ |  |____|  \ 
*/(###(*.
                                                                                                         
$((Get-Date).year) Baselayer Technology. LLC All Rights Reserved

$(Split-Path -Leaf $MyInvocation.ScriptName)

"@
}

Export-ModuleMember Logo

Function Prereqs (){
    $Modules = @(
        "PowerShellGet",
        "Invoke-Sqlcmd2"
    )

    ForEach($Module in $Modules){
        $installed = Get-Module -Name $Module
        if($installed){
            #Do nothing
        }
        else{
            Install-Module $Module -Force -Repository PSGallery
        }
    }
}

Export-ModuleMember Prereqs


Function SQLQuery($query) {    
    
    $Database = $Config.Database
    $ServerInstance = $Config.ServerInstance
    $Params = @{
        Database        = $Database
        ServerInstance  = $ServerInstance
        QueryTimeout    = 60        
        #OutputSqlErrors = $true
        Query           = $query
    }
    Try {
        Invoke-Sqlcmd2 @Params
    }
    Catch {
        Write-Host "Error Executing SQL Query : `n `n Database: $($Database) `n ServerInstance: $($ServerInstance) `n `n $($query) `n" 
        Write-Error $_ -ErrorAction Stop 
    }
}

Export-ModuleMember SQLQuery

Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

Export-ModuleMember Log

Function SQLFormatter($data) {
    #"Formatting Data..."
    $Keys = ($data | Get-Member -MemberType NoteProperty).Name
    #"Found $($Keys.Count) Unique Keys"
    ForEach ($Key in $Keys) {
        $KeyVal = $($data.$($Key))
        IF ($KeyVal -eq $null) {
            #Write-Host "VALUE IS NULL: Setting NULL"; $KeyVal = "NULL"
        }
        #Write-Host "KEY = $($Key); VALUE = $KeyVal; TYPE = $($KeyVal.GetType().name); LENGTH = $($KeyVal.length)"
        IF ($KeyVal.length -le 0) {
            #"Zero Length Setting NULL"
            $KeyVal = "NULL"
        }
        ELSEIF ($KeyVal.GetType().name -eq "DBNull") {
            #"DBNull Setting NULL"
            $KeyVal = "NULL"
        }

        IF (([string]$KeyVal).ToUpper() -eq "TRUE") {
            #"Converting Boolean"
            $KeyVal = 1
        }

        ELSEIF (([string]$KeyVal).ToUpper() -eq "FALSE") {
            #"Convererting Boolean"
            $KeyVal = 0
        }

        IF ($KeyVal -match "[^\d]" -and $KeyVal -ne "NULL") {
            $KeyVal = "'$($KeyVal)'"
        }

        $data.$($Key) = $KeyVal -replace "''", "'"
        

    }

    $data
}
Export-ModuleMember SQLFormatter

Function PSObjectifier {
    [cmdletbinding()]
    Param(
        [parameter(
            Position = 0, 
            Mandatory = $true, 
            ValueFromPipeline = $true)]
        $data
    )
    Process {
        $properties = ($data | Get-Member -MemberType Property).name
        $array = @()
        ForEach ($row in $data) {
            $obj = New-Object psobject
            ForEach ($prop in $properties) {
                #if($row.$($prop).GetType().name -eq "DBNull"){$row.$($prop) = ''}
                $obj | Add-Member -NotePropertyName $prop -NotePropertyValue $row.$($prop)
            }
            $array += $obj
        }
        $array
    }
}

Export-ModuleMember PSObjectifier

Function ExportWSToCSV ($excelFile, $csvLoc) {
    $excelFile
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile.FullName)
    foreach ($ws in $wb.Worksheets) {
        $n = $ws.Name
        "Saving to: $csvLoc\$n.csv"
        $ws.SaveAs("$csvLoc\$n.csv", 6)
    }
    $E.Quit()
}

Export-ModuleMember ExportWSToCSV

Function Merge-CSVFiles($CSVPath, $XLOutput) {
    $csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv | Sort-Object -Descending
    $Excel = New-Object -ComObject excel.application 
    $Excel.visible = $false
    $Excel.sheetsInNewWorkbook = $csvFiles.Count
    $workbooks = $excel.Workbooks.Add()
    $CSVSheet = 1

    Foreach ($CSV in $Csvfiles) {
        $worksheets = $workbooks.worksheets
        $CSVFullPath = $CSV.FullName
        $SheetName = ($CSV.name -split "\.")[0]
        $worksheet = $worksheets.Item($CSVSheet)
        $worksheet.Name = $SheetName
        $TxtConnector = ("TEXT;" + $CSVFullPath)
        Log "Merging Worksheet $SheetName"
        $CellRef = $worksheet.Range("A1")
        $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)
        $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
        $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
        $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
        $worksheet.QueryTables.item($Connector.name).delete()
        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        $CSVSheet++
    }

    $workbooks.SaveAs($XLOutput, 51)
    $workbooks.Saved = $true
    $workbooks.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Export-ModuleMember Merge-CSVFiles

Function SensorTypeSet($data){
    $SensorTypeSetSQL = @"
    DECLARE @RC int
    DECLARE @Sensor_Type_CD bigint = $($data.Sensor_Type_CD)
    DECLARE @Device_Type_CD bigint = $($data.Device_Type_CD)
    DECLARE @Sensor_Name nvarchar(100) = '$($data.Sensor_Name)'
    DECLARE @Sensor_Short_Name nvarchar(15) = '$(IF($data.Sensor_Short_Name.length -ge 15){$($data.Sensor_Type_CD.Substring(0,15))}ELSE{$($data.Sensor_Short_Name)})'
    DECLARE @Measurement_Type_CD int = $($data.Measurement_Type_CD)
    DECLARE @Last_Updated_By bigint = 616
    DECLARE @Is_Deleted bit = $(IF($data.Is_Deleted -eq "FALSE"){0}ELSE{1})
    DECLARE @Is_Critical bit = $(IF($data.Is_Critical -eq "FALSE"){0}ELSE{1})
    DECLARE @Rounding_Precision int = $($data.Rounding_Precision)
    DECLARE @Deadband_Percentage decimal(6,3) = $($data.Deadband_Percentage)
    DECLARE @Disable_History bit = $(IF($data.Disable_History -eq "FALSE"){0}ELSE{1})
    DECLARE @External_Tag_Format nvarchar(100) = $($data.External_Tag_Format)
    DECLARE @Setpoint_Timeout_Seconds int = $($data.Setpoint_Timeout_Seconds)
    DECLARE @Description nvarchar(1000) = $($data.Description)
    DECLARE @Package_PG uniqueidentifier = NULL
    DECLARE @Package_Name nvarchar(100) = NULL
    DECLARE @Sensor_Category_Type_CD bigint = $($data.Sensor_Category_Type_CD)
    DECLARE @Last_Updated_Document_ID bigint = NULL

    -- TODO: Set parameter values here.

    EXECUTE @RC = [cdm].[Device_Sensor_Type_Set] 
    @Sensor_Type_CD ,@Device_Type_CD ,@Sensor_Name ,@Sensor_Short_Name ,@Measurement_Type_CD 
    ,@Last_Updated_By ,@Is_Deleted ,@Is_Critical ,@Rounding_Precision ,@Deadband_Percentage 
    ,@Disable_History ,@External_Tag_Format ,@Setpoint_Timeout_Seconds 
    ,@Description ,@Package_PG ,@Package_Name ,@Sensor_Category_Type_CD ,@Last_Updated_Document_ID
    GO
"@
$SensorTypeSetSQL
}

Export-ModuleMember SensorTypeSet

