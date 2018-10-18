$scriptDir = split-path -parent $MyInvocation.MyCommand.Definition

$Config = (Get-Content "$scriptDir\buildconfig.json") -join "`n" | ConvertFrom-Json


Function ExecQuery ($Query){
    $sqlparams = @{
        ServerInstance = $Config.ServerInstance
        Database = $Config.Database
        QueryTimeout = 30
        OutputSQLErrors = $True
        Query = $Query
        
    }
    Invoke-Sqlcmd @sqlparams
}

Function InsertDevice ($params){
    $DeviceInsertSQL = @"
    DECLARE @Device_ID uniqueidentifier
    declare @now datetime2(7) = SYSUTCDATETIME()
    EXEC [cdm].[Device_Insert] 
    @Partition_ID = $($params.Partition_ID), 
    @Device_Type_CD = $($config.LeafDeviceType), 
    @Device_Name = 'Leaf Metrics', 
    @Host_Name = 'leafmetrics', 
    @IP_Address = NULL,
    @Port_Number = NULL, 
    @Device_State_CD = 1, 
    @Device_Up_Date = @now, 
    @Manufacturer_Name = NULL, 
    @Model_Number = NULL, 
    @PDN_KW_Rating = NULL, 
    @ERN_KW_Rating = NULL, 
    @Last_Updated_By = 616, 
    @Is_Deleted = 0, 
    @Date_In_Service = NULL, 
    @Main_Application_Purpose = NULL, 
    @Business_Owner = NULL, 
    @Estimate_Power_Usage = NULL, 
    @Operating_System = NULL, 
    @Is_Leased = 0, 
    @Lease_Expiration_Date = NULL,
    @Start_Number_U = NULL, 
    @Rack_Number_U = NULL, 
    @Comments = NULL, 
    @Is_Inventory_Only = 0, 
    @External_Identifier = NULL, 
    @Serial_Number = NULL, 
    @Version_ID = NULL, 
    @Section_Type_CD = NULL,
    @Manufacturer_Serial_Number = NULL,
    @MAC_Address = NULL,
    @Support_Vendor_ID = NULL,
    @External_Monitoring_System = NULL,
    @Maintenance_Interval_Type_CD = NULL,
    @Last_Maintenance_Date = NULL,
    @Description = NULL,
    @Last_Updated_Document_ID = NULL,
    @Device_Width = NULL,
    @Device_Depth = NULL,
    @Rack_Allow_Single_Post_Mount = 1,
    @Rack_Single_Post_Index = NULL,
    @Rack_Is_Mounted_On_Rear = 0,
    @Heat_Dissipation_BTU_Min = NULL,
    @Heat_Dissipation_BTU_Max = NULL,
    @Estimate_Power_Usage_Min = NULL,
    @Weight = NULL,
    @Rack_Start_Date = NULL,
    @Rack_End_Date = NULL,
    @Device_ID = @Device_ID OUTPUT
   
    SELECT @Device_ID as Device_ID
   
    GO
"@
$DeviceInsertSQL
}

Function InsertSensor ($params, $type, $URI){
    $SensorInsertSQL = @"
    DECLARE @Device_Sensor_ID uniqueidentifier,@SP_Return_CD bigint;
    EXEC cdm.Device_Sensor_Insert
    @Device_ID = '$($params.Device_ID)'
    ,@Monitor_Type_CD = 5
    ,@Operation_Type_CD = 2
    ,@Sensor_Type_CD = $($type.TypeCD)
    ,@External_Tag_Name = NULL
    ,@Reading_Type_CD = 1
    ,@Last_Updated_By = 616
    ,@External_Tag_Name_Setpoint = NULL
    ,@URI = '$($URI)'
    ,@Conversion_Formula_CD = NULL
    ,@Polling_Interval_Seconds = NULL
    ,@Sensor_Timeout_Seconds = NULL
    ,@URI_Setpoint = NULL
    ,@Comments = NULL
    ,@Disable_History = 0
    ,@Is_Deleted = 0
    ,@Setpoint_Min_Point = NULL
    ,@Setpoint_Max_Point = NULL
    ,@Is_Alarmed = 0
    ,@Is_Visible = 1
    ,@Ignore_Bad_Quality = 0
    ,@Bad_Quality_Priority_CD = 0
    ,@Comm_Failure_Priority_CD = 0
    ,@Bit_Type_CD = NULL
    ,@Description = NULL
    ,@Last_Updated_Document_ID = NULL
    ,@Device_Sensor_ID = @Device_Sensor_ID OUT
    ,@SP_Return_CD = @SP_Return_CD OUT;
   SELECT @Device_Sensor_ID as Device_Sensor_ID
   GO  
"@
$SensorInsertSQL
}

$LeafPartitionsQuery = @"
SELECT 
CPP.Full_Path,
P.Partition_ID,
P.Module_ID,
P.Partition_Name,
P.Partition_Short_Name,
PT.Description
FROM dbo.Partitions P
INNER JOIN dbo.Partition_Types PT ON P.Partition_Type_CD = PT.Partition_Type_CD
LEFT  JOIN dbo.Cache_Partition_Paths CPP ON P.Partition_ID = CPP.Partition_ID AND CPP.Path_Type = 2
WHERE P.Partition_Type_CD IN ($($Config.LeafPartitionTypes -join ","))
AND P.Is_Deleted = 0
"@

$LeafPartitions = ExecQuery -Query $LeafPartitionsQuery
$MissingLeafDevices = @()
$MismatchedURI = @()
ForEach($Partition in $LeafPartitions){
    $LeafDeviceQuery = @"
    SELECT 
    * 
    FROM dbo.Devices 
    WHERE Partition_ID = $($Partition.Partition_ID) 
    AND Device_Type_CD = $($Config.LeafDeviceType)
    AND Is_Deleted = 0
"@
    [array]$LeafDevice = ExecQuery -Query $LeafDeviceQuery
    IF($LeafDevice.count -ge 1){
        Write-Host -ForegroundColor Cyan "Found: $($LeafDevices.Count) :$($Partition.Partition_ID) : $($Partition.Full_Path)"      
    }
    ELSE{
        Write-Host -ForegroundColor magenta "No Leaf Device - Inserting : $($Partition.Partition_ID) : $($Partition.Full_Path)"
        $NewDevice = ExecQuery -Query $(InsertDevice -params $Partition) 
        $LeafDevice = $NewDevice
    }

    IF($LeafDevice.count -ge 1){
        ForEach($sensortype in $Config.LeafSensorTypes){
            IF($sensortype.VarName -ne $null){
                    $VarName = $sensortype.VarName
                    $VarExpand = Invoke-Expression -Command $VarName
                    $URI = $sensortype.URI.Replace($sensortype.PlaceHolder, $VarExpand)
            }
            ELSE{
                $URI = $sensortype.URI
            }
            
            $SensorExistQuery = "SELECT * FROM dbo.Device_Sensors WHERE Device_ID = '$($LeafDevice.Device_ID)' AND Sensor_Type_CD = $($SensorType.TypeCD)"
            $SensorExist = ExecQuery -Query  $SensorExistQuery     
            IF($SensorExist){
                IF($SensorExist.URI -eq $URI){
                    #do nothing
                }
                ELSE{
                    "URI Mismatch $($SensorType.TypeName)"
                    "old: $($SensorExist.URI)"
                    "new: $($URI)"
                    $MismatchedURI += [PSCustomObject]@{
                        path = $Partition.Full_Path
                        ID = $SensorExist.Device_Sensor_ID
                        SensorType = $sensortype.TypeName
                        OldURI = $SensorExist.URI
                        NewURI = $URI
                    }
                }
            }ElSE{
                "Inserting Sensor $($sensortype.TypeName)";""
                
                #(InsertSensor -params $LeafDevice -type $SensorType -URI $URI)
                $Insert = ExecQuery -Query (InsertSensor -params $LeafDevice -type $SensorType -URI $URI)
                "Added : $($Insert.Device_Sensor_ID) : $URI"
            }      
        }
    }
    ELSE{
        #Still Missing
        "Error Still Missing Device"
        RETURN
    }
}
""
"Missing Leaf Devices: $($MissingLeafDevices.count)"
""
"Mismatched URIs: $($MismatchedURI.count)"
$MismatchedURI | Format-Table * -AutoSize