#Library of functions to drive Deployment Tools

Function Logo() {
    Write-Host -ForegroundColor Blue @"
                                                                                                                                                         
    $(Split-Path -Leaf $MyInvocation.ScriptName)
"@
}
    
Export-ModuleMember Logo

Function Prereqs ($config) {
    $Repository = $config.PSModule.Repository
    Try {
        $Ping = $null
        $Ping = (Invoke-WebRequest -Uri $Repository).StatusCode
    }
    Catch { $Ping = $_ }
    $Modules = $config.PSModule.Modules
    
    if ($Ping.GetType().name -eq "ErrorRecord") {
        Write-Host -ForegroundColor Cyan "Error Encountered Connecting to Repository : $Repository"
        Write-Host -ForegroundColor Red $Ping.Exception
    }
    elseif ($Modules -eq $null -or $Modules.Count -le 0) {
        Write-Host -ForegroundColor Cyan "Error No Modules Listed in Config"
    }
    else {
        ForEach ($Module in $Modules) {
            $installed = $null
            $installed = Get-Module -ListAvailable -Name $Module
            $loaded = $null
            $loaded = Get-Module -Name $Module
            if ($installed -and $loaded) {
                Write-Host -ForegroundColor Cyan "Module: $Module - Already Loaded"
            }
            elseif ($installed -and $loaded -ne $true) {
                Write-Host -ForegroundColor Green "Module: $Module - Loading..."
                Try {
                    Import-Module $Module
                }
                Catch {$_ | Out-Null}
            }
            else {
                Write-Host -ForegroundColor Yellow "Module: $Module - Installing..."
                Install-Module $Module -Force -Repository PSGallery -Confirm:$false
                Try {
                    Import-Module $Module
                }
                Catch {$_ | Out-Null}
            }
        }        
    }
    ""
}

Export-ModuleMember Prereqs

Function LoginUser ($config) {
    $user = $config.username
    $password = $config.password
    $URI = "$($config.WebHost)/$($config.LoginMethod)"
    $winver = (Get-WmiObject Win32_OperatingSystem).Version 
    if([version]$winver -ge [version]"10.0.0.0"){
        $provider = Get-NetIPInterface | Where-Object { 
            $_.ConnectionState -eq "Connected" 
        } | Sort-Object InterfaceMetric | Select-Object -First 1

        $adapter = Get-NetAdapter | Where-Object { 
            $_.IfIndex -eq $Provider.IfIndex 
        }
        $macaddress = $adapter.MacAddress
    }else{
        $macaddress = (Get-WmiObject Win32_NetworkAdapter | where-Object{
            $_.MACAddress -ne $null -and
            $_.PhysicalAdapter -eq $true
        } | Select-Object -first 1).MacAddress.Replace(":","-")
    }
    $JSON = $Config.JSONLoginUser
    $JSON.LoginName = $user
    $JSON.Password = $password
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.MacAddresses = $macaddress
    $JSON.ComputerName = $env:COMPUTERNAME
    $JSON.UserAgent = [System.Environment]::OSVersion.VersionString

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    Invoke-RestMethod @params
}

Export-ModuleMember LoginUser

Function GetSelectMethod ($config, $user, $service, $method) {
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetSelectMethod

Function GetAppReleases ($config, $user) {
    $service = $config.MiscMethod
    $method = "App_Get_Releases"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Application_GUID("ca45379e-e857-4606-a956-bdf35f02cb81") -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetAppReleases


Function PartitionSearch ($config, $user, $searchstring) {
    $service = $config.PartitionMethod
    $method = "PH_Partition_Search"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Search_Text_Partition($searchstring) -Force
    $JSON | Add-Member Partition_ID($null) -Force
    $JSON | Add-Member Region_ID($null) -Force
    $JSON | Add-Member Application_CD(1) -Force
    $JSON | Add-Member Login_ID($user.Login_ID) -Force

    if($JSON.partitionID){
        $JSON.partitionID = $Null
    }

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember PartitionSearch

Function SensorSearch ($config, $user, $searchstring, $Partition_ID) {
    $service = $config.MiscMethod
    $method = "PH_Sensor_Search"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Search_Text_Partition($null) -Force
    $JSON | Add-Member Search_Text_Device($null) -Force
    $JSON | Add-Member Search_Text_Sensor($searchstring) -Force
    $JSON | Add-Member Partition_ID($Partition_ID) -Force
    $JSON | Add-Member Region_ID($null) -Force
    $JSON | Add-Member Application_CD(1) -Force
    $JSON | Add-Member Login_ID($user.Login_ID) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember SensorSearch

Function GetSensor ($config, $user, $sensor) {
    $service = $config.DeviceMethod
    $method = "SELECT_713"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member sensorID($sensor)

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetSensor


Function GetPartition ($config, $user, $partition) {
    $service = $config.PartitionMethod
    $method = "SELECT_103"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member partitionID($partition) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetPartition

Function GetPartitionChildren ($config, $user, $partition) {
    $service = $config.PartitionMethod
    if($Config.Release.Version_Number -ge [version]"3.7.0.0"){
    $method = "Partition_Get_Siblings"
    }else{
        $method = "Partition_Get_Children"
    }
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Partition_ID($partition.Partition_ID) -Force
    $JSON | Add-Member Location_ID($partition.Location_ID) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetPartitionChildren


Function GetDevices ($config, $user, $partition) {
    $Service = $config.DeviceMethod
    $Method = "SELECT_15003"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member -NotePropertyName "partitionID" -NotePropertyValue $partition -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetDevices

Function GetDeviceProperties ($config, $user, $deviceID) {
    $Service = $config.DeviceMethod
    $Method = "SELECT_181"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member -NotePropertyName "deviceID" -NotePropertyValue $deviceID -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetDeviceProperties

Function SetSensorType ($data, $config, $user) {
    #To do add logic to determine if sensor type cd is present in template
    $service = $config.DeviceMethod
    if ($data.Sensor_Type_CD -eq $Null -or $data.Sensor_Type_CD -eq "NULL") {
        $IsNew = $true
        $method = "AddDeviceSensorType"
        $data.Sensor_Type_CD = 0
    }
    else {
        $IsNew = $false
        $method = "ModifyDeviceSensorType"
    }
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = if([version]$config.Release.Version_Number -lt [version]"3.7.0.0"){$config.JSONSetSensorType332}else{$config.JSONSetSensorType}
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    #Set Poco properties from template
    $JSON.poco.Sensor_Type_CD = if ($data.Sensor_Type_CD -eq $Null -or
        $data.Sensor_Type_CD -eq "NULL") {
        0
    }
    else {$data.Sensor_Type_CD}
    $JSON.poco.Device_Type_CD = [int]$data.Device_Type_CD
    $JSON.poco.Sensor_Name = if($data.Sensor_Name){$data.Sensor_Name}else{$data.Sensor_Type}
    $JSON.poco.Sensor_Short_Name = if($data.Sensor_Short_Name){$data.Sensor_Short_Name}else{$data.Sensor_Type_Short}
    if ($JSON.poco.Sensor_Short_Name.length -gt 15) {
        Write-Host -ForegroundColor Yellow "Sensor_Short_Name: [$($data.Sensor_Short_Name)]LEN($($data.Sensor_Short_Name.length)) exceeds 15 char"
        Write-Error "Sensor_Short_Name: [$($data.Sensor_Short_Name)]LEN($($data.Sensor_Short_Name.length)) exceeds 15 char"
    }
    $JSON.poco.Measurement_Type_CD = if ($data.Measurement_Type_CD) {
        [int]$data.Measurement_Type_CD
    }
    elseif ($data.Measurement_Type_CD.length -eq 0 -and $data.Measurement_Type -eq $Null) {
        0 #default if field is left black "intentionally left blank"
    }
    else {
        ($Measurement_Types.Result | Where-Object {
                $_.Description -eq $data.Measurement_Type
            }).Measurement_Type_CD | Select-Object -First 1
        
    }
    if ($JSON.poco.Measurement_Type_CD -eq $Null) {
        Write-Host -ForegroundColor Yellow "Error: unable to match measurement type: $($data.Measurement_Type)"
        Write-Error "Error: unable to match measurement type: $($data.Measurement_Type)"
    }
    $JSON.poco.Last_Update_date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.poco.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    if ($data.Is_Critical -eq $Null) { 
        $data | add-member Is_Critical($false) -Force 
    }elseif($data.Is_Critical.length -lt 1){
        $data.Is_Critical = "FALSE"
    }
    $JSON.poco.Is_Critical = [System.Convert]::ToBoolean($data.Is_Critical)
    $JSON.poco.Rounding_Precision = if($data.Rounding){
        if($data.Rounding -eq $null -or $data.Rounding -eq "NULL"){
            [int]0
        }else{
            [int]$data.Rounding
        }
    }elseif($data.Rounding_Precision -ne $Null -and $data.Rounding_Precision.Length -gt 0){
        [int]$data.Rounding_Precision
    }else{
        [int]0
    }
    if ($data.Deadband_Percentage -eq $Null) { $data | add-member Deadband_Percentage($Null) -Force }
    $JSON.poco.Deadband_Percentage = if ($data.Deadband_Percentage -eq "NULL" -or $data.Deadband_Percentage -eq $Null) {$null}else {$data.Deadband_Percentage}
    if ($data.External_Tag_Format -eq $Null) {$data | add-member External_Tag_Format($null) -Force}
    $JSON.poco.External_Tag_Format = if ($data.External_Tag_Format -eq "NULL" -or $data.External_Tag_Format -eq $Null) {$null}else {$data.External_Tag_Format}
    if ($data.Disable_History -eq $Null) {
         $data | add-member Disable_History($false) -Force 
    }elseif($data.Disable_History.length -lt 1){
        $data.Disable_History = "FALSE"
    }
    $JSON.poco.Disable_History = [System.Convert]::ToBoolean($data.Disable_History)
    $JSON.poco.Setpoint_Timeout_Seconds = if ($data.SP_Timeout -eq "NULL") {$null}else {$data.SP_Timeout}
    $JSON.poco.Sensor_Category_Type_CD = if ($data.CatType -eq "NULL") {$null}else {$data.CatType}
    $JSON.poco.Description = if ($data.Description -eq "NULL") {$null}else {$data.Description}
    $JSON.poco.IsNew = $IsNew

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }
    InvokeRest -params $params
}

Export-ModuleMember SetSensorType

Function SetSensorType332 ($data, $config, $user) {
    #To do add logic to determine if sensor type cd is present in template
    $service = $config.DeviceMethod
    if ($data.Sensor_Type_CD -eq $Null -or $data.Sensor_Type_CD -eq "NULL") {
        $IsNew = $true
        $method = "AddDeviceSensorType"
        $data.Sensor_Type_CD = 0
    }
    else {
        $IsNew = $false
        $method = "ModifyDeviceSensorType"
    }
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSetSensorType332
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    #Set Poco properties from template
    $JSON.poco.Sensor_Type_CD = if ($data.Sensor_Type_CD -eq $Null -or
        $data.Sensor_Type_CD -eq "NULL") {
        0
    }
    else {$data.Sensor_Type_CD}
    $JSON.poco.Device_Type_CD = [int]$data.Device_Type_CD
    $JSON.poco.Sensor_Name = $data.Sensor_Name
    $JSON.poco.Sensor_Short_Name = $data.Sensor_Short_Name
    if ($JSON.poco.Sensor_Short_Name.length -gt 15) {
        Write-Host -ForegroundColor Yellow "Sensor_Short_Name: [$($data.Sensor_Short_Name)]LEN($($data.Sensor_Short_Name.length)) exceeds 15 char"
        Write-Error "Sensor_Short_Name: [$($data.Sensor_Short_Name)]LEN($($data.Sensor_Short_Name.length)) exceeds 15 char"
    }
    $JSON.poco.Measurement_Type_CD = if ($data.Measurement_Type_CD) {
        [int]$data.Measurement_Type_CD
    }
    elseif ($data.Measurement_Type_CD.length -eq 0 -and $data.Measurement_Type -eq $Null) {
        0 #default if field is left blank "intentionally left blank"
    }
    else {
        ($Measurement_Types.Result | Where-Object {
                $_.Description -eq $data.Measurement_Type
            }).Measurement_Type_CD | Select-Object -First 1
        
    }
    if ($JSON.poco.Measurement_Type_CD -eq $Null) {
        Write-Host -ForegroundColor Yellow "Error: unable to match measurement type: $($data.Measurement_Type)"
        Write-Error "Error: unable to match measurement type: $($data.Measurement_Type)"
    }
    $JSON.poco.Last_Update_date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.poco.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    if ($data.Is_Critical -eq $Null) { $data | add-member Is_Critical($false) -Force }
    $JSON.poco.Is_Critical = [System.Convert]::ToBoolean($data.Is_Critical)
    $JSON.poco.Rounding_Precision = if($data.Rounding -eq $Null -or $data.Rounding -eq "NULL"){0}else{$data.Rounding}
    if ($data.Deadband_Percentage -eq $Null) { $data | add-member Deadband_Percentage($Null) -Force }
    $JSON.poco.Deadband_Percentage = if ($data.Deadband_Percentage -eq "NULL" -or $data.Deadband_Percentage -eq $Null) {$null}else {$data.Deadband_Percentage}
    if ($data.External_Tag_Format -eq $Null) {$data | add-member External_Tag_Format($null) -Force}
    $JSON.poco.External_Tag_Format = if ($data.External_Tag_Format -eq "NULL" -or $data.External_Tag_Format -eq $Null) {$null}else {$data.External_Tag_Format}
    if ($data.Disable_History -eq $Null) { $data | add-member Disable_History($false) -Force }
    $JSON.poco.Disable_History = [System.Convert]::ToBoolean($data.Disable_History)
    $JSON.poco.Setpoint_Timeout_Seconds = if ($data.SP_Timeout -eq "NULL") {$null}else {$data.SP_Timeout}
    $JSON.poco.Sensor_Category_Type_CD = if ($data.CatType -eq "NULL") {$null}else {$data.CatType}
    $JSON.poco.IsNew = $IsNew

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }
    InvokeRest -params $params
}

Export-ModuleMember SetSensorType332


Function AddDeviceType ($data, $config, $user) {
    #To do add logic to determine if sensor type cd is present in template
    $service = $config.DeviceMethod
    $method = "AddDeviceType"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONAddDeviceType
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    #Set Poco properties from template
    $JSON.poco.Device_Type_Name = $data.Device_Type
    $JSON.poco.Device_Type_Short_Name = $data.Device_Type_Short
    $JSON.poco.Last_Updated_By = $user.Login_ID
    $JSON.poco.Last_Update_date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.poco.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    $JSON.poco.Description = if ($data.Description -eq "NULL") {$null}else {$data.Description}

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember AddDeviceType

Function GetDeviceSensors ($config, $user, $deviceID) {
    $Service = $config.DeviceMethod
    $Method = "SELECT_858"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member -NotePropertyName "deviceID" -NotePropertyValue $deviceID -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetDeviceSensors

Function InsertSensor ($data, $config, $user, $device, $sensortype) {
    $service = $config.DeviceMethod
    $method = "Device_Sensor_Insert"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONInsertSensor
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    #Set Poco properties from template
    $JSON.Device_ID = $device.Device_ID
    $JSON.Monitor_Type_CD = ($Monitor_Types.Result | Where-Object { 
            $_.Description -eq $data.Monitor_Type 
        }).Monitor_Type_CD

    if ($JSON.Monitor_Type_CD -eq $null) {
        Write-Error "Must Enter a Valid Monitor Type: $(
            $Monitor_Types.Result | Sort-Object Monitor_Type_CD | ForEach-Object {
                "$($_.Monitor_Type_CD):$($_.Description),"
            })"
    }
    $JSON.Sensor_Type_CD = $sensortype.Sensor_Type_CD
    $JSON.Reading_Type_CD = ($Reading_Types.Result | Where-Object { 
            $_.Description -eq $data.Reading_Type
        }).Reading_Type_CD

    if ($JSON.Reading_Type_CD.length -lt 1 -or $JSON.Reading_Type_CD -eq $null) {
        Write-Host -ForegroundColor Yellow "You must select a valid Reading Type:"
        Write-Host "$($Reading_Types.Result | Sort-Object Reading_Type_CD | ForEach-Object { "[ $($_.Reading_Type_CD) : $($_.Description) ],"})"
        Write-Error "Error: Method: InsertSensor Message: Invalid Reading_Type_CD"
    }        
    $JSON.Last_Updated_By = $user.Login_ID
    $JSON.URI = if ($data.BQ -ne "NULL" -and $data.BQ -ne $null -and $data.BQ.length -gt 1) {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)&QualityID=$($data.BQ)"
    }
    else {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)"
    }
    $JSON.URI_Setpoint = if ($data.SP -eq $true) {$JSON.URI}else {$null}
    $JSON.Conversion_Formula_CD = if ($data.Convert -eq $Null -or $data.Convert -eq "NULL") {$null}ELSE {$data.Convert}
    $JSON.Bad_Quality_Priority_CD = if ($data.BQ_Pri -eq $null) {0}ELSE {$data.BQ_Pri}
    $JSON.Comm_Failure_Priority_CD = if ($data.CF_Pri -eq $null) {0}ELSE {$data.CF_Pri}
    $JSON.Setpoint_Min_Point = if ($data.SPMin -eq "NULL") {$Null}else {$data.SPMin}
    $JSON.Setpoint_Max_Point = if ($data.SPMax -eq "NULL") {$null}else {$data.SPMax}
    $JSON.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    $JSON.Polling_Interval_Seconds = if ($data.Polling -eq "NULL") {$null}else {$data.Polling}

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember InsertSensor

Function UpdateSensor ($data, $config, $user, $device, $sensortype) {
    $service = $config.DeviceMethod
    $method = "Device_Sensor_Update"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONUpdateSensor
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    #Set Poco properties from template
    $JSON.Device_ID = $device.Device_ID
    $JSON.Device_Sensor_ID = $data.Device_Sensor_ID
    $JSON.Monitor_Type_CD = ($Monitor_Types.Result | Where-Object { 
            $_.Description -eq $data.Monitor_Type 
        }).Monitor_Type_CD
    if ($data.Monitor_Type -eq $null) {$JSON.Monitor_Type_CD = $data.Monitor_Type_CD}

    if ($JSON.Monitor_Type_CD -eq $null) {
        Write-Error "Must Enter a Valid Monitor Type: $(
            $Monitor_Types.Result | Sort-Object Monitor_Type_CD | ForEach-Object {
                "$($_.Monitor_Type_CD):$($_.Description),"
            })"
    }
    $JSON.Sensor_Type_CD = $sensortype.Sensor_Type_CD
    $JSON.Reading_Type_CD = ($Reading_Types.Result | Where-Object { 
            $_.Description -eq $data.Reading_Type
        }).Reading_Type_CD

    if ($data.ReadingType -eq $null -and $data.Reading_Type_CD) {
        $JSON.Reading_Type_CD = $data.Reading_Type_CD
    }

    if ($JSON.Reading_Type_CD.length -lt 1 -or $JSON.Reading_Type_CD -eq $null) {
        Write-Host -ForegroundColor Yellow "You must select a valid Reading Type:"
        Write-Host "$($Reading_Types.Result | Sort-Object Reading_Type_CD | ForEach-Object { "[ $($_.Reading_Type_CD) : $($_.Description) ],"})"
        Write-Error "Error: Method: InsertSensor Message: Invalid Reading_Type_CD"
    }        
    
    if($Config.Release.Version_Number -ge [version]"3.7.0.0"){
    $JSON.Last_Updated_By = $data.Last_Updated_By
    }else{
        $JSON = $JSON | Select-Object * -ExcludeProperty Last_Updated_By,Last_Updated_Document_ID
        $JSON | Add-Member Login_ID($User.Login_ID) -Force
    }
    $JSON.URI = if ($data.BQ -ne $null -and $data.BQ -ne "NULL" -and $data.BQ.length -gt 1) {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)&QualityID=$($data.BQ)"
    }
    else {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)"
    }$JSON.URI = if ($data.BQ -ne $null -and $data.BQ -ne "NULL" -and $data.BQ.length -gt 1) {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)&QualityID=$($data.BQ)"
    }
    else {
        "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.Monitor_Item_ID)"
    }
    $JSON.URI_Setpoint = if ($data.SP -eq $true -or $data.SP -eq "true") {
            $JSON.URI
        }elseif($data.SP -ne $true -and $data.SP -ne "TRUE" -and $data.SP.length -gt 5){
            "$($data.Monitor_URL)$($data.Monitor_Server_ID)$($data.SP)"
        }else {
            $null
        }   
    $JSON.Conversion_Formula_CD = if ($data.Convert -eq $Null -or $data.Convert -eq "NULL") {$null}ELSE {$data.Convert}
    $JSON.Bad_Quality_Priority_CD = if ($data.BQ_Pri -eq $null) {0}ELSE {$data.BQ_Pri}
    $JSON.Comm_Failure_Priority_CD = if ($data.CF_Pri -eq $null) {0}ELSE {$data.CF_Pri}
    $JSON.Setpoint_Min_Point = if ($data.SPMin -eq "NULL") {$Null}else {$data.SPMin}
    $JSON.Setpoint_Max_Point = if ($data.SPMax -eq "NULL") {$null}else {$data.SPMax}
    $JSON.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    $JSON.Polling_Interval_Seconds = if ($data.Polling -eq "NULL") {$null}else {$data.Polling}

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember UpdateSensor


Function InsertDevice ($data, $partition, $config, $user) {
    $service = $config.DeviceMethod
    $method = "Device_Insert"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONInsertDevice
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Last_Updated_By = $user.Login_ID
    $JSON.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    $JSON.Partition_ID = $partition.Partition_ID
    $JSON.Device_Type_CD = $data.Device_Type_CD
    $JSON.Device_Up_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.Device_Name = $data.Device_Name
    $JSON.Host_Name = (
        $data.Device_Name + 
        "." +
        $partition.partition_short_name +    
        ".anywhere.corp").Replace(" ", "").ToLower()
    

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember InsertDevice

Function SetPartition ($data, $config, $user) {
    $service = $config.PartitionMethod
    $method = "Partition_Set"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONPartitionSet
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Last_Updated_By = $user.Login_ID
    $JSON.Is_Deleted = [System.Convert]::ToBoolean($data.Is_Deleted)
    $JSON.Partition_ID = $data.Partition_ID
    $JSON.Activation_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.Location_ID = $data.Location_ID
    $JSON.Parent_Partition_ID = $data.Parent_Partition_ID
    $JSON.Partition_Name = $data.Partition_Name
    $JSON.Partition_Short_Name = $data.Partition_Short_Name
    $JSON.Partition_Type_CD = $data.Partition_Type_CD
    $JSON.Status_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"

    if($data.Partition_Type_CD -eq 12){
        $JSON.Partition_Module_Type_CD = $data.PMT_CD
        $JSON.Power_CD = $data.PMT_PWR_CD
        $JSON.Redundancy_ERN_CD = $data.PMT_ERN_CD
        $JSON.Redundancy_PDN_CD = $data.PMT_PDN_CD
        $JSON.Optimization_CD = $data.PMT_OPT_CD
    }
    
    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember SetPartition

Function SetThresh ($data, $config, $user) {
    $service = $config.DeviceMethod
    $method = "Device_Sensor_Thresh_Set"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONDevice_Sensor_Thresh_Set
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Last_Updated_By = $user.Login_ID
    $JSON.Device_Sensor_ID = $data.Device_Sensor_ID
    $JSON.Sort_Order = $data.Sort_Order
    $JSON.Priority_CD = $data.Priority_CD
    $JSON.Min_Point = $data.Min_Point
    $JSON.Max_Point = $data.Max_Point

    if($Config.Release.Version_Number -lt [version]"3.7.0.0"){
        $JSON = $JSON | Select-Object * -ExcludeProperty Last_Updated_By,Last_Updated_Document_ID
        $JSON | Add-Member Login_ID($user.Login_ID) -Force
    }
    
    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember SetThresh

Function GetThresh ($data, $config, $user) {
    $service = $config.DeviceMethod
    $method = "SELECT_1568"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member sensorID($data.Device_Sensor_ID) -Force
     

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetThresh

Function DeleteThresh ($id) {
    $service = $config.DeviceMethod
    $method = "Device_Sensor_Threshold_Delete"
    #Set Rest connection params
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONThreshDelete
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Last_Updated_By = $user.Login_ID
    $JSON.Threshold_ID = $id.Threshold_ID
    
    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }
    
    InvokeRest -params $Params
}

Export-ModuleMember DeleteThresh

Function DestroyDevice ($config, $user, $device, $taskid) {
    $service = $Config.MiscMethod
    $method = "Task_Set"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONDestroyDevice
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Task_ID = $taskid.guid
    $JSON.Start_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.Parameter_XML = (
        '<?xml version="1.0" encoding="utf-8"?>' + 
        '<EntityDestructionParameters xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
        '<EntityTypeToDestroy>Device</EntityTypeToDestroy>' +
        "<EntityID>$($Device.Device_ID)</EntityID>" +
        "<RequestedBy>$($user.Login_ID)</RequestedBy>" +
        '</EntityDestructionParameters>')

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember DestroyDevice

Function DestroySensor ($config, $user, $sensor, $taskid) {
    $service = $Config.MiscMethod
    $method = "Task_Set"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONDestroyDevice
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Task_ID = $taskid.guid
    $JSON.Start_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"
    $JSON.Parameter_XML = (
        '<?xml version="1.0" encoding="utf-8"?>' + 
        '<EntityDestructionParameters xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
        '<EntityTypeToDestroy>Sensor</EntityTypeToDestroy>' +
        "<EntityID>$($sensor)</EntityID>" +
        "<RequestedBy>$($user.Login_ID)</RequestedBy>" +
        '</EntityDestructionParameters>')

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember DestroySensor


Function TaskProgress ($config, $user, $taskid) {
    $service = $Config.MiscMethod
    $method = "Task_Progress_Set"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONTaskProgressSet
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON.Task_ID = $taskid.guid
    $JSON.Progress_Date = "$([System.DateTime]::UtcNow.GetDateTimeFormats("O"))"

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember TaskProgress

Function GetTask ($config, $user, $taskid) {
    $service = $config.MiscMethod
    $method = "Task_Get"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Task_ID($taskid.guid) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetTask

Function TempThresh ($data) {
    $ThreshCompressed = @()
    if ($data.Thresh1 -and $data.Thresh1 -ne "NULL") {$ThreshCompressed += "1|$($data.Thresh1)"}
    if ($data.Thresh2 -and $data.Thresh2 -ne "NULL") {$ThreshCompressed += "2|$($data.Thresh2)"}
    if ($data.Thresh3 -and $data.Thresh3 -ne "NULL") {$ThreshCompressed += "3|$($data.Thresh3)"}
    if ($data.Thresh4 -and $data.Thresh4 -ne "NULL") {$ThreshCompressed += "4|$($data.Thresh4)"}
    if ($data.Thresh5 -and $data.Thresh5 -ne "NULL") {$ThreshCompressed += "5|$($data.Thresh5)"}

    if ($ThreshCompressed.Count -ge 1) {
        $ThreshArray = @()
        
        ForEach ($thresh in $ThreshCompressed) {
            $split = $thresh.split("|")
            
            $ThreshArray += [PSCustomObject]@{
                Sort_Order  = $split[0]
                Priority_CD = $split[1]
                Min_Point   = $split[2]
                Max_Point   = $split[3]
            }
        }
        $ThreshArray
    }
    else {
        $null
    }
}

Export-ModuleMember TempThresh

Function CompareSensorType ($config, $user, $sensortype, $template) {
    $match = $true
    $attributes = @()
    if($sensortype -and $template){
        #Compare Measurement Type
        $temp_measure_cd = ($Measurement_Types.Result | Where-Object {
            $_.Description -eq $template.Measurement_Type
        }).Measurement_Type_CD

        if($temp_measure_cd.Count -gt 1){
            $temp_measure = $Measurement_Types.Result | Where-Object {
                $_.Description -eq $template.Measurement_Type
            } | Out-Gridview -passthru -title "Select the correct measurement type for Sensor: $($sensortype.Sensor_Name)"
            $temp_measure_cd = $temp_measure.Measurement_Type_CD
        }

        if($temp_measure_cd -ne $sensortype.Measurement_Type_CD){
            $match = $false
            $oldMT = $Measurement_Types.Result | Where-Object {$_.Measurement_Type_CD -eq $sensortype.Measurement_Type_CD}
            $newMT = $Measurement_Types.Result | Where-Object {$_.Measurement_Type_CD -eq $temp_measure_cd}
            $attributes += "Measurement_Type_CD( old: $($oldMT.Measurement_Type_CD):$($oldMT.Measurement_Name) | new: $($newMT.Measurement_Type_CD):$($newMT.Measurement_Name) )"
        }
        #Compare Sensor Name
        if($template.Sensor_Type -ne $sensortype.Sensor_Name){
            $match = $false
            $attributes += "Sensor_Name( old: $($sensortype.Sensor_Name) | new: $($template.Sensor_Type))"
        }
        #Compare Sensor Short Name
        if($template.Sensor_Type_Short -ne $sensortype.Sensor_Short_Name){
            $match = $false
            $attributes += "Short_Name($($template.Sensor_Type_Short))"
        }
        #Compare Rounding Precision
        if($template.Rounding -ne $sensortype.Rounding_Precision){
            $match = $false
            $attributes += "Rounding($($template.Rounding))"
        }
    }
    [pscustomobject]@{
        match      = $match
        attributes = $attributes
    }
}

Export-ModuleMember CompareSensorType
Function CompareSensor ($config, $user, $sensor, $template) {
    $match = $true
    $attributes = @()
    if ($sensor -and $template) {
        #compare reading types
        $readingtype = $Reading_Types.Result | Where-Object {
            $_.Reading_Type_CD -eq $Sensor.Reading_Type_CD
        }
        if ($readingtype.Description -ne $template.Reading_Type) {
            $match = $false
            $attributes += "Reading_Type"
        }

        #compare Monitor Type
        $monitortype = $Monitor_Types.Result | Where-Object {
            $_.Monitor_Type_CD -eq $Sensor.Monitor_Type_CD
        }
        if ($monitortype.Description -ne $template.Monitor_Type) {
            $match = $false
            $attributes += "Monitor_Type"
        }
        #compare URI
        $tempuri = "$($template.Monitor_URL)$($template.Monitor_Server_ID)$($template.Monitor_Item_ID)$(if($template.BQ.Length -gt 5){"&QualityId=$($template.BQ)"})"
        if ($tempuri -ne $Sensor.URI) {
            $match = $false
            $attributes += "URI"
        }
        #compare URI Setpoint
        if ($template.SP -eq $true) {$setpoint = $tempuri}else {$setpoint = $null}
        if ($sensor.URI_Setpoint -ne $setpoint) {
            $match = $false
            $attributes += "URI_Setpoint"
        }
        #Compare Setpoint Min
        if ($Sensor.Setpoint_Min_Point -ne (toNull($template.SPMin))) {
            $match = $false
            $attributes += "Setpoint_Min_Point"
        }
        #Compare Setpoint Max
        if ($Sensor.Setpoint_Max_Point -ne (toNull($template.SPMax))) {
            $match = $false
            $attributes += "Setpoint_Max_Point"
        }
        #Compare Polling
        if ($Sensor.Polling_Interval_Seconds -ne (toNull($template.Polling))) {
            $match = $false
            $attributes += "Polling"
        }
        #Compare deleted
        if ($Sensor.Is_Deleted -ne [System.Convert]::ToBoolean($template.Is_Deleted)) {
            $match = $false
            $attributes += "Is_Deleted"
        }
        #compare Conversion Formula
        if ($Sensor.Conversion_Formula_CD -ne (toNull($template.Convert))) {
            $match = $false
            $attributes += "Conversion_Formula"
        }
        #Compare BQ Priority
        if ($template.BQ_Pri -and $Sensor.Bad_Quality_Priority_CD -ne $template.BQ_Pri) {
            $match = $false
            $attributes += "Bad_Quality_Priority"
        }
        #Compare CF Priority
        if ($template.CF_Pri -and $sensor.Comm_Failure_Priority_CD -ne $template.CF_Pri) {
            $match = $false
            $attributes += "Comm_Fail_Priority"
        }
    }
    [pscustomobject]@{
        match      = $match
        attributes = $attributes
    }
}

Export-ModuleMember CompareSensor
Function GetLastAppLog ($config, $user) {
    $service = $config.MiscMethod
    $method = "App_Logs_Get"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | Add-Member Is_Closed($false) -Force
    $JSON | Add-Member Lowest_Log_ID($null) -Force
    $JSON | Add-Member Requires_Attention($null) -Force
    $JSON | Add-Member Retrieval_Amount(1) -Force
    $JSON | Add-Member Search_Text($null) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    $Result = InvokeRest -params $Params
    if ($Result.ResultCode -eq 0 -and
        $Result.Result[0]) {
        $service = $config.MiscMethod
        $method = "App_Log_Long_Desc_Get"
        $URI = "$($config.WebHost)/$($service)$($method)"
        $JSON = $config.JSONAppLogLongDesc
        $JSON.Ticket = $user.Ticket
        $JSON.LoginID = $user.Login_ID
        $JSON.Environment = $config.Environment
        $JSON.AppID = $config.AppID  
        $JSON.Log_ID = $Result.Result[0].Log_ID

        $Params = @{
            Uri         = $URI
            Method      = "Post"
            Body        = ($JSON | ConvertTo-Json)
            ContentType = "application/json; charset=utf-8"
            ErrorAction = "Inquire"
        }
        
        $longdesc = InvokeRest -params $Params
        if ($longdesc.ResultCode -eq 0 -and
            $longdesc.Result[0]) {
            $longdesc.Result[0].Long_Description
        }
        else {
            Write-Error "Error: Method:App_Log_Long_Desc_Get: $($Result.ResultMessage)"
        }
    }
    else {
        Write-Error "Error: Method:GetLastAppLog: $($Result.ResultMessage)"
    }
}

Export-ModuleMember GetLastAppLog

Function GetDeviceSensorTypes ($config, $user, $devicetype) {
    $service = $config.DeviceMethod
    $method = "Device_Sensor_Types_Get"
    $URI = "$($config.WebHost)/$($service)$($method)"
    $JSON = $config.JSONSelectMethod
    $JSON.Ticket = $user.Ticket
    $JSON.LoginID = $user.Login_ID
    $JSON.Environment = $config.Environment
    $JSON.AppID = $config.AppID
    $JSON | add-member Device_Type_CD($devicetype.Device_Type_CD) -Force
    $JSON | Add-Member Is_Deleted($Null) -Force
    $JSON | Add-Member Sensor_Type_CD($Null) -Force

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }

    InvokeRest -params $Params
}

Export-ModuleMember GetDeviceSensorTypes


Function ConvertKVP ($data) {
    [System.Management.Automation.PSCustomObject]$Obj = $null
    ForEach ($line in $data) {
        $obj
    }
}

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

Function ExportWSToCSV ($excelFile, $csvLoc) {
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile.FullName)
    foreach ($ws in $wb.Worksheets) {
        $n = $ws.Name
        "Saving to: .\$((get-item $csvLoc).name)\$n.csv"
        $ws.SaveAs("$csvLoc\$n.csv", 23)
    }
    ""
    $E.Quit()
}

Export-ModuleMember ExportWSToCSV

#General Retry Function for API Calls made for DataManagement External Service
Function InvokeRest($params) {
    $attempt = 3
    $FailCount = 0
    $interval = 5000
    While ($attempt -gt 0 -and -not $success) {
        Try {
            $Result = Invoke-RestMethod @params
            if ($Result.ResultCode -eq 0 -and
                $Result.ResultMessage -eq $null) {
                $success = $true
            }
            else {
                $success = $false
                $FailCount++
                if ($FailCount -ge 3) {
                    Write-Host -ForegroundColor Green "Updating Login Ticket..."
                    $Body = $params.body | convertfrom-json
                    $Body.Ticket = (LoginUser -config $Config).Ticket
                    $Params.Body = $Body | ConvertTo-Json
                }
                Write-Host -ForegroundColor red "$(if($Result.ResultMessage){$Result.ResultMessage})"
                Write-Host -ForegroundColor Yellow "Attempts Remaining: $($attempt) | Retrying in $($interval)ms..."
                Start-Sleep -Milliseconds $interval
            }
        }
        Catch {
            $success = $false
            $FailCount++
            Write-Host -ForegroundColor red "$(if($Result.ResultMessage){$Result.ResultMessage})"
            Write-Host -ForegroundColor Yellow "Attempts Remaining: $($attempt) | Retrying in $($interval)ms..."
            Start-Sleep -Milliseconds $interval
            $Result = [pscustomobject]@{
                ResultCode    = 4
                ResultMessage = $_
            }
        }
        
        $attempt--
    }
    
    $Result
}

Export-ModuleMember InvokeRest

Function SetInstance($Config) {
    $Instance = $Config.Instances | Where-Object {
        $_.InstanceName -eq $Config.Instance
    }
    if ($Instance) {
        $Config.WebHost = $Instance.WebHost
        $Config.Environment = $Instance.Environment
        $Config.Username = $Instance.username
        $Config.Password = $Instance.password
        $Config.Description = $Instance.Description
        Write-Host -ForegroundColor Yellow "Setting Instance: $($Instance.Description)"
    }
    else {
        Write-Error "Error:  match InstanceName to Config.Instances[]"
    }
    $Config
}

Export-ModuleMember SetInstance

Function toNull ($value) {
    if ($Value -like 'null') {
        return $null
    }
    return $value
}

Export-ModuleMember toNull

Function LeadTrail ($Row){
    #Leading and Trailing Space Validator
    $ExtraSpaceRegex = "^(\s+)(\w+)|(\s+)$"
    $LeadColumns = @()
    $RowDef = $Row | Get-Member | Where-Object {
        $_.MemberType -eq 'NoteProperty'
    }
    $LeadTrail = $false
    For($i=0; $i -lt $RowDef.Length; $i++){
        $column = $RowDef[$i].Name
        if($Row.$column -match $ExtraSpaceRegex){
            $LeadTrail = $true
            $LeadColumns += $column
        }
    }
    [pscustomobject]@{
        ContainsLeading = $LeadTrail
        Columns = $LeadColumns
    }
}

Export-ModuleMember LeadTrail
