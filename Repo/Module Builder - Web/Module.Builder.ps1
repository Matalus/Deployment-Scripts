#Module Builder - Reads Smart Template format to ease sensor deployment
#Created by Matt Hamende 2018

$ErrorActionPreference = "stop"

#Auto Elevation
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

#Define Root Path
$RunDir = split-path -parent $MyInvocation.MyCommand.Definition

#Generic Logging function --
Function Log($message, $color) {
    if ($color) {
        Write-Host -ForegroundColor $color "$(Get-Date -Format u) | $message"
    }
    else {
        "$(Get-Date -Format u) | $message"
    }
}

#Set Current Root Dir --
Set-Location $RunDir

#Load function library from modules folder
Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "..\Modules\Functions" -DisableNameChecking

#Load config from JSON in project root dir
Log "Loading Configuration..."
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -config $Config

$instance_confirm = ("yes", "no") | Out-GridView -PassThru -Title "Is this the correct instance? $($Config.Instance) : $($Config.Description)" #Read-Host -Prompt "Is this the correct instance? y/n"

if ($instance_confirm -like "*n*") {
    Log "Select the correct instance" "Magenta"


    $instance_params = @(
        "InstanceName",
        "Description",
        "WebHost",
        "Environment",
        "Username",
        @{N = "Password"; E = {"****"}}
    )
    $select_instance = $Config.Instances | Select-Object $instance_params | Out-GridView -PassThru -Title "Select Instance"
    <# -- Legacy Host menu selection
    $instance_count = 1
    $menu = @{} 
    ForEach($Instance in $Config.Instances){
        $menu.Add($instance_count, $Instance)
        Write-Host -ForegroundColor Cyan "$instance_count : $($Instance.InstanceName) | $($Instance.Description) | $($Instance.WebHost)"
        $instance_count++
        ""
    }
    $key = Read-Host -Prompt "Enter Selection Number and Press Enter to continue"
    $Config.Instance = $menu[[int]$key].InstanceName
    #>
    $Config.Instance = $select_instance.InstanceName
    
    $Config = SetInstance -Config $Config
    Log "Instance changed to: $($Config.Description ) : $($Config.WebHost)" "Green"
}

#Load Modules in PSModules section of config.json
Log "Loading Modules..."
Prereqs -config $Config

#Delete temp files
Get-ChildItem $RunDir\Temp -Filter "*.csv" | Remove-Item -Force -ErrorAction SilentlyContinue

#Set location of template file
[array]$importList = Get-ChildItem $RunDir\Templates -Filter "*.xls*"
Clear-Host
Write-Host -ForegroundColor Gray "Module Builder -:- Created by Matt Hamende 2018" 
Write-Host -ForegroundColor Gray "##########################################################"
Write-Host -ForegroundColor Green "Connected to Instance: $($Config.Description)"
Write-Host -ForegroundColor Magenta "   +-- WebHost: $($Config.WebHost)"
Write-Host -ForegroundColor Magenta "   +-- Environment: $($Config.Environment)"
Write-Host -ForegroundColor Magenta "   +-- AppID: $($Config.AppID)"
Write-Host -ForegroundColor Gray "##########################################################"

""

if ($importList.Count -gt 1) {
    Write-Host -ForegroundColor Yellow "Select A Template to Build"
    <#
    $enum = 1
    $menu = @{} 
    ForEach ($file in $importList) {
        $menu.Add($enum, $file)
        ""
        Write-Host -ForegroundColor Cyan "$($enum): $($file.Name) : $($file.Length) : $($file.LastWriteTime)" 
        $enum++
    }
    ""
    $key = Read-Host -Prompt "Enter Selection Number and Press Enter to continue"
    $importList = $menu[[int]$key]
    #>
    $importList = $importList | Out-GridView -PassThru -Title "Select Template"
}

#Kill Orphaned Excel Procs
[array]$Procs = Get-WmiObject Win32_Process -Filter "name = 'Excel.exe'"
$embeddedXL = $Procs | ? {$_.CommandLine -like "*/automation -Embedding*"}
$embeddedXL | ForEach-Object { $_.Terminate() | Out-Null }

#Load CSV from template file
Log "Extracting Template Data..."
ForEach ($xl in $importList) {
    Write-Host -ForegroundColor Magenta "Importing Data from Excel : $($xl.name)..."
    ExportWSToCSV -excelFile $xl -csvLoc $RunDir\Temp
}

#Sensor Integration Template
[Array]$FullTemplate = Import-Csv $RunDir\Temp\Template.csv -Encoding UTF8

#[Array]$FullTemplate = Import-Excel -Path $importList.FullName -WorksheetName "Template" -DataOnly

[Array]$Template = $FullTemplate | Where-Object {
    $_.Update -eq 1
}

#Login to Visualizer
Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10) {
    Log "Retrieved Login Ticket" "Green"
}
else {
    GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:LoginUser: Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

Log "Getting Release Info..."
$Release = GetAppReleases -config $Config -user $User
$Config | Add-Member Release($Release.Result) -Force


#Retrieve Partition Types Array
Log "Getting Partition Types Data..."
$params = @{
    config  = $Config
    user    = $User
    Service = $Config.MiscMethod
    Method  = "SELECT_919"
}

$All_Partition_Types = $Null
$All_Partition_Types = GetSelectMethod @params
#Retrieve Device Types Array
Log "Getting Device Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "SELECT_1127"
}
$All_Device_Types = $Null
$All_Device_Types = GetSelectMethod @params

#Retrieve Measurement Types Array
Log "Getting Measurement Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.MiscMethod
    method  = "SELECT_897"
}
$Measurement_Types = GetSelectMethod @params

#Retrieve Sensor Type Array
<# -- Changed to Cache on Demand and only load sensor type data relevant to device types being built
   -- Gets called just before the sensor type is matched
Log "Getting Sensor Types Data..."
$params = @{
    config  = $config 
    user    = $User 
    service = $config.DeviceMethod
    method  = "Device_Sensor_Types_Get"
}

#Ordered list of properties to select for template
#Inline Expressions are formatted to allow results to be copy / pasted to the template tab
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

if (-not $Sensor_Types -or $Error) {
    #Selects based on properties list and sorts objects descending by type cd
    $Sensor_Types = (GetSelectMethod @params).Result |
        Select-Object -Property $SelectProperties | Sort-Object Sensor_Type_CD -Descending
}
$Error.Clear()
#>
[array]$Sensor_Types = $null

#Retrieve Monitor Type Array
Log "Getting Monitor Types Data..."
$params = @{
    config  = $Config
    user    = $User
    service = $Config.DeviceMethod
    method  = "SELECT_1049"
}

$Monitor_Types = GetSelectMethod @params

#Retrieving Reading Type Array
Log "Getting Reading Types Data..."
$params = @{
    config  = $Config
    user    = $User
    service = $Config.DeviceMethod
    method  = "SELECT_1048"
}

$Reading_Types = GetSelectMethod @params

#Retrieving Conversio Formulas Array
Log "Getting Conversion Formulas"
$params = @{
    config  = $Config
    user    = $User
    service = $Config.MiscMethod
    method  = "SELECT_1122"
}

$Conversion_Formulas = GetSelectMethod @params
""

Log "Found: $($FullTemplate.Count) Template Rows"

Log "Validating Template..." "Cyan"
$ValidationError = 0
$LeadRowCount = 1 #Starts at 1 due to header row
ForEach ($Row in $FullTemplate) {
    $LeadRowCount++
    $LeadResult = LeadTrail $Row 
    if ($LeadResult.ContainsLeading -eq $true) {
        $ValidationError++
        ForEach ($column in $LeadResult.Columns) {
            Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : The field "$column" contains extra trailing or leading spaces that may cause mismatches: [$($row.$column)]"
        }
    }

    if ($Row.Device_Type_Short.Length -gt 15) {
        #Checks Device_Type_Short length
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Sensor_Type_Short: [$($Row.Device_Type_Short)]($($Row.Device_Type_Short.Length)) : exceeds maximum chars(15)"
        $ValidationError++
    }

    if ($Row.Sensor_Type_Short.Length -gt 15) {
        #Checks Sensor_Type_Short length
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Device_Type_Short: [$($Row.Sensor_Type_Short)]($($Row.Sensor_Type_Short.Length)) : exceeds maximum chars(15)"
        $ValidationError++
    }


    if ($Row.Measurement_Type -notin $Measurement_Types.Result.Description) {
        #Checks for invalid measurement types
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Measurement Type: [$($Row.Measurement_Type)] : does not exist in this instance of RunSmart"
        $ValidationError++
    }

    if ($Row.Monitor_Type -notin $Monitor_Types.Result.Description) {
        #Checks for invalid monitor types
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Monitor Type: [$($Row.Monitor_Type)] is not a valid monitor type"
        $ValidationError++
    }

    if ($Row.Reading_Type -notin $Reading_Types.Result.Description) {
        #Checks for invalid reading Types
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Reading Type: [$($Row.Reading_Type)] is not a valid reading type"
        $ValidationError++
    }

    if ($Row.Convert -ne "NULL" -and $row.Convert -notin $Conversion_Formulas.Result.Conversion_Formula_CD) {
        #Checks for invalid conversion codes
        Write-Host -ForegroundColor Yellow "Row $($LeadRowCount) : Conversion Formula CD: [$($Row.Convert)] is not a valid conversion formula cd"
        $ValidationError++
    }


}

if ($ValidationError -ge 1) {
    Write-Error "Error: Found $ValidationError Rows with Errors please resolve and try again"
}
else {
    Log "Template is Valid" "Green"
}

$Script:AltRootID = $Null

#Loop through items in template
$Count = 0
ForEach ($Row in $Template) {
    
    #Set Alternate Partition ID for loop variable already in memory
    if ($Script:AltRootID -ne $Null) {
        $Row.Partition_ID = $Script:AltRootID
        $Row.Sensor_Type_CD = $Null
        $Row.Device_Type_CD = $Null
        $Row.Device_Sensor_ID = $Null
        $Row.Device_ID = $Null
    }
     
    #Where-Object query on All_Device_Types to match Device_Type_CD
    $Count++
    ""
    if (-not ($Row.Partition_ID -match "\d+")) {
        GetLastAppLog -config $Config -user $User; Write-Error "Error: You must enter a valid Partition_ID in the template"
    }

    Write-Host -ForegroundColor cyan "$(Get-Date -Format u) | $Count. Partition:$($Row.Partition_ID) | Path:$($row.Parent_Path) | DeviceType:$($Row.Device_Type) | DeviceName:$($Row.Device_Name) | Sensor:$($Row.Sensor_Type)"
    
    #Fetch Data for Top Level Partition
    $GetPartitionParams = @{
        user      = $User
        config    = $Config
        partition = $Row.Partition_ID
    }

    $Partition = $Null
    $Partition = GetPartition @GetPartitionParams

    #Checks for bad return code and errors
    if ($Partition.ResultCode -eq 0 -and
        $Partition.ResultMessage -eq $Null -and
        $Partition.Result -ne $Null) {
        #Log "--> Found: Partition: $($Partition.Result.Partition_Name) : $($Partition.Result.Partition_ID)"
    }
    else {
        if ($Partition.Result -eq $Null) {
            Log "Unable to locate root partition ID: $($Row.Partition_ID)" "Yellow"
            #$SearchConfirm = Read-Host -Prompt "Search for Partition in current instance? y/n"
            [array]$SearchConfirm = ("Yes", "No") | Out-GridView -PassThru -Title "Search for Partition in current instance? y/n"
            if ($SearchConfirm -like "*y*" -and $SearchConfirm.Count -lt 2) {
                #$PartitionString = Read-Host -Prompt "Enter Partition Search String"

                if ([appdomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like "*VisualBasic*"} -eq $null) {
                    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") #Load VB Assemblies
                }

                if ($Last_SearchString) {
                    $PartitionString = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Partition Search String", "Partition Search", $Last_SearchString)
                    $Last_SearchString = $PartitionString
                }
                else {
                    $PartitionString = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Partition Search String", "Partition Search")
                    $Last_SearchString = $PartitionString
                }
                $PartitionResult = $null
                $retry = 0
                While ($PartitionResult.Result.Count -lt 1) {
                    
                    if ($retry -ge 1) {
                        $PartitionString = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Partition Search String", "Partition Search: No Results Please Try a different search term", $Last_SearchString)
                        $Last_SearchString = $PartitionString
                    }
                    $PartitionResult = PartitionSearch -config $Config -user $User -searchstring $PartitionString
                    $retry++
                }
                <# -- Legacy Host Select method
                $enum = 0
                $partmenu = @{}
                ForEach($Partition in $PartitionResult.Result){
                    $enum++
                    $partmenu.Add($enum, $Partition)
                    Write-Host -ForegroundColor Cyan "$enum : Partition_ID: $($Partition.Partition_ID) | $($Partition.Path)"
                }
                #>
                ""
                While ($PartitionResult.Result.Count -gt 0) {
                    if ($PartitionResult.Result.Count -gt 0) {
                        $PartitionResult.Result = $PartitionResult.Result | Out-GridView -PassThru -Title "Select a Partition to set Root Partition"
                        #Retrieves info for selected Partition
                        $GetPartitionParams = @{
                            user      = $User
                            config    = $Config
                            partition = $PartitionResult.Result.Partition_ID
                        }
                    
                        $Partition = $Null
                        $Partition = GetPartition @GetPartitionParams
                    }
                    else {
                        Write-Error "Partition Search Returned no results"
                    }
                    if ($PartitionResult.Result.Count -gt 1) {
                        Write-Host -ForegroundColor "You Must Select Only 1 Partition"
                    }
                }   

                
            


                <# -- Legacy Host Select method
                $PartitionKey = Read-Host -Prompt "Select a Partition to set Root Partition"
                if($PartitionKey.Length -lt 1){
                    Write-Error "Check the hierarchy to make sure this partition exists in the current instance: $($Config.Description) $($Config.WebHost)"

                }
                #>

                #$Partition = $partmenu[[int]$PartitionKey] # -- Legacy Host Select method
                $Script:AltRootID = $Partition.Result.Partition_ID
                Log "Updating and Sanitizing Template..." "Magenta"
                $FullTemplate | ForEach-Object { 
                    $_.Partition_ID = $Partition.Result.Partition_ID;
                    $_.Sensor_Type_CD = $Null;
                    $_.Device_Type_CD = $Null;
                    $_.Device_Sensor_ID = $Null;
                    $_.Device_ID = $Null;
                } 
                
                $Row.Partition_ID = $Partition.Result.Partition_ID
                $Row.Sensor_Type_CD = $Null
                $Row.Device_Type_CD = $Null
                $Row.Device_Sensor_ID = $Null
                $Row.Device_ID = $Null
                #Sanitizes Template for new environment
                
                [Array]$Template = $FullTemplate | Where-Object {
                    $_.Update -eq 1
                } #Recreate update array
                
            }
            else {
                GetLastAppLog -config $Config -user $User; Write-Error "Error:GetPartition: Method: ResultCode: $($Partition.ResultCode) ResultMessage: $($Partition.ResultMessage)"
            }
        }
    }
    if ($Row.Parent_Path.Length -ge 1) {
        #Attempt to locate child partitions
        [array]$TempPartitions = $Row.Parent_Path.Split("\")
        [array]$TempTypes = $Row.Partition_Types.Split("\")

        $LastRow = $Partition.Result
        For ($i = 0; $i -lt $TempPartitions.Length; $i++) {
            $ChildParams = @{
                user      = $User
                config    = $Config
                partition = $LastRow
            }
            $children = GetPartitionChildren @ChildParams
            <##
        if([version]$Config.Release.Version_Number -lt [version]"3.7.0.0"){
            $children = (GetSelectMethod -config $Config -user $User -service $Config.PartitionMethod -method "SELECT_1200")
            $trimChildren = $children.Result | Where-Object{
                $_.Parent_Partition_ID -eq $LastRow.Partition_ID
            }
            $children.Result = $trimChildren
            
        }else{
            $children = GetPartitionChildren @ChildParams
        }
        #>
            if ($children.Resultcode -eq 0 -and
                $children.ResultMessage -eq $null) {
                Try {    
                    $childType = $All_Partition_Types.Result.PartitionType | Where-Object { $_.Description -eq $TempTypes[$i] }
                }
                Catch {Write-Error "Unable to Match Partition Type: $($TempTypes[$i])"}
                $childMatch = $Null
                $childMatch = $children.Result | Where-Object {
                    $_.Partition_Name -eq $TempPartitions[$i] -and
                    $_.Partition_Type_CD -eq $childType.Partition_Type_CD
                }

                if ($childMatch) {
                    Log "--> Found: Child Partition: $($childMatch.Partition_Name) : $($childMatch.Partition_ID)"
                    $LastRow = $childMatch

                }
                else {
                    Log "--> Unable to Locate Child Partition $($TempPartitions[$i]) ($($childtype.Description)) - Inserting" "Cyan"
                
                    $partitionData = [pscustomobject]@{
                        Location_ID          = $LastRow.Location_ID 
                        Partition_ID         = $Null
                        Parent_Partition_ID  = $LastRow.Partition_ID
                        Partition_Type_CD    = $childType.Partition_Type_CD
                        Partition_Name       = $TempPartitions[$i]
                        Partition_Short_Name = $TempPartitions[$i]
                    }
                
                    $PartitionSetParams = @{
                        user   = $User
                        config = $Config
                        data   = $partitionData
                    }
                    $SetPartition = SetPartition @PartitionSetParams
                    if ($SetPartition.ResultCode -eq 0 -and
                        $SetPartition.ResultMessage -eq $Null) {
                        Log "--> Success! - Inserted Partition: $($SetPartition.Result.Partition_ID)" "Green"
                        $LastRow = $SetPartition.Result
                    }
                    else {
                        Write-Error "Error: SetPartition ResultCode: $($SetPartition.ResultCode) ResultMessage: $($SetPartition.ResultMessage)"
                    }
                }
            }
            else {
                GetLastAppLog -config $Config -user $User;
                Write-Error "Error:GetPartitionChildren: ResultCode: $($child.ResultCode) ResultMessage: $($child.ResultMessage)"
            }
    
            
        }
    }
    $DeviceType = $Null
    [array]$DeviceType = $All_Device_Types.Result | Where-Object {
        $_.Device_Type_Name -eq $Row.Device_Type
    }

    if ($Row.Parent_Path.Length -lt 1 -and ($Row.Partition_Types -eq "NULL" -or $Row.Partition_Types -eq $Null -or $Row.Partition_Types.Length -lt 1)) {
        if ($Partition.Result -ne $Null) {
            $LastRow = $Partition.Result
        }
        else {
            $LastRow = $Partition
        }
    }

    #if Device Type not present, inserts generic "Environment" device type based on template
    if ($DeviceType -eq $Null) {
        #            GetLastAppLog -config $Config -user $User; Write-Error "Error: $($Row.Device_Type) - doesn't exist"
        Log "--> Device Type: $($Row.Device_Type) - doesn't exist, Inserting..." "Cyan"

        $AddDevTypeParams = @{
            config = $Config
            user   = $User
            data   = $Row
        }

        #Executes function that calls webservice method to insert device type
        $AddDevType = $Null
        $AddDevType = AddDeviceType @AddDevTypeParams

        #Checks for bad return code and errors
        if ($AddDevType.ResultCode -eq 0 -and
            $AddDevType.ResultMessage -eq $Null) {
            Log "--> Success! Added Device Type: $($AddDevType.Result.Device_Type_Name) : $($AddDevType.Result.Device_Type_CD)" "Green"
            #Sets Device Type variable for current row
            $DeviceType = $AddDevType.Result
            #Update Device Type CD in Template
            $FullTemplate[$($FullTemplate.indexof($row))].Device_Type_CD = $DeviceType.Device_Type_CD
            Log "Refreshing Device Types..." "Yellow"
            $params = @{
                config  = $config 
                user    = $User 
                service = $config.MiscMethod
                method  = "SELECT_1127"
            }
            $All_Device_Types = $Null
            $All_Device_Types = GetSelectMethod @params

        }
        else {
            GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:AddDeviceType: ResultCode: $($AddDevType.ResultCode) ResultMessage: $($AddDevType.ResultMessage)"
        }


    }
    else {
        Log "--> Found: Device Type: $($DeviceType.Device_Type_Name) : $($DeviceType.Device_Type_CD)"
        #Update Device Type CD in Template
        $FullTemplate[$($FullTemplate.indexof($row))].Device_Type_CD = $DeviceType.Device_Type_CD
    }
    #Sets Device_Type_CD for Template Updates
    if ($Row.Device_Type_CD.Length -lt 1) {
        $Row.Device_Type_CD = $DeviceType.Device_Type_CD
    }
    
    #WebService Query Devices in Partition
    $params = @{
        config    = $Config
        user      = $User
        partition = $LastRow.Partition_ID
    }
    $Devices = $Null
    $Devices = GetDevices @params

    #Determine if query was successful and match current device
    if ($Devices.ResultCode -eq 0 -and
        $Devices.ResultMessage -eq $Null) {
        $DeviceMatch = $Null
        [array]$DeviceMatch = $Devices.Result | Where-Object {
            $_.Device_Name -eq $Row.Device_Name -and
            $_.Device_Type_CD -eq $DeviceType.Device_Type_CD
        }
    }
    else {
        GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:GetDevices Message:$($Devices.ResultMessage)"
    }

    #determines if device exists or needs to be inserted
    if ($DeviceMatch) {
        $DeviceDetailParams = @{
            config   = $Config
            user     = $User
            deviceID = $DeviceMatch.Device_ID
        }
        $DeviceDetails = $Null
        $DeviceDetails = GetDeviceProperties @DeviceDetailParams
        Log "--> Found: Device: $($DeviceMatch.Device_Name) : $($DeviceMatch.Device_ID)"
        #Update Sensor Type CD on Template
        $FullTemplate[$($FullTemplate.indexof($row))].Device_ID = $DeviceDetails.Result.Device_ID
    }
    else {
        #            GetLastAppLog -config $Config -user $User; Write-Error "Error: Device Match is Null"
        if ($Row.Device_Name -in $Devices.Result.Device_Name) {
            Write-Error "Error: This Partition contains a device with the same name but a different partition type - please validate partition types string"
        }
        Log "--> Unable to locate Device: $($Row.Device_Name) - Inserting..." "Cyan"

        #Append Device Type CD
        $Row | Add-Member Device_Type_CD($DeviceType.Device_Type_CD) -Force

        $InsertDeviceParams = @{
            data      = $Row
            config    = $Config
            user      = $User
            partition = $LastRow
        }

        $InsertDevice = $Null
        $InsertDevice = InsertDevice @InsertDeviceParams

        if ($InsertDevice.ResultCode -eq 0 -and
            $InsertDevice.ResultMessage -eq $Null) {
            Log "--> Success! Added Device: $($Row.Device_Name) : $($InsertDevice.Result.Device_ID)" "Green"

            $DeviceDetailParams = @{
                config   = $Config
                user     = $User
                deviceID = $InsertDevice.Result.Device_ID
            }
            $DeviceDetails = $Null
            $DeviceDetails = GetDeviceProperties @DeviceDetailParams
    
        }
        else {
            GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:GetDeviceProperties: ResultCode: $($InsertDevice.ResultCode) ResultMessage: $($InsertDevice.ResultMessage)"
        }
    }
    #Determines if sensortype exists

    #Loads Only Sensor Types related to current Device Type
    $GetDeviceSensorTypesParams = @{
        user       = $User
        config     = $Config
        devicetype = $DeviceDetails.Result
    }
    $DeviceSensorTypes = GetDeviceSensorTypes @GetDeviceSensorTypesParams
    if ($DeviceSensorTypes.ResultCode -eq 0 -and
        $DeviceSensorTypes.ResultMessage -eq $Null) {
        ForEach ($type in $DeviceSensorTypes.Result) {
            if ($type.Sensor_Type_CD -notin $Sensor_Types.Sensor_Type_CD) {
                #add type to array
                $Sensor_Types += $type
            }
        }
    }
    else {
        Write-Error "Error: Method:GetDeviceSensorTypes ResultCode:$($DeviceSensorTypes.ResultCode) ResultMessage:$($DeviceSensorTypes.ResultMessage)"
    }

    $SensorType = $Null

    if ($Row.Sensor_Type_CD.Length -gt 0) {
        $SensorType = $Sensor_Types | Where-Object {
            $_.Sensor_Type_CD -eq $Row.Sensor_Type_CD -and
            $_.Device_Type_CD -eq $Row.Device_Type_CD
        }
        if ($SensorType -eq $Null -and [int]$Row.Sensor_Type_CD -lt 100000) {
            Log "Unable to locate factory Sensor Type: $($Row.Sensor_Type_CD) Searching for Custom Type with same name" "Yellow"
            $SensorType = $Sensor_Types | Where-Object {
                $_.Sensor_Name -eq $Row.Sensor_Type -and
                $_.Device_Type_CD -eq $DeviceType.Device_Type_CD
            }
        }
    }
    else {
        $SensorType = $Sensor_Types | Where-Object {
            $_.Sensor_Name -eq $Row.Sensor_Type -and
            $_.Device_Type_CD -eq $DeviceType.Device_Type_CD
        }
    }
    if ($SensorType) {
        Log "--> Found: SensorType:" 
        Write-host -ForegroundColor Magenta "                            +-- Sensor_Type_CD: $($SensorType.Sensor_Type_CD)"
        Write-host -ForegroundColor Magenta "                            +-- Name: $($SensorType.Sensor_Name)"
        Write-host -ForegroundColor Magenta "                            +-- Last_Updated_Date: $($SensorType.Last_Update_Date)"
        #Update Sensor Type CD on Template
    
        if ($Row.Sensor_Type_CD.Length -lt 1 -or $Row.Sensor_Type_CD -eq $null) {
            $FullTemplate[$($FullTemplate.indexof($row))].Sensor_Type_CD = $SensorType.Sensor_Type_CD
        }

        $CompareSTParams = @{
            template   = $Row
            sensortype = $SensorType
            user       = $User
            config     = $Config
        }
    
        $CompareSensorType = CompareSensorType @CompareSTParams

        #Purge cached Sensor Type
        $Sensor_Types = $Sensor_Types | Where-Object {
            $_.Sensor_Type_CD -eq $SensorType.Sensor_Type_CD
        }

        if ($CompareSensorType.Match -eq $false) {
            Log "Updating Sensor Type: $($SensorType.Sensor_Name) : $($SensorType.Sensor_Type_CD)" "Yellow"
            Log "$($CompareSensorType.attributes -join ", ")" "Yellow"

            $SetSensorTypeParams = @{
                data   = $Row
                config = $Config
                user   = $User
            }
            if ([int]$SensorType.Sensor_Type_CD -lt 100000 -and $Config.Instance -eq "BL-FA") {
                Log "This is a Factory Sensor Type : do you want to update this"
                #$updateST = Read-Host -Prompt "Update? y / n"
                $updateST = ("Yes", "No") | Out-GridView -PassThru -Title "$($SensorType.Sensor_Type_CD) : $($SensorType.Sensor_Name): is a Factory Sensor Type : do you want to update this?"
            }
            else {
                $updateST = 'y'
            }
        
            if ($updateST -like "*y*") {
                $SetSensorType = SetSensorType @SetSensorTypeParams
            }
        }
    }
    else {
        Log "--> Unable to locate Sensor Type : $($Row.Sensor_Type) - Inserting..." "Cyan"

        #Append device type data
        $row | Add-Member Device_Type_CD($DeviceType.Device_Type_CD) -Force
        $row | Add-Member Sensor_Type_CD("NULL") -Force
        $row | Add-Member Sensor_Name($row.Sensor_Type) -Force
        $row | Add-Member Sensor_Short_Name($Row.Sensor_Type_Short) -Force

        $SetSensorTypeParams = @{
            data   = $Row
            config = $Config
            user   = $User
        }
         
        $SetSensorType = SetSensorType @SetSensorTypeParams


        if ($SetSensorType.ResultCode -eq 0 -and
            $SetSensorType.ResultMessage -eq $Null) {
            Log "--> Success! Added Sensor Type: $($SetSensorType.Result.Sensor_Name) : $($SetSensorType.Result.Sensor_Type_CD)" "Green"
            $SensorType = $SetSensorType.Result
            Log "Refreshing Sensor Types..." "Yellow"
            $Sensor_Types += $SensorType

            #Update Sensor Type CD on Template
            $FullTemplate[$($FullTemplate.indexof($row))].Sensor_Type_CD = $SensorType.Sensor_Type_CD
            
            #Selects based on properties list and sorts objects descending by type cd
            <#
            $params = @{
                config  = $config 
                user    = $User 
                service = $config.DeviceMethod
                method  = "Device_Sensor_Types_Get"
            }
            
            $Sensor_Types = (GetSelectMethod @params).Result |
                Select-Object -Property $SelectProperties | Sort-Object Sensor_Type_CD -Descending
            #>


        }
        else {
            GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:SetSensorType: ResultCode: $($SetSensorType.ResultCode) ResultMessage: $($SetSensorType.ResultMessage)"
        }    
    }

    $GetSensorParams = @{
        config   = $Config
        user     = $User
        deviceID = $DeviceDetails.Result.Device_ID
    }
    $DeviceSensors = $Null
    $DeviceSensors = GetDeviceSensors @GetSensorParams

    if ($DeviceSensors.ResultCode -eq 0 -and
        $DeviceSensors.ResultMessage -eq $Null) {
        $Sensor = $Null
        $Sensor = $DeviceSensors.Result | Where-Object {
            $_.Sensor_Type_CD -eq $SensorType.Sensor_Type_CD
        }
        if ($Sensor) {
            Log "--> Found: Sensor:" 
            Write-host -ForegroundColor Magenta "                            +-- Device_Sensor_ID: $($Sensor.Device_Sensor_ID)"
            Write-host -ForegroundColor Magenta "                            +-- URI:  $($Sensor.URI)"
            #Update Sensor ID on Template
            $FullTemplate[$($FullTemplate.indexof($row))].Device_Sensor_ID = $Sensor.Device_Sensor_ID

            #TODO Compare and Update Sensor
            $CompareSensorParams = @{
                config   = $Config
                user     = $User
                sensor   = $Sensor
                template = $Row

            }
            $compare = CompareSensor @CompareSensorParams
            if ($compare.match -eq $false) {
                $SensorUpdateParams = @{
                    data       = $Row
                    config     = $Config
                    user       = $User
                    device     = $DeviceDetails.Result
                    sensortype = $SensorType
                }
                Log "Updating Sensor - $($compare.attributes -join ", ")" "Cyan"
                $SensorUpdate = UpdateSensor @SensorUpdateParams
                if ($SensorUpdate.ResultCode -eq 0 -and
                    $SensorUpdate.ResultMessage -eq $Null) {
                    #Do Nothing
                }
                else {
                    Write-Error "Error: Method:UpdateSensor ResultCode: $($SensorUpdate.ResultCode) ResultMessage: $($SensorUpdate.ResultMessage)"
                }
            }
        }
        else {
            #            GetLastAppLog -config $Config -user $User; Write-Error "Error: Unable to locate sensor: $($SensorType.Sensor_Name)"
            Log "--> Sensor: $($SensorType.Sensor_Name) : not found - Inserting... " "Cyan"

            if ($SensorType -and $Row.Monitor_Item_ID.Length -ge 4) {

                #Insert Sensor Params
                $InsertSensorParams = @{
                    data       = $Row
                    config     = $Config
                    user       = $User
                    device     = $DeviceDetails.Result
                    sensortype = $SensorType
                }
                $InsertSensor = $Null
                $InsertSensor = InsertSensor @InsertSensorParams

                #Checks for bad return code and errors
                if ($InsertSensor.ResultCode -eq 0 -and
                    $InsertSensor.ResultMessage -eq $Null) {
                    Log "--> Success! Added Sensor: $($SensorType.Sensor_Name) : $($InsertSensor.Result.Device_Sensor_ID)" "Green"
                    $Sensor = $InsertSensor.Result
                    #Update Sensor ID on Template
                    $FullTemplate[$($FullTemplate.indexof($row))].Device_Sensor_ID = $Sensor.Device_Sensor_ID

                }
                else {
                    GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:InsertSensor: ResultCode: $($InsertSensor.ResultCode) ResultMessage: $($InsertSensor.ResultMessage)"
                }    
                
            }
            else {
                if ($Row.Monitor_Item_ID.Length -eq 0) {
                    GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:InsertSensor: Monitor_Item_ID: Length cannot be zero"
                }
                else {
                    GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:InsertSensor: Missing Sensor Type Config Data"
                }
            }
        }

        #Get Thresholds from template
        $TempThresh = $Null
        [array]$TempThresh = TempThresh -data $row -sensor $Sensor

        #Get Thresholds from RunSmart
        $GetThreshParams = @{
            data   = $Sensor
            user   = $User
            config = $Config
        }
        $SensorThresh = $Null
        $SensorThresh = GetThresh @GetThreshParams

        if ($SensorThresh.Result) {
            
            #If Thresholds exist that shouldn't delete
            if ($SensorThresh -and $Row.ThreshIgnore -ne $true) {
                ForEach ($thresh in $SensorThresh.Result) {
                    if ($thresh.Sort_Order -notin $TempThresh.Sort_Order) {
                        Log "--> Threshold: $($thresh.Threshold_ID) doesn't match - Deleting..." "Yellow"
                        $threshDelete = DeleteThresh -id $thresh
                        if ($threshDelete.ResultCode -eq 0 -and
                            $threshDelete.ResultMessage -eq $Null) {
                            #do nothing
                        }
                        else {
                            GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:DeleteThresh: ResultCode:$($DeleteThresh.ResultCode) ResultMessage:$($DeleteThresh.ResultMessage)"
                        }
                    }
                }
            }
        }

        if ($Row.ThreshIgnore -eq $true) {
            Log "Threshold Ignore is enabled - configure thresholds beyond Sort_Order:5 for this sensor in RunSmart" "Yellow"
        }
    
        if ($TempThresh.Count -ge 1) {    
            ForEach ($thresh in $TempThresh) {
                $thresh | Add-Member Device_Sensor_ID($Sensor.Device_Sensor_ID) -Force
                $match = $true
                $compare = $Null
                $compare = $SensorThresh.Result | Where-Object {
                    $_.Sort_Order -eq $thresh.Sort_Order
                }

                #sanitize data
                if ($thresh.Min_Point -like "*inf*") {
                    $thresh.Min_Point = $Null
                }

                if ($thresh.Max_Point -like "*inf*") {
                    $thresh.Max_Point = $Null
                }

                if ($compare) {
                    $threshstring = "Sort $($compare.Sort_Order) " +
                    "| $(if($compare.Priority_CD -eq 0){"OK"}else{"P$($compare.Priority_CD)"}) " +
                    "| [$(if($compare.Min_Point -eq $null){"-inf"}else{$compare.Min_Point})." +
                    ".$(if($compare.Max_Point -eq $null){"inf"}else{$compare.Max_Point})]"
                    Log "--> Found: Threshold: $threshstring" "Blue"
                    #Begin compare block
                    if ($compare.Priority_CD -ne $thresh.Priority_CD) {
                        $match = $false
                    }

                    if ($compare.Min_Point -ne $thresh.Min_Point) {
                        $match = $false
                    }

                    if ($compare.Max_Point -ne $thresh.Max_Point) {
                        $match = $false
                    }
                }
                else {
                    $match = $false
                }

                if ($match -eq $false -and $compare) {
                    Log "--> Threshold: $($compare.Threshold_ID) doesn't match - Deleting..." "Yellow"
                    $threshDelete = DeleteThresh -id $compare
                    if ($threshDelete.ResultCode -eq 0 -and
                        $threshDelete.ResultMessage -eq $Null) {
                        #do nothing
                    }
                    else {
                        GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:DeleteThresh: ResultCode:$($DeleteThresh.ResultCode) ResultMessage:$($DeleteThresh.ResultMessage)"
                    }
                }

                if ($match -eq $false) {
                    $ThreshStringInsert = "Sort $($thresh.Sort_Order) " +
                    "| $(if($thresh.Priority_CD -eq 0){"OK"}else{"P$($thresh.Priority_CD)"}) " +
                    "| [$(if($thresh.Min_Point -eq $null){"-inf"}else{$thresh.Min_Point})." +
                    ".$(if($thresh.Max_Point -eq $null){"inf"}else{$thresh.Max_Point})]"
                    Log "--> Inserting New Threshold: $($ThreshStringInsert)" "Green"

                    $NewThreshParams = @{
                        data   = $thresh
                        config = $Config
                        user   = $User
                    }

                    $SetThresh = SetThresh @NewThreshParams
                    if ($SetThresh.ResultCode -eq 0 -and
                        $SetThresh.ResultMessage -eq $null) {
                        #do nothing
                    }
                    else {
                        GetLastAppLog -config $Config -user $User; Write-Error "Error: Method:SetThresh ResultCode:$($SetThresh.Resultcode) ResultMessage:$($SetThresh.ResultMessage)"
                    }
                }
            }
        }
    }
    else {
        Write-Host -ForegroundColor Yellow "                               +-- Thresholds: null"
    }
    "----------------------------------------------------------------------"
    $Row.Update = 0
}

Log "Updating Template..." "Yellow"

$ExportParams = @{
    Path          = $importList.FullName
    WorkSheetName = "Template"
    AutoSize      = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableStyle    = "Medium21"
    TableName     = "Template"
    ClearSheet    = $true
    PassThru      = $true
}

$XL = $FullTemplate | Sort-Object Parent_Path, Device_Name, Sensor_Type | Export-Excel @ExportParams
$XL.Workbook.Worksheets.MoveToStart("Template")
<# -- Macro Code that automatically sets the "update" column, IF I CAN GET THE DAMN THING WORKING
if ($XL.Workbook.VbaProject -eq $Null){
    $XL.Workbook.CreateVBAProject()
}
if ($XL.Workbook.VbaProject.Modules["AutoUpdate"] -eq $Null){
    $Macro = $XL.Workbook.VbaProject.Modules.AddModule("AutoUpdate")
    $Macro.Code = 'Private Sub Worksheet_Change(ByVal Target As Range)\r\n  If Target.Row > 1 Then Cells(Target.Row, "A") = 1\r\nEnd Sub\r\n'
}
#$XL.Workbook.Worksheets["Template"].CodeModule.Name = "AutoUpdate"
#>
$XL.Save()
$XL.Dispose()


"" ; Log "Done!" "Green"
