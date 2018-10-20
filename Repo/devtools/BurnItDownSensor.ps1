
$ErrorActionPreference = "stop"

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition
Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

Set-Location $RunDir

Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Set-Location $RunDir
Import-Module "..\Modules\Functions" -DisableNameChecking
Log "Loading Modules..."

Log "Loading Configuration..."
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -config $Config


Prereqs -config $Config

#Code above this line is mandatory



#Logs in
Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10) {
    Log "Retrieved Login Ticket"
}
else {
    Write-Error "Error Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

$PartitionString = Read-Host -Prompt "Enter Partition Search String"

$PartitionResult = PartitionSearch -config $Config -user $User -searchstring $PartitionString


$enum = 0
$menu = @{}
ForEach($Partition in $PartitionResult.Result){
    $enum++
    $menu.Add($enum, $Partition)
    Write-Host -ForegroundColor Cyan "$enum : Partition_ID: $($Partition.Partition_ID) | $($Partition.Path)"
}
""


$PartitionKey = Read-Host -Prompt "Select a Partition to search inside of"
$Partition = $menu[[int]$PartitionKey]

$SensorSearchStr = Read-Host -Prompt "Enter Sensor Search String"



$SensorSearchParams = @{
    Config = $Config
    User = $User
    Partition_ID = $Partition.Partition_ID
    searchstring = $SensorSearchStr
}
$SensorResults = SensorSearch @SensorSearchParams
""
Write-Host -ForegroundColor Yellow "Found: $($SensorResults.Result.Count) Sensors"
""
$SensorResults.Result | Format-Table

$Commit = Read-Host -Prompt "To Destroy These Sensors Type DESTROY (Case Sensitive)"

if($Commit -ceq "DESTROY"){
    $SensorIDs = $SensorResults.Result.Device_Sensor_ID
    Write-Host -ForegroundColor Red "DESTROYING SENSORS..."
}else{
    $SensorIDs = $Null
    Write-Host -ForegroundColor Magenta "Aborting..."
    RETURN
}


$TaskStatus = @{
    1 = "New"
    2 = "InProgress"
    3 = "Error"
    4 = "Completed"
    5 = "Cancelled"
}

ForEach ($sensor in $SensorIDs) {
    Log "Destroying Sensor: $sensor"
    $taskid = [guid]::NewGuid()
    $DestroySensorParams = $Null
    $DestroySensorParams = @{
        config = $Config
        user   = $User
        sensor = $sensor
        taskid = $taskid

    }
    Log "Creating Task: $($DestroySensorParams.taskid.guid)"
    $Destroy = DestroySensor @DestroySensorParams

    if ($Destroy.ResultCode -eq 0 -and
        $Destroy.ResultMessage -eq $Null) {
            
        $TaskParams = @{
            config = $Config
            user   = $User
            taskid = $taskid
        }
            
        $SetProg = TaskProgress @TaskParams
            
        $laststatus = 1
        do {
            $status = $Null
            Start-Sleep -Milliseconds 500
            $Task = $Null
            $Task = GetTask @TaskParams
            $status = $Task.Result[0].Status_CD
            if ($status -ne $laststatus) {
                Log "Task $($taskid.Guid) Status: $($TaskStatus[$status])"
            }
            $laststatus = $status
        }While ($status -ne 4 -and $status -ne 5)
        if ($status -eq 5) {
            Write-Error "Error: Processing Task - Please reviewe TaskProcessor Logs"
        }
        ""
    }
    else {
        Write-Error "Error: Method:DestroySensor Message:$($Destroy.ResultMessage)"
    }

}




