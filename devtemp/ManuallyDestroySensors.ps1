
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

$SensorResults = @(
    '145CF29B-9396-E811-80D5-005056BF0489',
    '135CF29B-9396-E811-80D5-005056BF0489',
    '125CF29B-9396-E811-80D5-005056BF0489',
    '115CF29B-9396-E811-80D5-005056BF0489',
    '105CF29B-9396-E811-80D5-005056BF0489',
    '0F5CF29B-9396-E811-80D5-005056BF0489',
    '0E5CF29B-9396-E811-80D5-005056BF0489',
    '0D5CF29B-9396-E811-80D5-005056BF0489',
    '0C5CF29B-9396-E811-80D5-005056BF0489',
    '0B5CF29B-9396-E811-80D5-005056BF0489',
    '0A5CF29B-9396-E811-80D5-005056BF0489',
    '015CF29B-9396-E811-80D5-005056BF0489',
    'FD5BF29B-9396-E811-80D5-005056BF0489',
    'FC5BF29B-9396-E811-80D5-005056BF0489',
    'FB5BF29B-9396-E811-80D5-005056BF0489',
    'FA5BF29B-9396-E811-80D5-005056BF0489',
    'E790E495-9396-E811-80D5-005056BF0489',
    'E690E495-9396-E811-80D5-005056BF0489',
    'E490E495-9396-E811-80D5-005056BF0489',
    'E390E495-9396-E811-80D5-005056BF0489',
    'E290E495-9396-E811-80D5-005056BF0489',
    'DE90E495-9396-E811-80D5-005056BF0489',
    'D690E495-9396-E811-80D5-005056BF0489',
    'D190E495-9396-E811-80D5-005056BF0489',
    'D090E495-9396-E811-80D5-005056BF0489',
    'CF90E495-9396-E811-80D5-005056BF0489',
    'CE90E495-9396-E811-80D5-005056BF0489',
    'CA90E495-9396-E811-80D5-005056BF0489',
    'C990E495-9396-E811-80D5-005056BF0489',
    'C890E495-9396-E811-80D5-005056BF0489',
    'C390E495-9396-E811-80D5-005056BF0489',
    'C290E495-9396-E811-80D5-005056BF0489',
    'C190E495-9396-E811-80D5-005056BF0489',
    'C090E495-9396-E811-80D5-005056BF0489',
    'BC90E495-9396-E811-80D5-005056BF0489',
    'BB90E495-9396-E811-80D5-005056BF0489',
    'BA90E495-9396-E811-80D5-005056BF0489',
    'AA90E495-9396-E811-80D5-005056BF0489'
)

""
$SensorResults | Format-Table

$Commit = Read-Host -Prompt "To Destroy These Sensors Type DESTROY (Case Sensitive)"

if($Commit -ceq "DESTROY"){
    $SensorIDs = $SensorResults
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




