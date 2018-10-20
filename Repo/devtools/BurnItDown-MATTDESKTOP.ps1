

$ErrorActionPreference = "stop"
$Partition_ID = 14806

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition
Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

Set-Location $RunDir

Log "Loading Configuration..."
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -Config $Config

Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Set-Location $RunDir
Import-Module "..\Modules\Functions" -DisableNameChecking
Log "Loading Modules..."
Prereqs -config $Config

#Code above this line in mandatory

#Logs in
Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10) {
    Log "Retrieved Login Ticket"
}
else {
    Write-Error "Error Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

#Gets lists of Devices

$params = @{
    config    = $Config
    user      = $User
    partition = $Partition_ID
}
$Devices = $Null
Log "Querying Devices..."
$Devices = GetDevices @params
Log "Found: $($Devices.Result.Count) Devices"

$TaskStatus = @{
    1 = "New"
    2 = "InProgress"
    3 = "Error"
    4 = "Completed"
    5 = "Cancelled"
}

if ($Devices.Result) {
    ForEach ($Device in $Devices.Result) {
        Log "Destroying Device: $($Device.Device_Name) : $($Device.Device_ID)"
        $taskid = [guid]::NewGuid()
        $DestroyDeviceParams = $Null
        $DestroyDeviceParams = @{
            config = $Config
            user   = $User
            device = $Device
            taskid = $taskid

        }
        Log "Creating Task: $($DestroyDeviceParams.taskid.guid)"
        $Destroy = DestroyDevice @DestroyDeviceParams

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
            if($status -eq 5){
                Write-Error "Error: Processing Task - Please reviewe TaskProcessor Logs"
            }
            ""
        }
        else {
            Write-Error "Error: Method:DestroyDevice Message:$($Destroy.ResultMessage)"
        }

    }
}
else {Write-Error "No Devices Found"}




