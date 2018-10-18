#Define Root Path
$RunDir = split-path -parent $MyInvocation.MyCommand.Definition
Function Log($message, $color) {
    if ($color) {
        Write-Host -ForegroundColor $color "$(Get-Date -Format u) | $message"
    }
    else {
        "$(Get-Date -Format u) | $message"
    }
}


#Set Current Root Dir
Set-Location $RunDir

#Load function library from modules folder
Log "Loading Function Library"
Remove-Module functions -ErrorAction SilentlyContinue
Import-Module "..\Modules\Functions" -DisableNameChecking -Force

#Load config from JSON in project root dir
Log "Loading Configuration..."
$Config = (Get-Content "..\config.json") -join "`n" | ConvertFrom-Json
$Config = SetInstance -config $Config

#Load Modules in PSModules section of config.json
Log "Loading Modules..."
Prereqs -config $Config

Log "Logging into Data Management Service..."
$User = LoginUser -config $Config
if ($User.Ticket -ne $Null -and $User.Ticket.Length -ge 10) {
    Log "Retrieved Login Ticket" "Green"
}
else {
    Write-Error "Error: Method:LoginUser: Retrieving Login Ticket: $($User.FailedLoginMessage)"
}

#Test below this line
#########################################################################


<#
 $Test = GetSelectMethod -config $Config -user $User -service $COnfig.PartitionMethod -method "SELECT_1200"

if ($test.Result) {
    $Test.Result | ft
}
else {
    "Result = NULL"
}
#>
$partition = @{
    Partition_ID = 13
    Location_ID = $null
}
$Test2 = GetPartitionChildren -config $Config -user $User -partition $partition
if ($test2.Result) {
    $Test2.Result | ft Partition_Name,Partition_ID,Parent_Partition_ID
}
else {
    "Result = NULL"
}