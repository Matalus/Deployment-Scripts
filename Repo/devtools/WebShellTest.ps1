$ErrorActionPreference = "stop"

$RunDir = split-path -parent $MyInvocation.MyCommand.Definition

#Generic Logging function
Function Log($message) {
    "$(Get-Date -Format u) | $message"
}

#Set Current Root Dir
Set-Location $RunDir

#Load function library from modules folder
Log "Loading Function Library"
Remove-Module WSFunctions -ErrorAction SilentlyContinue
Import-Module ".\WSFunctions.psm1" -DisableNameChecking

#Load config from JSON in project root dir
Log "Loading Configuration..."
$Config = (Get-Content ".\wsconfig.json") -join "`n" | ConvertFrom-Json
#-------------------------------------------------------------------------------------------

Log "Authenticating with WebShell : $($config.WebHost)"
$user = WSLoginUser -config $Config
if($user.Ticket){
    Log "Successfully Retrived Login Ticket"
}else{
    Write-Error "Error Retrieving Login Ticiket : $($user.FailedLoginMessage)"
}

$params = @{
    config = $Config
    user = $user
    partition = $Null
}

Log "Getting Partition Data..."
$GetPartition = WSGetPartition @params
