Function WSLoginUser ($config){
    $URI = "$($config.WebHost)/$($config.LoginMethod)"
    $JSON = [pscustomobject]@{
       LoginName = $config.Username
       NewPassword = $null
       Password = $config.Password 
    }

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }
    Invoke-RestMethod @params
}

Export-ModuleMember WSLoginUser

Function WSGetPartition ($config, $user, $partition){
    $URI = "$($config.WebHost)/$($config.ExecMethod)"
    $JSON = $Config.JSONGetLocInfo
    $JSON.Ticket = $user.Ticket

    $params = @{
        Uri         = $URI
        Method      = "Post"
        Body        = ($JSON | ConvertTo-Json)
        ContentType = "application/json; charset=utf-8"
        ErrorAction = "Inquire"
    }
    Invoke-RestMethod @params
}

Export-ModuleMember WSGetPartition