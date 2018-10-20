﻿function Get-Properties($Object, $MaxLevels="5", $PathName = "`$JSON", $Level=0){
    if ($Level -eq 0) { 
        $oldErrorPreference = $ErrorActionPreference
        $ErrorActionPreference = "SilentlyContinue"
    }
    $props = @()
    $rootProps = $Object | Get-Member -ErrorAction SilentlyContinue | Where-Object { $_.MemberType -match "Property"} 
    $rootProps | ForEach-Object { $props += "$PathName.$($_.Name)" }
    if ($Level -lt $MaxLevels){
        $typesToExclude = "System.Boolean", "System.String", "System.Int32", "System.Char"
        $props += $rootProps | ForEach-Object {
                    $propName = $_.Name;
                    $obj = $($Object.$propName)
                    $type = ($obj.GetType()).ToString()
                    if (!($typesToExclude.Contains($type) ) ){
                        $childPathName = "$PathName.$propName"
                        if ($obj -ne $null) {
                            Get-Properties -Object $obj -PathName $childPathName -Level ($Level + 1) -MaxLevels $MaxLevels }
                        }
                    }
    }
    if ($Level -eq 0) {$ErrorActionPreference = $oldErrorPreference}
    $props
}

Function Sanitize-Object($object,$propsarray){
    $count = 0
    ForEach($propfilter in $propsarray){
       $props = $object | Where-Object{ $_ -like "*$propfilter" }
       ForEach($path in $props){
            [string]$pathstring = $path.Replace('$','')
            $count++
            ""
            "$count : Cheking Path: $pathstring"
            $splitpath = $path.Split(".")
            $propname = $splitpath[$($splitpath.length -1)]
            $rootpath = $path.Replace(".$propname","")
            $invoke = Invoke-Expression -command $path
            if($invoke){
                $invokecount = 0
                ForEach($item in $invoke){                
                    $invokepath = "$rootpath[$invokecount].$propname"
                    if($item){
                        Write-Host -ForegroundColor Yellow "   + enum $invokecount : value is not null: $item"
                        $nullcmd = $invokepath + ' = $null'
                        Invoke-Expression -command $nullcmd
                        if((Invoke-Expression -command $Invokepath) -eq $null){
                            Write-Host -ForegroundColor White "      + value is null; value= $(Invoke-Expression -command $Invokepath)"
                        }else{
                            Write-Host -ForegroundColor Red "      + value is not null; value = $(Invoke-Expression -command $Invokepath)"
                        }
                    }else{
                        write-host -ForegroundColor Cyan "   + enum $invokecount : value is already null"
                    }
                    $invokecount++  
                }
            }else{
                Write-Host -ForegroundColor White "   +   Path: $pathstring is null"
            }
        }    
    }
}

$JSON = (Get-Content 'C:\Users\MattH\OneDrive - BASELAYER\Deployment Scripts\config.json') -join "`n" | ConvertFrom-Json

$filters = @("Username","Password","LoginName","WebHost")

$FullObject = Get-Properties -Object $JSON

Sanitize-Object -object $FullObject -propsarray $filters


