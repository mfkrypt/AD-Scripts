function Test-IsPortOpen {
    param(
        [string]$Name,
        [int]$Port
    )
    
    $mgr = New-Object -ComObject "HNetCfg.FwMgr"
    $allow = $null
    $mgr.IsPortAllowed($Name, 2, $Port, "", 6, [ref]$allow, $null)
    $allow
}

# . .\test_internal_ports.ps1
# Test-IsPortOpen -Name "MyApp" -Port 80