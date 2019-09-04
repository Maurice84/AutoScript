Function Maintenance-WindowsServer-Overview-Applications-3-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
        Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
            -Arguments ' -Applications $global:Applications -Servers $global:Servers' `
    )
}