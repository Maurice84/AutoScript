Function Maintenance-WindowsServer-Send-Message-3-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
            Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
            -Arguments ' -Message $global:Message -Servers $global:Servers' `
    )
}