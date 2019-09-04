Function Messaging-Exchange-Create-Domains-3-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
      Script-Function-Confirmation `
        -Name ($MyInvocation.MyCommand).Name `
    )
}