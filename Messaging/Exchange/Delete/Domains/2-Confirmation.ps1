Function Messaging-Exchange-Delete-Domains-2-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
      Script-Function-Confirmation `
        -Name ($MyInvocation.MyCommand).Name `
    )
}