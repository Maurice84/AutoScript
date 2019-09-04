Function Messaging-Exchange-Move-Mailboxes-7-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
        Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
    )
}