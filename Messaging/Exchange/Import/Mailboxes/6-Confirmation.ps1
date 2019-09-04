Function Messaging-Exchange-Import-Mailboxes-6-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
        Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
    )
}