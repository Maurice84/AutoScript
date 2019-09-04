Function Messaging-Exchange-Archive-Mailboxes-3-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
            Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
    )
}