Function Identity-ActiveDirectory-Delete-Accounts-4-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
            Script-Function-Confirmation `
            -Name ($MyInvocation.MyCommand).Name `
    )
}