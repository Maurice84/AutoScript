Function Identity-ActiveDirectory-Purge-Accounts-2-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
      Script-Function-Confirmation `
        -Name ($MyInvocation.MyCommand).Name `
    )
}