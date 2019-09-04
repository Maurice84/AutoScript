Function Identity-ActiveDirectory-Modify-ProfilePaths-6-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
    Script-Function-Confirmation `
      -Name ($MyInvocation.MyCommand).Name `
    )
}