Function Identity-ActiveDirectory-Create-Groups-4-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
    Script-Function-Confirmation `
      -Name ($MyInvocation.MyCommand).Name `
    )
}