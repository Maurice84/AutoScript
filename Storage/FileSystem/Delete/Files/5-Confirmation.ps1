Function Storage-FileSystem-Delete-Files-5-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
      Script-Function-Confirmation `
        -Name ($MyInvocation.MyCommand).Name `
      )
}