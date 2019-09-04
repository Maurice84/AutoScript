Function Software-MicrosoftOutlook-Modify-AutoComplete-4-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
    Script-Function-Confirmation `
      -Name ($MyInvocation.MyCommand).Name `
    )
}