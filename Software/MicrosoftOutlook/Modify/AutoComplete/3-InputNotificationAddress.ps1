Function Software-MicrosoftOutlook-Modify-AutoComplete-3-InputNotificationAddress {
    # ============
    # Declarations
    # ============
    $Task = "Email address for notification and to send test mail"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "  > Please enter an email address for notification and to send test mail (required to process AutoComplete files): "
        $global:TestEmailAddress = Read-Host
        $InputKey = $global:TestEmailAddress
        If (!$InputKey) {
            Write-Host "    Please enter an email address" -ForegroundColor Red
            $InputKey = $null
        }
    } Until ($InputKey)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TestEmailAddress
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}