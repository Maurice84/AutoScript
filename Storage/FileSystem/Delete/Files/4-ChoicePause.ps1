Function Storage-FileSystem-Delete-Files-4-SelectPause {
    # ============
    # Declarations
    # ============
    $Task = "Pause processing when a file is found"
    # =========
    # Execution
    # =========
    [string]$global:PauseProcessing = Script-Function-ChoicePause -Name ($MyInvocation.MyCommand).Name -Task $Task
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:PauseProcessing
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}