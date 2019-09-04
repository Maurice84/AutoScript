Function Storage-FileSystem-Delete-Files-3-InputFilter {
    # ============
    # Declarations
    # ============
    $Task = "Select which file(s) you would like to delete"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "  > Please enter the search filter for the file(s) what needs to be deleted (begin with . when using extensions): "
        $global:SearchFilter = Read-Host
        $InputKey = $global:SearchFilter
        If (!$InputKey) {
            Write-Host "    Please enter the search filter" -ForegroundColor Red
            $InputKey = $null
        }
    } Until ($InputKey)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:SearchFilter
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}