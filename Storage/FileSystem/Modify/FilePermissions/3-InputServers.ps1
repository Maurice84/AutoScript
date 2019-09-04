Function Storage-FileSystem-Modify-FilePermissions-3-InputServers {
    # ============
    # Declarations
    # ============
    $Task = "Select servers to set the permissions"
    # =========
    # Execution
    # =========
    [array]$global:Servers = Script-Function-InputServers -Name ($MyInvocation.MyCommand).Name -Task $Task
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ([string]$Servers.Count + " server(s)")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}