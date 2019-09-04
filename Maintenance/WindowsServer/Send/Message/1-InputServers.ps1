Function Maintenance-WindowsServer-Send-Message-1-InputServers {
    # ============
    # Declarations
    # ============
    $Task = "Select servers to send a message to all users"
    # =========
    # Execution
    # =========
    $global:Servers = Script-Function-InputServers -Name ($MyInvocation.MyCommand).Name -Task $Task
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ([string]$Servers.Count + " server(s)")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}