Function Script-Function-InputServers {
    param (
        [string]$Name,
        [string]$Task
    )
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name $Name
    Write-Host -NoNewLine "  > Please enter the server name(s) you would like to use (comma-separated) or leave blank for current server: " -ForegroundColor Yellow
    $InputServers = Read-Host
    if ($InputServers) {
        $InputServers = $InputServers -Replace ', ', ','
        $InputServers = $InputServers -Split ','
    }
    else {
        $InputServers = $env:computername
    }
    foreach ($Server in $InputServers) {
        $Servers += $Server
    }
    # ==========
    # Finalizing
    # ==========
    return $Servers
}