Function Storage-FileSystem-Delete-Files-2-SelectDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select the (network)drives where you want to delete $global:Choice"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If (!$global:Server) {
        Write-Host -NoNewLine "  > Please enter a server name to browse to the drives (or leave blank for current server): "
        $global:Server = Read-Host
        If (!$global:Server) {
            $global:Server = $env:computername
        }
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Functie "Selecteren" -Server $global:Server -Type "Drive"
    $global:Drive = $global:Drive
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Drive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}