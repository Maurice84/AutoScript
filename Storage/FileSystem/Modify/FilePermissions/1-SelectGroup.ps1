Function Storage-FileSystem-Modify-FilePermissions-1-SelectGroup {
    # ============
    # Declarations
    # ============
    $Task = "Select a group which you like to use"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -CurrentTask $Task -Functie "Selecteren" -FunctionName ($MyInvocation.MyCommand).Name -Server $global:PDC -Type "Groepen"
    $global:Group = ($global:Array | Select-Object Name).Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Group
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}