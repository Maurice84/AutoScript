Function Storage-FileSystem-Restore-3-SelectTargetDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select the target (network)drive where you want to restore to"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Drive" -Functie "Selecteren"
    $global:DriveTarget = $global:Drive
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:DriveTarget
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}