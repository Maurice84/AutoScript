Function Identity-ActiveDirectory-Create-Groups-1-SelectGroup {
    # ============
    # Declarations
    # ============
    $Task = "Select a group to duplicate"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Server $global:PDC -Type "Groepen" -Functie "Selecteren"
    $global:SelectieGroep = $global:Array.Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:SelectieGroep
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
