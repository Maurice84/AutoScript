Function Software-SyncBackPro-Create-Profiles-ForAccounts-3-SelectTemplate {
    # ============
    # Declarations
    # ============
    $Task = "Select a SyncBackPro profile you like use as a template"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Bestand" -Functie "Selecteren" -Filter "SyncBackPro"
    $global:FileTemplate = $global:Array.Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $FileTemplate.Split('\')[-1]
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}