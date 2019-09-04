Function Storage-FileSystem-Modify-AutoComplete-2-SelectCSV {
    # ============
    # Declarations
    # ============
    $Task = "Select the CSV-file(s) with domain account(s) which you want to use to rename the AutoComplete files"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "CSV" -Functie "Selecteren" -Type "Bestand"
    $global:FileCSV = ($global:Array | Select-Object Name).Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $FileCSV.Split('\')[-1]
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}