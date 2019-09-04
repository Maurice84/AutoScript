Function Identity-ActiveDirectory-Recover-Accounts-2-SelectDrive {
    $global:Titel = ($MyInvocation.MyCommand).Name
    Script-Module-SetHeaders -Name $Titel
    $Task = "Select a (network)drive where the CSV-file with archived domain account(s) is located"
    $global:Subtitel = {
        Write-Host ("- " + $Task + ":")
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "CSV" -Type "Drive" -Functie "Selecteren"
    $global:TitelDrive = "Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '$Drive' -ForegroundColor Yellow"
    $global:TitelDrive = [scriptblock]::Create($global:TitelDrive)
    Identity-ActiveDirectory-Recover-Accounts-3-SelectCSV
}