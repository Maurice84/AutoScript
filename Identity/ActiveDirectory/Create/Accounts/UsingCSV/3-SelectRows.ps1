Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-3-SelectRows {
    #$global:Titel = ($MyInvocation.MyCommand).Name
    #Script-Module-SetHeaders -Name $Titel
    #$Task = "Selecteer de nieuwe domeinaccount(s) die u wilt aanmaken:"
    #$global:Subtitel = {$TitelDrive.Invoke();$TitelFile.Invoke();Write-Host ("- " + $Task + ":")}  
    #$Subtitel.Invoke()
    #Write-Host
    #Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $File -Type "CSV" -Functie "Markeren"
    #$global:Accounts = $global:Array
    #$global:AantalAccounts = $Aantal
    #$global:TitelAccounts = "Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '$AantalAccounts' -ForegroundColor Yellow"
    #$global:TitelAccounts = [scriptblock]::Create($global:TitelAccounts)
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}