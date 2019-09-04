Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-7-SetGroups {
    # ============
    # Declarations
    # ============
    $Task = "Select security groups for the domain account(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -ne "Mailbox account") {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $Filter -Type "Groepen" -Functie "Markeren"
        $global:Groups = $global:Array
        $global:AantalGroups = $Aantal
    } Else {
        $global:AantalGroups = "n.a."
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:AantalGroups
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}