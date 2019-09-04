Function Identity-ActiveDirectory-Delete-Accounts-3-SelectHomefolder {
    # ============
    # Declarations
    # ============
    $Task = "Select the homefolder(s) to archive"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "Folder" -Type "Homefolders" -Functie "Markeren" -Objecten $global:ArrayAccounts
    $global:ArrayHomefolders = $global:Array
    $global:AantalHomefolders = $global:Aantal + " (" + (($ArrayHomefolders.HomePath | Where-Object { $_ -notlike "Geen homefolder*" }) -join ", ") + ")"
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:AantalHomefolders
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}