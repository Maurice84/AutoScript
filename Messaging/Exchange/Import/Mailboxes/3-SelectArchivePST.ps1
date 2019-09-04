Function Messaging-Exchange-Import-Mailboxes-3-SelectArchivePST {
    # ============
    # Declarations
    # ============
    $Task = "Select the PST-file to import into a mailbox"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "PST" -Type "Bestand" -Functie "Selecteren"
    $global:File = $global:Array.PST
    $global:Path = $global:Array.Name
    $global:UNC = $global:Array.FileUNC
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:File
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}