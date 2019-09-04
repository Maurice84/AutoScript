Function Messaging-Exchange-Import-Mailboxes-2-SelectPST {
    # ============
    # Declarations
    # ============
    $Task = "Select PST-files to import into the mailbox(es)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "PST" -Type "Bestand" -Functie "Markeren"
    Do {
        Write-Host
        Write-Host -NoNewLine "  Would you like to start the import of the "; Write-Host -NoNewLine "selected" -ForegroundColor Green; Write-Host -NoNewLine " PST-files? (Y/N): "
        [string]$Choice = Read-Host
        $Input = @("Y","N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    # ==========
    # Finalizing
    # ==========
    Switch ($Choice) {
        "Y" {
            $global:File = $null
            $global:Mailboxes = $global:Array
            $global:AantalPST = $global:Aantal
            Set-Variable $global:VarHeaderName -Value $global:AantalPST
            Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 5]
        }
        "N" {
            Write-Host
            Write-Host "  Now returning to the previous menu..."
            $Pause.Invoke()
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
        }
    }
}