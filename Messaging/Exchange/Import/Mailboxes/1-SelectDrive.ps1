Function Messaging-Exchange-Import-Mailboxes-1-SelectDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select the (network)drive where the CSV-file(s) with mailbox(es) is located"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "PST" -Type "Drive" -Functie "Selecteren"
    Do {
        Write-Host
        Write-Host -NoNewLine "  Would you like to import mail-items from a PST-file into a folder in a mailbox? You can select a mailbox and enter a foldername (Y/N): "
        [string]$Choice = Read-Host
        $Input = @("Y", "N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Drive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Switch ($Choice) {
        "Y" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 2]
        }
        "N" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
    }
}