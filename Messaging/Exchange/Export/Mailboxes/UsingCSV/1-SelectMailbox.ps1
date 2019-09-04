Function Messaging-Exchange-Export-Mailboxes-UsingCSV-1-SelectMailbox {
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
    Do {
        Write-Host -NoNewLine "- Would you like to manually "; Write-Host -NoNewLine "S" -ForegroundColor Yellow; Write-Host -NoNewLine ")elect mailbox(es) or like to use a ("; Write-Host -NoNewLine "C" -ForegroundColor Yellow; Write-Host -NoNewLine ")SV-file? Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
        [string]$Choice = Read-Host
        $Input = @("C", "S", "X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "S" {
            # -------------------------------------------------------------
            # Go to function Messaging > Export Mailboxes to export mailbox
            # -------------------------------------------------------------
            Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Export-Mailboxes-*-SelectMailbox" })
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "CSV" -Type "Drive" -Functie "Selecteren"
    $global:CSVDrive = $Drive
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:CSVDrive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}