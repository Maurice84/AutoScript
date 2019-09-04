Function Identity-ActiveDirectory-Recover-Accounts-1-SelectAccount {
    # ============
    # Declarations
    # ============
    $Task = "Select a domain account to recover"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "Would you like to use a "; Write-Host -NoNewLine "CSV-file" -ForegroundColor Yellow; Write-Host -NoNewLine " to recover a domain account? (Y/N): "
        $Choice = Read-Host
        $Input = @("Y"; "N") -contains $Choice
        If (!$Input) {
            Write-Host "  Please use the letters above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            $global:AccountHerstelCSV = $true
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
        "N" {
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "AccountHerstel" -Functie "Markeren"
            If ($CSV) {
                $global:AccountHerstelCSV = $true
                Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
            }
            Else {
                $global:ArrayAccounts = $global:Array
                If ($ArrayAccounts.Count -gt 1) {
                    $global:AantalAccounts = [string]$global:Aantal
                }
                Else {
                    $global:AantalAccounts = $ArrayAccounts.Name
                }
                # ==========
                # Finalizing
                # ==========
                Set-Variable $global:VarHeaderName -Value $global:AantalAccounts
                Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
                Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 3]
            }
        }
    }
}