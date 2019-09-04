Function Identity-ActiveDirectory-Modify-ProfilePaths-5-SetLanguage {
    # ============
    # Declarations
    # ============
    $Task = "Select a language for the accounts where the Country attribute has not been populated"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "    1. Dutch      2. English     3. German" -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Select an option or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
        [string]$Choice = Read-Host
        $Input = @("1", "2", "3", "X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letter/numbers from above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "1" {
            $global:Taalkeuze = "Nederlands"
            $global:Language = "NL"
        }
        "2" {
            $global:Taalkeuze = "Engels"
            $global:Language = "US"
        }
        "3" {
            $global:Taalkeuze = "Duits"
            $global:Language = "DE"
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Taalkeuze
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}