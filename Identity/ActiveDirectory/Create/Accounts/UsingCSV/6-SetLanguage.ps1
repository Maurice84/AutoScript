Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-6-SetLanguage {
    # ============
    # Declarations
    # ============
    $Task = "Select a language for the domain account(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "    1. Dutch      2. English     3. German" -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Select an option or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
        [string]$InputChoice = Read-Host
        $InputKey = @("1", "2", "3", "X") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letter/numbers from above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
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
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Taalkeuze
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}