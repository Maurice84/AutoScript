Function Identity-ActiveDirectory-Create-Accounts-1-SetType {
    # ============
    # Declarations
    # ============
    $Task = "Select a profile type for the new account"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "  P. Profile duplication " -ForegroundColor Magenta; Write-Host "(recommended)" -ForegroundColor Gray;
    Write-Host
    Write-Host -NoNewLine "  Standard profiles:" -ForegroundColor Gray; Write-Host
    Write-Host -NoNewLine "  1. Domain user " -ForegroundColor Yellow; Write-Host "(Remote Desktop workspace)" -ForegroundColor Gray
    Write-Host -NoNewLine "  2. Mailbox account " -ForegroundColor Yellow; Write-Host "(Shared mailbox)" -ForegroundColor Gray
    Write-Host -NoNewLine "  3. Webmail user " -ForegroundColor Yellow; Write-Host "(OWA or Outlook access only)" -ForegroundColor Gray
    Write-Host
    Write-Host -NoNewLine "  Custom profiles:" -ForegroundColor Gray; Write-Host
    Write-Host -NoNewLine "  4. Admin account " -ForegroundColor Yellow; Write-Host "(3rd party/consultants)" -ForegroundColor Gray
    Write-Host -NoNewLine "  5. Test user " -ForegroundColor Yellow; Write-Host "(testing applications and settings)" -ForegroundColor Gray
    Write-Host -NoNewLine "  6. Service account " -ForegroundColor Yellow; Write-Host "(services)" -ForegroundColor Gray
    Write-Host
    Write-Host
    Do {
        Write-Host -NoNewLine "  Please make a choice. Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
        $InputChoice = Read-Host
        $InputKey = @("1"; "2"; "3"; "4"; "5"; "6"; "P"; "X") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "P" { $global:Profiel = "Profielkopie" }
        "1" { $global:Profiel = "Domein gebruiker" }
        "2" { $global:Profiel = "Mailbox account" }
        "3" { $global:Profiel = "Webmail gebruiker" }
        "4" { $global:Profiel = "Beheerder account" }
        "5" { $global:Profiel = "Test gebruiker" }
        "6" { $global:Profiel = "Service account" }
        "X" { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Profiel
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}