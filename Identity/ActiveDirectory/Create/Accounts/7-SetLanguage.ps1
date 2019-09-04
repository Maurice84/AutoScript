Function Identity-ActiveDirectory-Create-Accounts-7-SetLanguage {
    $global:Titel = ($MyInvocation.MyCommand).Name
    Script-Module-SetHeaders -Name $Titel
    $Task = "Select a language for the new account"
    If ($Profiel -eq "Profielkopie" -OR $Profiel -eq "Service account") {
        If ($Profiel -eq "Profielkopie") {
            If ($Language -eq "NL") {
                $global:Taalkeuze = "Nederlands"
            }
            If ($Language -eq "US") {
                $global:Taalkeuze = "Engels"
            }
            If ($Language -eq "DE") {
                $global:Taalkeuze = "Duits"
            }
            If (!$Language) {
                $global:SetLanguage = $true
            }
            $global:TitelTaalinstelling = "Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '$Taalkeuze (gedetecteerd)' -ForegroundColor Yellow"
        }
        If ($Profiel -eq "Service account") {
            $global:TitelTaalinstelling = "Write-Host -NoNewLine ('- ' + '$Task' + ': n.a.') -ForegroundColor Gray"
        }
    }
    If ($SetLanguage -OR ($Profiel -ne "Profielkopie" -AND $Profiel -ne "Service account")) {
        $global:Subtitel = { $TitelProfiel.Invoke(); $TitelOU.Invoke(); $TitelGegevens.Invoke(); $TitelGebruikersnaam.Invoke();
            Write-Host -NoNewLine ("- " + $Task + ":") -ForegroundColor Yellow
        }
        $Subtitel.Invoke()
        Write-Host
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
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
        $global:TitelTaalinstelling = "Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '$Taalkeuze' -ForegroundColor Yellow"
    }
    $global:TitelTaalinstelling = [scriptblock]::Create($global:TitelTaalinstelling)
    Identity-ActiveDirectory-Create-Accounts-8-SetGroups
}