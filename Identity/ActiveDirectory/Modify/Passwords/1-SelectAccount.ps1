Function Identity-ActiveDirectory-Modify-Passwords-1-SelectAccount {
    # ============
    # Declarations
    # ============
    $Task = "Select accounts to modify the passwords"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $global:OUFilter -Type "AccountsBulk" -Functie "Markeren"
    Write-Host
    Do {
        Write-Host -NoNewLine "Are you certain you want to change the passwords for the "; Write-Host -NoNewLine "SELECTED ACCOUNTS" -ForegroundColor Red; Write-Host -NoNewLine "?"; Write-Host
        Write-Host -NoNewLine "Use "; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to change the OU, or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N/X): "
        [string]$Choice = Read-Host
        $Input = @("Y", "N", "X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    # ==========
    # Finalizing
    # ==========
    Switch ($Choice) {
        "N" {
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "Y" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
}