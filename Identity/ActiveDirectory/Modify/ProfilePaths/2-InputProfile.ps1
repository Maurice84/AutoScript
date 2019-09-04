Function Identity-ActiveDirectory-Modify-ProfilePaths-2-InputProfile {
    # ============
    # Declarations
    # ============
    $Task = "Please enter the UNC path to the mandatory profile (leave blank for UPD)"
    # =========
    # Execution
    # =========
    Do {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        $global:ProfilePath = Read-Host
        If ($global:ProfilePath) {
            If (!(Test-Path ($global:ProfilePath + ".V2"))) {
                Write-Host "  The UNC path could not be reached. Please enter the correct UNC path" -ForegroundColor Red
                $Pause.Invoke()
            }
            Else {
                $Choice = $true
            }
        }
        Else {
            $global:ProfilePath = "n.a."
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:ProfilePath
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}