Function Identity-ActiveDirectory-Modify-ProfilePaths-3-InputHomefolder {
    # ============
    # Declarations
    # ============
    $Task = "Please enter the UNC path to the homefolders"
    # =========
    # Execution
    # =========
    Do {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        $global:HomefolderPath = Read-Host
        If (!(Test-Path $global:HomefolderPath)) {
            Write-Host "  The UNC path could not be reached. Please enter the correct UNC path" -ForegroundColor Red
            $Pause.Invoke()
        }
        Else {
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:HomeFolderPath
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
