Function Identity-ActiveDirectory-Modify-ProfilePaths-4-InputDrive {
    # ============
    # Declarations
    # ============
    $Task = "Please enter the drive letter to the homefolder of the domain account(s)"
    # =========
    # Execution
    # =========
    Do {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        $global:HomefolderDrive = Read-Host
        If ($global:HomefolderDrive.Length -ne 1) {
            Write-Host "  The drive letter should only be 1 character" -ForegroundColor Red
            $Pause.Invoke()
        }
        Else {
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:HomeFolderDrive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}