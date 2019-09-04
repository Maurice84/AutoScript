Function Identity-ActiveDirectory-Create-Groups-2-InputName {
    # ============
    # Declarations
    # ============
    $Task = "Select a name for the group"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewline "  > Please enter a new name for the group" -ForegroundColor Yellow
        $global:GroepNaam = Read-Host
        $CheckGroepNaam = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADGroup -Filter {Name -eq '$global:GroepNaam'} -Server '$global:PDC'")
        )
        If ($CheckGroepnaam) {
            Write-Host "    This group already exist, please enter an unused group name" -ForegroundColor Red; Write-Host
        }
        Else {
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:GroepNaam
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}