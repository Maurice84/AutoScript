Function Identity-ActiveDirectory-Create-Groups-5-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # -----------------------------------------------------------------
    Write-Host "Step 1/2 - Creating new group" -ForegroundColor Magenta
    # -----------------------------------------------------------------
    Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
        "New-ADGroup -Name '$global:GroepNaam'`
        -SamAccountName '$global:GroepNaam'`
        -GroupCategory Security`
        -GroupScope Global`
        -DisplayName '$global:GroepNaam'`
        -Path '$global:GroupOUPath'`
        -Server '$global:PDC'")
    )
    If ($?) {
        Write-Host -NoNewLine "- OK: Group "; Write-Host -NoNewLine $global:GroepNaam -ForegroundColor Yellow; Write-Host -NoNewLine " successfully created in "; Write-Host -NoNewLine $global:Subtree -ForegroundColor Yellow; Write-Host
    }
    Else {
        Write-Host "- ERROR: An error occurred creating group $global:MailboxGroep in $global:Subtree, please investigate!" -ForegroundColor Magenta
        $Problem = $true
    }
    if (!$Problem) {
        Write-Host
        # ----------------------------------------------------------------------------------------
        Write-Host "Step 2/2 - Adding domain account(s) to the new group" -ForegroundColor Magenta
        # ----------------------------------------------------------------------------------------
        $Membership = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
        "Get-ADGroupMember -Identity '$global:SelectieGroep' -Recursive -Server '$global:PDC'")
        )
        ForEach ($Account in $Membership) {
            $AccountSamAccountName = $Account.SamAccountName
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                    "Add-ADGroupMember '$global:GroepNaam' '$AccountSamAccountName' -Server '$global:PDC'")
            )
            If ($?) {
                Write-Host -NoNewLine "- OK: Account "; Write-Host -NoNewLine $Account.Name -ForegroundColor Yellow; Write-Host -NoNewLine " successfully added to the group "; Write-Host $GroepNaam -ForegroundColor Yellow
            }
            Else {
                Write-Host
                Write-Host ("- ERROR: An error occurred adding account " + $Account.Name +" to the group, please investigate!") -ForegroundColor Magenta
                $Pause.Invoke()
                BREAK
            }
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}