Function Identity-ActiveDirectory-Modify-Accounts-4-SetGroups {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Selected account: "; Write-Host $global:WeergaveNaam -ForegroundColor Yellow
    Write-Host
    If (!$global:OUSelectie) {
        Do {
            If ($global:OUFilter -ne "*") {
                Write-Host -NoNewLine "  Would you like to use the previous entered OU filter: "; Write-Host -NoNewLine $global:OUFilter -ForegroundColor Yellow; Write-Host -NoNewLine " to select groups? (Y/N): "
            }
            Else {
                Write-Host -NoNewLine "  Would you like to search groups through all Organizational Units? (Y/N): "
            }
            $InputChoice = Read-Host
            $InputKey = @("Y"; "N") -contains $InputChoice
            If (!$InputKey) {
                Write-Host "  Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($InputKey)
        Write-Host -NoNewLine "  "
        Switch ($InputChoice) {
            "N" {
                Write-Host -NoNewLine ("Please enter the name (or a part) of the OU with groups, or press Enter for all groups: ")
                $global:OUFilter = Read-Host
                Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
                If ($global:OUFilter.Length -eq 0) {
                    $global:OUFilter = "*"
                    $global:OUFormat = "All groups"
                }
                Else {
                    $global:OUFormat = $global:OUFilter
                }
            }
        }
        $global:OUSelectie = $true
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $global:OUFilter -Type "Groepen" -Functie "Markeren" -Objecten $global:GroepenArray
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    [array]$global:GroepenNieuw = $global:Array
    $global:GroepenHuidig = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "(Get-ADPrincipalGroupMembership '$global:PreWin2000Naam' | Where-Object {`$`_.Name -ne `"Domain Users`" -AND `$`_.Name -ne `"Domeingebruikers`"} | Select-Object Name | Sort-Object Name).Name")
    )
    $global:GroepenNieuwNamen = @()
    $global:GroepenNieuwNamen = $GroepenNieuw.Name
    $global:Fout = $null
    ForEach ($GroepName in $GroepenNieuwNamen) {
        # -----------------------
        # Adding group membership
        # -----------------------
        If ($GroepenHuidig -notcontains $GroepName) {
            $AddGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Add-ADGroupMember -Identity $GroepName -Members '$global:PreWin2000Naam'")
            )
            $GroepenHuidig = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "(Get-ADPrincipalGroupMembership '$global:PreWin2000Naam' | Where-Object {`$`_.Name -ne `"Domain Users`" -AND `$`_.Name -ne `"Domeingebruikers`"} | Select-Object Name | Sort-Object Name).Name")
            )
            If ($GroepenHuidig -contains $GroepName) {
                Write-Host -NoNewLine "OK: Group "; Write-Host -NoNewLine $GroepName -ForegroundColor Yellow; Write-Host " has been added to the account"
            }
            Else {
                Write-Host -NoNewLine "ERROR: An error occurred adding group $GroepName to the account, please investigate!" -ForegroundColor Magenta
                Write-Host
                $Pause.Invoke()
                Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 2]
            }
        }
    }
    ForEach ($GroepName in $GroepenHuidig) {
        # -------------------------
        # Removing group membership
        # -------------------------
        If ($GroepenNieuwNamen -notcontains $GroepName) {
            $RemoveGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Remove-ADGroupMember -Identity $GroepName`
                -Members $global:PreWin2000Naam`
                -WarningAction SilentlyContinue`
                -ErrorAction SilentlyContinue`
                -Confirm:0")
            )
            $GroepenHuidig = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "(Get-ADPrincipalGroupMembership '$global:PreWin2000Naam' | Where-Object {`$`_.Name -ne `"Domain Users`" -AND `$`_.Name -ne `"Domeingebruikers`"} | Select-Object Name | Sort-Object Name).Name")
            )
            If ($GroepenHuidig -notcontains $GroepName) {
                Write-Host -NoNewLine "OK: Group "; Write-Host -NoNewLine $GroepName -ForegroundColor Yellow; Write-Host " has been removed from the account"
            }
            Else {
                Write-Host ("ERROR: An error occurred removing $GroepName from the account, please investigate!") -ForegroundColor Magenta
                Write-Host
                $Pause.Invoke()
                Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 2]
            }
        }
    }
    Script-Module-ReplicateAD
    $global:GroepenArray = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
        "(Get-ADPrincipalGroupMembership '$PreWin2000Naam' | Where-Object {`$`_.Name -ne `"Domain Users`" -AND `$`_.Name -ne `"Domeingebruikers`"} | Select-Object Name | Sort-Object Name).Name")
    )
    If ($global:GroepenArray.Count -gt 5) {
        $global:Groepen = [string]$global:GroepenArray.Count + " groups"
    }
    Else {
        $global:Groepen = $global:GroepenArray -join ', '
    }
    Write-Host
    Write-Host "Now returning to the group selection menu." -ForegroundColor Green
    $Pause.Invoke()
    If (!$global:GroepenArray) {
        $global:Groepen = "n.a."
        Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
    }
    Else {
        Invoke-Expression -Command ($MyInvocation.MyCommand).Name
    }
}