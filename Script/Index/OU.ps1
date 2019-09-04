Function Script-Index-OU {
    param (
        [string]$Path
    )
    # -----------------
    # Declare variables
    # -----------------
    $InputKey = $null
    $InputSelection = $null
    $global:Array = $null
    $global:Array = @()
    $Objects = $null
    $Objects = @()
    Write-Host -NoNewLine "Loading domain OUs..."
    If (!$Path) {
        $Path = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADDomain).DistinguishedName"))
    }
    [array]$Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADOrganizationalUnit -SearchBase '$Path' -SearchScope Onelevel -Filter * | Where-Object {`
        `$`_.Name -notlike 'Dom*Controllers' -AND `$`_.Name -notlike '*Exchange*'} | Select-Object Name, DistinguishedName")
    )
    If ($Objects.Length -eq 1) {
        $Path = $Objects.DistinguishedName
    }
    Do {
        # -----------------------------------------------------------------------------
        # Converting DistinguishedName (OU=..,OU=...,DC=...) to CanonicalName (\OU\OU\)
        # -----------------------------------------------------------------------------
        $global:Subtree = Script-Convert-DN-to-CN -Path $Path -Debug $true
        [array]$Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADOrganizationalUnit -SearchBase '$Path' -SearchScope Onelevel -Filter * | Where-Object {`
            `$`_.Name -notlike 'Dom*Controllers' -AND `$`_.Name -notlike '*Exchange*'} | Select-Object Name, DistinguishedName")
        )
        Script-Module-SetHeaders -Name $Titel
        $Subtitel.Invoke()
        Write-Host
        Write-Host -NoNewLine "  Current selected OU: "; Write-Host -NoNewLine $global:Subtree -ForegroundColor Yellow; Write-Host
        Write-Host
        $global:Array = @()
        $MaxLength = 0
        $Counter = 1
        ForEach ($OUObject in $Objects) {
            $Properties = @{
                Nr   = $Counter;
                Name = $OUObject.Name;
                DN   = $OUObject.DistinguishedName
            }
            $global:ArrayOUObject = New-Object PSObject -Property $Properties
            $global:Array += $global:ArrayOUObject
            $Counter++
            If ($OUObject.Name.Length -gt $MaxLength) {
                $MaxLength = $OUObject.Name.Length
            }
        }
        If ($global:Array) {
            If ($global:Array.Count -le 10) {
                $ColumnSize = $global:Array.Count
            }
            Else {
                $ColumnSize = "{0:F0}" -f ($global:Array.Count / 2)
            }
            For ($CounterC1 = 0; $CounterC1 -lt $ColumnSize; $CounterC1++) {
                $NrC1 = $global:Array[$CounterC1].Nr
                $NrC2 = $global:Array[$CounterC1 + $ColumnSize].Nr
                If ($NrC1 -ge 100) { $NrC1 = (" " * 2) + $NrC1 + ". " } ElseIf ($NrC1 -ge 10) { $NrC1 = (" " * 3) + $NrC1 + ". " } Else { $NrC1 = (" " * 4) + $NrC1 + ". " }
                If ($NrC2 -ge 100) { $NrC2 = (" " * 2) + $NrC2 + ". " } ElseIf ($NrC2 -ge 10) { $NrC2 = (" " * 3) + $NrC2 + ". " } ElseIf (!$NrC2) { $NrC2 = $null } Else { $NrC2 = (" " * 4) + $NrC2 + ". " }
                $ColumnText = [string]$NrC1
                $OrgUnitC1 = $global:Array[$CounterC1].Name
                $OrgUnitC2 = $global:Array[$CounterC1 + $ColumnSize].Name
                $ColumnText += $OrgUnitC1
                If ($global:Array.Count -gt 10) {
                    $Length = $MaxLength - $OrgUnitC1.Length
                    $Spaces = (" " * $Length)
                    $ColumnText += $Spaces + $NrC2 + $OrgUnitC2
                }
                Write-Host $ColumnText -ForegroundColor Yellow
            }
        }
        Else {
            Write-Host -NoNewLine "  There are no OUs found, please select this or go to the previous OU" -ForegroundColor Gray; Write-Host
        }
        Write-Host
        Do {
            Write-Host -NoNewline "  Browse through the OUs and select your choice using "; Write-Host -NoNewline "Y" -ForegroundColor Yellow; Write-Host -NoNewline " or use "; Write-Host -NoNewline "B" -ForegroundColor Yellow; Write-Host -NoNewline " to go back to the previous OU"; Write-Host
            Write-Host -NoNewLine "  Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
            [string]$InputChoice = Read-Host
            If ($InputChoice -as [int]) {
                If ([int]$InputChoice -eq "0" -OR [int]$InputChoice -gt $global:Array.Count) {
                    Write-Host "  Please use the letters above as input" -ForegroundColor Red; Write-Host
                }
                Else {
                    $InputKey = $InputChoice
                    $Path = $global:Array[$InputChoice - 1].DN
                }
            }
            Else {
                If ($InputChoice -eq "X") {
                    If ($Titel -like "*ActiveDirectory*Modify*Account*") {
                        Identity-ActiveDirectory-Modify-Accounts-3-Menu
                    }
                    Else {
                        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
                    }
                }
                If ($InputChoice -eq "Y") {
                    $InputKey = $InputChoice
                    $InputSelection = "Ja"
                }
                If ($InputChoice -eq "B") {
                    $InputKey = $InputChoice
                    Write-Host ($Path.Split(",")).Count -ForegroundColor Yellow
                    if (($Path.Split(",")).Count -gt 3) {
                        $Path = $Path.Substring($Path.IndexOf(','))
                        $Path = $Path.Substring(1)
                    }
                    else {
                        $global:Subtree = $null
                    }
                }
            }
        } Until ($InputKey)
    } While (!$InputSelection)
    $global:OUPath = $Path
    # -----------------------------------------------------------------------------------
    # Detectie van een prefix (Easy-Cloud Entry/Custom of On-Premise), op basis van de OU
    # -----------------------------------------------------------------------------------
    #If ($global:OUPath -like "*OU=Customers,OU=EasyCloud Entry*") {
    #    $global:Prefix = $global:OUPath.Split(',')[-6].Split('=')[1]
    #} Else {
    #$global:Prefix = $global:OUPath.Split(',')[-3].Split('=')[1]
    #}
    #If ($global:Prefix -like "*-*") {
    #    $global:Prefix = $Prefix.Split('-')[0]
    #    If ($global:Prefix.Length -ge 4) {
    #        $global:Prefix = $null
    #    }
    #}
    #Else {
    #    $global:Prefix = $null
    #}
}
