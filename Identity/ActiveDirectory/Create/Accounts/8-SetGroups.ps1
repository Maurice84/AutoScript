Function Identity-ActiveDirectory-Create-Accounts-8-SetGroups {
    # ============
    # Declarations
    # ============
    $global:Groups = @()
    $Task = "Select group membership for the new account"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -ne "Profielkopie") {
        If ($global:Profiel -ne "Mailbox account" -AND $global:Profiel -ne "Webmail gebruiker") {
            If ($global:Profiel -eq "Beheerder account" -OR $global:Profiel -eq "Service account") {
                $global:Filter = "*"
            }
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $Filter -Type "Groepen" -Functie "Markeren"
            $global:Groups = $Array
        }
    }
    Else {
        Write-Host
        Write-Host -NoNewLine "Detecting group membership of the selected domain account..."
        $global:AllGroups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
            "Get-ADPrincipalGroupMembership '$SamAccountName'`
            | Where-Object {`$`_.Name -ne `"Domain Users`" -AND `$`_.Name -ne `"Domeingebruikers`"}`
            | Select-Object Name")
        )
        ForEach ($global:Group in $AllGroups) {
            $global:Properties = @{Name = $Group.Name }
            $global:Object = New-Object PSObject -Property $global:Properties
            $global:Groups += $global:Object
        }
    }
    # ==========
    # Finalizing
    # ==========
    If ($global:Groups.Length -eq 1) {
        $global:Aantal = [string]$Groups.Length + " group"
    }
    ElseIf ($global:Groups.Length -ge 2) {
        $global:Aantal = [string]$Groups.Length + " groups"
    }
    Else {
        $global:Aantal = "n.a."
    }
    Set-Variable $global:VarHeaderName -Value $global:Aantal
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
