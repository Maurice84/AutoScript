Function Messaging-Exchange-Create-Mailboxes-6-SetPermissionsGroupOU {
    # ============
    # Declarations
    # ============
    $Task = "Select an OU where the permission group must be created"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    $global:OUPath = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
        "(Get-ADUser '$SamAccountName' | Select-Object @{n='ParentContainer';e={`$`_.DistinguishedName -replace '^.+?,(CN|OU.+)','`$1'}}).ParentContainer")
    )
    Script-Index-OU
    $global:MailboxGroep = "MBX-" + $global:SamAccountName
    $global:GroupOUPath = $global:OUPath
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ($global:Subtree + " (" + $global:MailboxGroep + ")")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}