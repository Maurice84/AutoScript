Function Messaging-Exchange-Grant-Admins-MailboxPermissions-1-SelectMailbox {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter * -Type "Mailbox" -Functie "Markeren" -SkipMenu $true
    Write-Host -NoNewLine "Detecting Administrators group..."
    $global:Mailboxes = $global:Array
    If ($global:Mailboxes) {
        $global:GroupAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADGroup -Filter * | Where-Object {`$`_.Name -like `"*-Admins`"}).Name"))
        If (!$global:GroupAdmins) {
            $global:GroupAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq `"Domain Admins`"}).Name"))
        }
        If ($global:GroupAdmins) {
            Write-Host "Found: $global:GroupAdmins"
        }
        Else {
            Write-Host
            Write-Host "ERROR: An error occurred detecting the Administrators group, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
        Do {
            Write-Host
            Write-Host -NoNewLine "Would you like to grant Full Access permission for the detected Administrators group to all mailbox(es)? (Y/N): " -ForegroundColor Yellow
            [string]$Choice = Read-Host
            $Input = @("Y", "N") -contains $Choice
            If (!$Input) {
                Write-Host "  Please use the letters above as input" -ForegroundColor Red
            }
        } Until ($Input)
        # ==========
        # Finalizing
        # ==========
        Switch ($Choice) {
            "Y" {
                Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
            }
            "N" {
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
    }
    Else {
        Write-Host "ERROR: An error occurred indexing the mailbox(es), please investigate!" -ForegroundColor Red
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
}