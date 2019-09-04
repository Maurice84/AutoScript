Function Messaging-Exchange-Create-Domains-4-Start {
    # =========
    # Execution
    # =========
    If (!$CSV) {
        Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    }
    # ---------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/2 - Creating accepted mail domain" -ForegroundColor Magenta; Write-Host
    # ---------------------------------------------------------------------------------------------------
    $CheckDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain -DomainController '$PDC'"))
    If (!($CheckDomain | Where-Object { $_.DomainName -eq $Emaildomein }) -AND !($CheckDomain | Where-Object { $_.Name -eq $NameEmaildomein })) {
        $AddDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-AcceptedDomain -Name '$NameEmaildomein' -DomainName '$Emaildomein' -DomainType:Authoritative -DomainController '$PDC'"))
        $CheckDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain -DomainController '$PDC'"))
        If ($CheckDomain | Where-Object { $_.Name -eq $NameEmaildomein }) {
            Write-Host -NoNewLine "OK: Mail domain "; Write-Host -NoNewLine $Emaildomein -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating mail domain $Emaildomein, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Mail domain $NameEmaildomein already exists" -ForegroundColor Gray
    }
    Write-Host
    # --------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/2 - Applying mail domain to selected mailbox" -ForegroundColor Magenta; Write-Host
    # --------------------------------------------------------------------------------------------------------------
    $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
    If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $true) {
        $AddEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddressPolicyEnabled 0 -DomainController '$PDC'"))
        $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $false) {
            Write-Host " - OK: Email Address Policy removed from selected mailbox"
        }
        Else {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Disable-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -Confirm:0"))
            Write-Host " - ERROR: An error occurred removing the Email Address Policy from the selected mailbox (required)! Therefore the mailbox has been deleted. Please recreate the mailbox correctly." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Email Address Policy already removed from selected mailbox" -ForegroundColor Gray
    }
    $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
    If ($CheckEmailAddress | Where-Object { $_.EmailAddresses -notlike "*smtp:$Email*" } | Select-Object EmailAddresses) {
        $AddEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses @{Add = '$Email' } -DomainController '$PDC'"))
        $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If ($CheckEmailAddress | Where-Object { $_.EmailAddresses -like "*smtp:$Email*" } | Select-Object EmailAddresses) {
            Write-Host -NoNewLine " - OK: Mail Address "; Write-Host -NoNewLine $Email -ForegroundColor Yellow; Write-Host " added successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred adding mail address $Email, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Mail Address $Email already added" -ForegroundColor Gray
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    If ($CSV) {
        Write-Host -NoNewLine "Proceeding to the next mail domain in 2 seconds..." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    Else {
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
}