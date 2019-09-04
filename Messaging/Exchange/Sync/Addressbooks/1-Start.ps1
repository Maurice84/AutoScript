Function Messaging-Exchange-Sync-Addressbooks-1-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    # ----------------------------------------------------------------
    # Updating the (Global) Address Lists and Offline Address Books(s)
    # ----------------------------------------------------------------
    Write-Host -NoNewLine "- Updating the Global Address List(s)..." -ForegroundColor Gray
    Script-Connect-Server -Module "Exchange" -Command (
        [scriptblock]::Create("Get-GlobalAddressList | Update-GlobalAddressList -WarningAction SilentlyContinue")
    )
    Write-Host; Write-Host -NoNewLine "- Updating the Address List(s)..." -ForegroundColor Gray
    Script-Connect-Server -Module "Exchange" -Command (
        [scriptblock]::Create("Get-AddressList | Update-AddressList -WarningAction SilentlyContinue")
    )
    Write-Host; Write-Host -NoNewLine "- Updating the Offline Address Book(s)..." -ForegroundColor Gray
    Script-Connect-Server -Module "Exchange" -Command (
        [scriptblock]::Create("Get-OfflineAddressBook | Update-OfflineAddressBook -WarningAction SilentlyContinue")
    )
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}