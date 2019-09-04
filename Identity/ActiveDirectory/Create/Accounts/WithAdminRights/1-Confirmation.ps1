Function Identity-ActiveDirectory-Create-Accounts-WithAdminRights-1-Confirmation {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host " The following tasks will be executed:"
    Write-Host
    Write-Host " - Create Fine Grained Password Policy for Administrators (180 days)"
    Write-Host " - Create Administrator accounts with random passwords"
    Write-Host " - Duplicate Administrator permissions to the Administrator accounts"
    Write-Host " - Notify users with the created Administrator accounts"
    Write-Host
    Do {
        Write-Host -NoNewLine "Would you like to proceed? (Y/N): "
        [string]$InputChoice = Read-Host
        $InputKey = @("Y", "N") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "N" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}