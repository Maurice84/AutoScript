Function Script-Module-SetCredentials {
    param (
        [string]$Credentials,
        [string]$Type,
        [string]$Office365Admin,
        [string]$OnPremiseAdmin
    )
    # =========
    # Execution
    # =========
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "CredentialManager"
    # -------------------
    # Loading credentials
    # -------------------
    If ($Credentials -ne "Aanmaken") {
        If ($CredManager) {
            $global:CredTarget = $Credentials
            $UserName = (Get-StoredCredential -Target $global:CredTarget -WarningAction SilentlyContinue).UserName
            If ($global:CredTarget -like "Office365-*") {
                $Office365Admin = $UserName
            }
            If ($global:CredTarget -like "OnPremise-*") {
                $OnPremiseAdmin = $UserName
            }
            $global:Credential = Get-StoredCredential -Target $global:CredTarget -WarningAction SilentlyContinue
            If ($Office365Admin) {
                $global:Office365Credential = $global:Credential
            }
            If ($OnPremiseAdmin) {
                $global:OnPremiseCredential = $global:Credential
            }
        } Else {
            $global:Credential = Get-Credential -Message "Please enter the $Type admin credentials"
        }
    }
    # ------------------
    # Create credentials
    # ------------------
    If ($Credentials -eq "Aanmaken") {
        If ($Office365Admin) {
            $Type = "Office365"
            $User = $Office365Admin
        }
        If ($OnPremise) {
            $Type = "OnPremise"
            $User = $OnPremiseAdmin
        }
        Do {
            $global:GetCred = Get-Credential -Message "Please enter the $Type admin credentials" -User $User
            If (!$global:GetCred) {
                Write-Host "Please use above as input" -ForegroundColor Red;Write-Host
            }
        } Until ($GetCred)
        If ($CredManager) {
            $Target = $Type + "-" + $GetCred.UserName
            New-StoredCredential -Credential $GetCred -Target $Target | Out-Null
        }
    }
}