Function Messaging-Exchange-Move-Mailboxes-1-InputSource {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Do {
        $Choice = Read-Host "- Is the source an Exchange server located in this network? (Y/N)"
        $Input = @("Y", "N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange" -Forced $true
            Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
            $CheckHybrid = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-HybridConfiguration | Select-Object OnPremisesSmartHost"))
            If ($CheckHybrid -notlike "*.*") {
                $EWSVirtualDirectory = (Get-PSSession -Name "Exchange").ComputerName + "\EWS (Default Web Site)"
                $GetExternalUrl = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-WebServicesVirtualDirectory '$EWSVirtualDirectory' | Select-Object ExternalUrl"))
                $global:ServerSource = ($GetExternalUrl.ExternalUrl).Host
            }
            Else {
                $global:ServerSource = ($CheckHybrid.OnPremisesSmartHost).Domain
            }
        }
        "N" {
            Do {
                $Choice = Read-Host "- Is the source Office 365 (Exchange Online)? (Y/N)"
                $Input = @("Y", "N") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "Y" {  
                    $global:ServerSource = "Office 365"
                }
                "N" {
                    $global:ServerSource = Read-Host "- Please enter the external DNS name of an Exchange server as source (i.e. mail.customer.nl)"
                }
            }
        }
    }
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Please enter the source server: " -ForegroundColor Gray; Write-Host $global:ServerSource -ForegroundColor Yellow
    $Task = "Select an Administrator credential to connect to the source"
    If ($ServerSource -eq "Office 365") {
        $CredType = "Office365"
    }
    Else {
        $CredType = "OnPremise"
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $CredType -Type "Credentials" -Functie "Selecteren"
    Script-Module-SetCredentials -Credentials $Credential -Type $CredType
    $global:CredentialSource = $Credential
    $global:CredentialSourceUPN = $Credential.UserName
    If ($CredTarget -like "*@*") {
        $global:DomainSource = ($CredTarget.Replace($CredType + '-', '')).Split('@')[1]
    }
    ElseIf ($global:ServerSource -ne "Office 365") {
        $global:DomainSource = ($global:ServerSource).Split('.')[1, 2] -Join "."
    }
    Else {
        $global:DomainSource = Read-Host "- Please enter the mail domain which is accepted at $ServerSource (i.e. customer.com)"
    }
    # ==========
    # Finalizing
    # ==========
    $Task = "Please enter the source server"
    Set-Variable $global:VarHeaderName -Value ($global:ServerSource + " (" + $global:CredentialSourceUPN + ")")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}