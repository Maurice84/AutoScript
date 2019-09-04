Function Messaging-Exchange-Move-Mailboxes-2-InputTarget {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Do {
        $Choice = Read-Host "- Is the target Office 365 (Exchange Online)? (Y/N)"
        $Input = @("Y","N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {  
            $global:ServerDestination = "Office 365"
        }
        "N" {
            Do {
                $Choice = Read-Host "- is the target an Exchange server located in this network? (Y/N)"
                $Input = @("Y","N") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
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
                        $global:ServerDestination = ($GetExternalUrl.ExternalUrl).Host
                    } Else {
                        $global:ServerDestination = $CheckHybrid.OnPremisesSmartHost
                    }
                }
                "N" {
                    $global:ServerDestination = Read-Host "- Please enter the external DNS name of an Exchange server as target (i.e. mail.customer.nl)"
                }
            }
        }
    }
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Please enter the target server: " -ForegroundColor Gray; Write-Host $global:ServerDestination -ForegroundColor Yellow
    $Task = "Select an Administrator credential to connect to the source"
    If ($ServerDestination -eq "Office 365") {
        $CredType = "Office365"
    } Else {
        $CredType = "OnPremise"
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $CredType -Type "Credentials" -Functie "Selecteren"
    Script-Module-SetCredentials -Credentials $Credential -Type $CredType
    $global:CredentialDestination = $Credential
    $global:CredentialDestinationUPN = $Credential.UserName
    If ($CredTarget -like "*@*") {
        $global:DomainDestination = ($CredTarget.Replace($CredType + '-','')).Split('@')[1]
    } ElseIf ($global:ServerDestination -ne "Office 365") {
        $global:DomainDestination = ($global:ServerDestination).Split('.')[1,2] -Join "."
    } Else {
        $global:DomainDestination = Read-Host "- Please enter the mail domain which is accepted at $ServerDestination (i.e. customer.com)"
    }
    $global:DomainDestination = ($CredTarget.Replace($CredType + '-','')).Split('@')[1]
    # ==========
    # Finalizing
    # ==========
    $Task = "Please enter the target server"
    Set-Variable $global:VarHeaderName -Value ($global:ServerDestination + " (" + $global:CredentialDestinationUPN + ")")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}