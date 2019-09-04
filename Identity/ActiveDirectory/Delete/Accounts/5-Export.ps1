Function Identity-ActiveDirectory-Delete-Accounts-5-Export {
    # ============
    # Declarations
    # ============
    $Task = "Moving selected homefolder(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:AantalHomefolders -ne "n.a.") {
        $global:GroupAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "(Get-ADGroup -Filter * | Where-Object {`$`_.Name -like `"Dom*dmin*`"}).Name")
        )
        Write-Host "- Moving selected homefolder(s)..." -ForegroundColor Gray
        Write-Host
        ForEach ($Homefolder in $ArrayHomefolders) {
            $global:Account = ($Homefolder.Name).Replace("|", "-")
            $global:HomePath = $Homefolder.HomePath
            If ($HomePath -notlike "Geen homefolder*") {
                $global:HomePathRoot = $HomePath.Split('\')[0..($HomePath.Split('\').Length - 2)] -join "\"
                $global:Folder = $HomePath[($HomePathRoot.Length + 1)..$HomePath.Length] -join ""
                $global:Destination = $HomePathRoot + "\- Archief\" + $Account
                Write-Host (" - " + $Account + ": Checking presence of a centralized archive folder in the root of the homefolder share...")
                If (!(Get-Item $Destination -ErrorAction SilentlyContinue)) {
                    Write-Host -NoNewLine ("  " + $Account + ": Creating centralized archive folder: "); Write-Host -NoNewLine $global:Destination -ForegroundColor Yellow; Write-Host "..."
                    New-Item $Destination -Type Directory | Out-Null
                    If ($?) {
                        $global:Destination += "\" + $global:Folder
                    }
                    Else {
                        Write-Host (" - " + $global:Account + ": Centralized archive folder could not be created, please investigate!") -ForegroundColor Red
                    }
                }
                Else {
                    Write-Host -NoNewLine (" - " + $global:Account + ": Centralized archive folder found: "); Write-Host -NoNewLine $global:Destination -ForegroundColor Yellow; Write-Host "..."
                    $global:Destination += "\" + $global:Folder
                }
                Write-Host -NoNewLine (" - " + $global:Account + ": Moving homefolder to "); Write-Host -NoNewLine $global:Destination -ForegroundColor Yellow; Write-Host -NoNewLine "..."
                Move-Item $HomePath $global:Destination -Force -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                If (!$?) {
                    Write-Host
                    Write-Host (" - " + $global:Account + ": Homefolder could not be moves to the centralized archive folder, please investigate!") -ForegroundColor Magenta
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                }
                Write-Host
                # ---------------------------------
                # Modifying owner of the Homefolder
                # ---------------------------------
                If ((Get-ACL $global:Destination).Owner -notlike "*$global:GroupAdmins") {
                    ICACLS "$global:Destination" /setowner "$global:GroupAdmins" /T | Out-Null
                    If (!$?) {
                        Write-Host " - ERROR: An error occurred setting the owner permission of $global:GroupAdmins, please investigate!" -ForegroundColor Magenta
                        $Pause.Invoke()
                        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                    }
                }
                # ------------------------------
                # Adding Inheritance permissions
                # ------------------------------
                $global:ACL = (Get-Item $global:Destination).GetAccessControl('Access')
                $ACL.SetAccessRuleProtection($False, $False)
                Set-ACL -Path "$global:Destination" -ACLObject $ACL
                If (!$?) {
                    Write-Host (" - " + $global:Account + ": Inheritance could not be enabled, please investigate!") -ForegroundColor Magenta
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                }
                Write-Host (" - " + $global:Account + ": Homefolder successfully moved") -ForegroundColor Green
                Write-Host
                $Status = "Successful"
            }
        }
        Do {
            Write-Host -NoNewLine "  Would you like to delete the domain account(s) since the homefolders are successfully moved? Use "; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N): "
            [string]$InputChoice = Read-Host
            $InputKey = @("Y", "N") -contains $InputChoice
            If (!$InputKey) {
                Write-Host "Please use the letters above as input" -ForegroundColor Red
            }
        } Until ($InputKey)
        Switch ($InputChoice) {
            "N" {
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
    } else {
        $Status = "n.a."
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $Status
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}