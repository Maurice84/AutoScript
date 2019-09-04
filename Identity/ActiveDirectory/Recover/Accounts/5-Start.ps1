Function Identity-ActiveDirectory-Recover-Accounts-5-Start {
    # =========
    # Execution
    # =========
    If (!$global:AccountHerstelCSV) {
        ForEach ($Account in $global:Array) {
            $Task = "Recovering domain account " + $Account.Name
            Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
            If ($Account.DN -eq "Bestaat niet meer") {
                Write-Host
                Write-Host "  Caution: The Organizational Unit for this account does not exist anymore! Please select another Organizational Unit in the next menu:" -ForegroundColor Red
                $Pause.Invoke()
                Script-Index-OU
                $OU = $global:OUPath
            }
            Else {
                $OU = $Account.DistinguishedName
            }
            $ObjectGUID = $Account.GUID
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADObject -Filter {ObjectGUID -eq '$ObjectGUID'} -IncludeDeletedObjects | Restore-ADObject -TargetPath '$OU' -Confirm:`$false")
            )
            If ($?) {
                Write-Host -NoNewLine (" - " + $Account.Name) -ForegroundColor Yellow; Write-Host ": Recovered successfully"
            }
            Else {
                $Problem = $true
                Write-Host (" - " + $Account.Name + ": Could not recover domain account, please investigate!") -ForegroundColor Magenta
            }
        }
    }
    Else {
        $CSVFile = Import-CSV -Path "$global:File"
        ForEach ($CSVLine in $CSVFile) {
            $global:GivenName = $CSVLine.'GivenName'
            $global:SurName = $CSVLine.'Surname'
            $global:Initials = $CSVLine.'Initials'
            $global:SamAccountName = $CSVLine.'SamAccountName'
            $global:UPN = $CSVLine.'UserPrincipalName'
            $global:Description = $CSVLine.'Description'
            $global:Office = $CSVLine.'Office'
            $global:Title = $CSVLine.'Title'
            $global:Department = $CSVLine.'Department'
            $global:Company = $CSVLine.'Company'
            $global:HomePage = $CSVLine.'HomePage'
            $global:StreetAddress = $CSVLine.'StreetAddress'
            $global:POBox = $CSVLine.'P.O.Box'
            $global:City = $CSVLine.'City'
            $global:State = $CSVLine.'State'
            $global:PostalCode = $CSVLine.'PostalCode'
            $global:Country = $CSVLine.'Country'
            $global:OfficePhone = $CSVLine.'OfficePhone'
            $global:MobilePhone = $CSVLine.'MobilePhone'
            $global:HomePhone = $CSVLine.'HomePhone'
            $global:Fax = $CSVLine.'Fax'
            $global:IPPhone = $CSVLine.'IPPhone'
            $global:Pager = $CSVLine.'Pager'
            $global:OUPath = $CSVLine.'DistinguishedName'
            If ($global:Surname.Length -ne 0) {
                $global:Name = $global:GivenName + " " + $global:Surname
            }
            Else {
                $global:Name = $global:GivenName
            }
            $Task = "Recovering domain account " + $global:Name
            Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
            # -------------------------
            # Indexing group membership
            # -------------------------
            $global:Groups = @()
            $Counter = 1
            Do {
                $Group = "Group " + $Counter + ":"
                If ($CSVLine.$Group) {
                    $Object = New-Object PSObject
                    $Object | Add-Member -type NoteProperty -Name 'Name' -Value $CSVLine.$Group
                    $global:Groups += $Object
                }
                $Counter++
            } Until (!$CSVLine.$Group)
            If ($global:OUPath) {
                $Length = $global:OUPath.Split(",").Count
                $global:OUPath = $global:OUPath.Split(",")[1..$Length] -join ","
                # ------------------------------------------------------------------------------
                # Converting DistinguishedName (OU=..,OU=...,DC=...) to CanonicalName (\OU\OU\)
                # ------------------------------------------------------------------------------
                $global:Subtree = Script-Convert-DN-to-CN -Path $global:OUPath
            }
            $CheckOU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Try {Get-ADOrganizationalUnit '$OUPath' -ErrorAction SilentlyContinue | Select-Object DistinguishedName | Out-Null} Catch {`$False}")
            )
            If ($CheckOU -eq $False) {
                Write-Host
                Write-Host "  Caution: The Organizational Unit for this account does not exist anymore! Please select another Organizational Unit in the next menu:" -ForegroundColor Magenta
                $Pause.Invoke()
                Script-Index-OU
            }
            # ----------------------------------------------------------------------
            # Go to function ActiveDirectory > Create Accounts to create the account
            # ----------------------------------------------------------------------
            Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "*Create-Accounts-*-Start" })
        }
    }
    Write-Host
    If (!$Problem) {
        Write-Host "  Domain account(s) successfully recovered. Please be advised that groups that have been deleted after the removal of the" -ForegroundColor Yellow
        Write-Host "  domain account are not recovered. Check the CSV-file if this information is needed. Also note homefolder data," -ForegroundColor Yellow
        Write-Host "  this should manually be moved to the correct homefolder (check the AD Object for the homefolder path)." -ForegroundColor Yellow
        Write-Host
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}