Function Identity-ActiveDirectory-Overview-Accounts-2-Export {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If ($Titel -like "*ActiveDirectory*Overview*" -AND !$AzureADConnected) {
        Write-Host "Indexing groups and membership..." -ForegroundColor Magenta
        $ArrayAccounts = $global:Array
        # ---------------------------------------------------------------------------------------------------------------------------------------------
        # Checking if 1 of more accounts need to be exported. If it's only 1, then we use groupnames. If there's more, we use group selection using 'X'
        # ---------------------------------------------------------------------------------------------------------------------------------------------
        If ($ArrayAccounts.Count) {
            $ArrayTotaal = $ArrayAccounts.Count
        }
        Else {
            $ArrayTotaal = 1
        }
        If ($ArrayTotaal -ge 2) {
            $ExportType = "Groepmarkeringen"
        }
        Else {
            $ExportType = "Groepnamen"
        }
    }
    Else {
        If ($AzureADConnected) {
            $ArrayAccounts = $global:Array
        }
        Else {
            $ExportType = "Groepnamen"
        }
    }
    If ($ExportType -eq "Groepmarkeringen") {
        Write-Host -NoNewLine
        # -------------------------------------------------------
        # Group selection: Filter groups using prefix if detected
        # -------------------------------------------------------
        $Groups = @()
        $CounterAccounts = 0
        #If ($EasyCloud) {
        #    ForEach ($Account in $ArrayAccounts) {
        #        $CounterAccounts++
        #        $Prefix = $Account.ExtensionAttribute1
        #        Script-Module-SetHeaders -Name $Titel
        #        $Subtitel2 = {
        #            Write-Host -NoNewLine (" - Bezig met inventariseren van de groeplidmaatschap voor Easy-Cloud account " + $CounterAccounts + "/" + $ArrayAccounts.Count + "...")
        #        }
        #        $Subtitel.Invoke()
        #        $Subtitel2.Invoke()
        #        If (!$Groups) {
        #            $Groups += Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADGroup -Filter * | Where-Object {`$`_.Name -like `"$Prefix-*`" -OR `$`_.Name -like `"*-$Prefix-*`"}"))
        #        }
        #    }
        #    If ($Groups) {
        #        $Groups = $Groups | Select-Object Name | Sort-Object Name -Unique
        #    }
        #} Else {
        Write-Host -NoNewLine " - Indexing all groups..."
        $GetGroups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADGroup -Filter * | Select-Object Name | Sort-Object Name"))
        #}
        # -------------------------------------------
        # Group selection: Detecting group membership
        # -------------------------------------------
        $Groups = @()
        ForEach ($Group in $GetGroups) {
            # ----------------------------------------------------------------------------------------------------------------------------------------------
            # If there are special characters found we need to replace them like apostroph ' because this will fail the indexing of the members.
            # To fix this we need to convert the character with a temporary text like <char39>. During indexing we convert it back with [char]39.ToString()
            # ----------------------------------------------------------------------------------------------------------------------------------------------
            if ($Group.Name -like "*'*") {
                $GroupName = ($Group.Name).Replace("'", "<char39>")
            }
            else {
                $GroupName = $Group.Name
            }
            $Groups += "[" + $GroupName + "];"
        }
        if (Test-Path "C:\_TEMP_GroupsAndMembers.csv") {
            Write-Host " Temporary saved Group membership detected" -ForegroundColor Cyan
            Do {
                Write-Host -NoNewLine "   > Would you like to use this? (Y/N): "
                $Choice = Read-Host
                $Input = @("Y"; "N") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "Y" {
                    $GroupArray = Import-Csv -Path "C:\_TEMP_GroupsAndMembers.csv" -Delimiter ";"
                    $Skip = $true
                }
            }
        }
        if (!$Skip) {
            $GroupArray = @()
            $GroupArray = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "`$Groups=@();
                ForEach(`$Temp in '$Groups'.Split(';')) {
                    `$Temp = ((`$Temp).Replace(' [','')).Replace('] ','');
                    `$Temp = ((`$Temp).Replace('[','')).Replace(']','');
                    `$Temp = (`$Temp).Replace('<char39>',([char]39).ToString());
                    `$Groups += `$Temp;
                };
                `$Titel = '$Titel';
                `$Subtitel = [scriptblock]::Create('$Subtitel');
                `$Subtitel2 = [scriptblock]::Create('$Subtitel2');
                `$CounterGroups=0;
                `$GroupArray=@();
                ForEach (`$Group in `$Groups) {
                    `$CounterGroups++;
                    Clear;
                    Write-Host (`"=`" * `$Titel.Length) -ForegroundColor Gray;
                    Write-Host `$Titel -ForegroundColor Gray;
                    Write-Host (`"=`" * `$Titel.Length) -ForegroundColor Gray;
                    Write-Host 
                    `$Subtitel3 = {
                        Write-Host;
                        Write-Host -NoNewLine (`" - Indexing group membership `" + `$CounterGroups + `"/`" + `$Groups.Count + `"...`");
                    };
                    `$Subtitel.Invoke();
                    `$Subtitel2.Invoke();
                    `$Subtitel3.Invoke();
                    `$GroupName = `$Group | %{`$`_.Name};
                    `$Members = Try {Get-ADGroupMember -Identity `"`$Group`" -Recursive -ErrorAction SilentlyContinue | Where-Object {`$`_.ObjectClass -eq `"user`"} | Select-Object Name} Catch {`$false};
                    If (`$Members -AND `$Members -ne `$false) {
                        `$GroupProperties = @{Name=`$Group; Members=`$Members | %{`$`_.Name}}
                        `$GroupObject = New-Object PSObject -Property `$GroupProperties
                        `$GroupArray += `$GroupObject
                    }
                };`$GroupArray")
            )
            $GroupArray | Select-Object Name, @{Name = "Members"; Expression = { $_.Members -join "," } } | Export-Csv "C:\_TEMP_GroupsAndMembers.csv" -Delimiter ";" -NoTypeInfo
        }
    }
    Write-Host
    # -----------------------------------
    # Declare the header for the CSV-file
    # -----------------------------------
    If ($ExportMigratie) {
        $ExportHeader = '"GivenName";"Surname";"UserPrincipalName";"PrimarySmtpAddress";"LastLogonDate";"New UserPrincipalName";"Comments engineer";'
        $ExportHeader += '"Is this a mailbox or user account? (Mailbox/User)";"Does this account need to be active? (Yes/No)";'
        $ExportHeader += '"If not, should we copy the data to an Archive folder? (Yes/No)";"External login allowed? Yes/No";'
        $ExportHeader += '"Comments customer (please add extra info based on comments engineer)"'
    }
    Else {
        $ExportHeader = '"Date"'
    }
    If ($ExportType -OR $AzureADConnected) {
        If ((($ArrayAccounts | % { $_.Name }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Name"' }
        If ((($ArrayAccounts | % { $_.GivenName }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"GivenName"' }
        If ((($ArrayAccounts | % { $_.Surname }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Surname"' }
        If ((($ArrayAccounts | % { $_.Initials }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Initials"' }
        If ((($ArrayAccounts | % { $_.Description }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Description"' }
        If ((($ArrayAccounts | % { $_.SamAccountName }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"SamAccountName"' }
        If ((($ArrayAccounts | % { $_.UserPrincipalName }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"UserPrincipalName"' }
        If ((($ArrayAccounts | % { $_.DN }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"OrganizationalUnit"' }
        If ($ArrayAccounts | % { $_.LastLogonDate }) { $ExportHeader += ';"LastLogonDate"' }
        If ((($ArrayAccounts | % { $_.Office365Guid }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Office365Guid"' }
        If ((($ArrayAccounts | % { $_.License }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Office365License"' }
        If ((($ArrayAccounts | % { $_.Email }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Email"' }
        If ((($ArrayAccounts | % { $_.Office }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Office"' }
        If ((($ArrayAccounts | % { $_.Title }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Title"' }
        If ((($ArrayAccounts | % { $_.Department }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Department"' }
        If ((($ArrayAccounts | % { $_.Company }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Company"' }
        If ((($ArrayAccounts | % { $_.HomePage }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"HomePage"' }
        If ((($ArrayAccounts | % { $_.StreetAddress }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"StreetAddress"' }
        If ((($ArrayAccounts | % { $_.POBox }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"P.O.Box"' }
        If ((($ArrayAccounts | % { $_.City }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"City"' }
        If ((($ArrayAccounts | % { $_.State }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"State"' }
        If ((($ArrayAccounts | % { $_.PostalCode }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"PostalCode"' }
        If ((($ArrayAccounts | % { $_.Country }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Country"' }
        If ((($ArrayAccounts | % { $_.OfficePhone }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"OfficePhone"' }
        If ((($ArrayAccounts | % { $_.MobilePhone }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"MobilePhone"' }
        If ((($ArrayAccounts | % { $_.HomePhone }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"HomePhone"' }
        If ((($ArrayAccounts | % { $_.Fax }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"Fax"' }
        If ((($ArrayAccounts | % { $_.HomePath }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"HomePath"' }
        If ((($ArrayAccounts | % { $_.HomeDrive }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"HomeDrive"' }
        If ((($ArrayAccounts | % { $_.RDPPath }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"RDPPath"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute1 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute1"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute2 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute2"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute3 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute3"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute4 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute4"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute5 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute5"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute6 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute6"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute7 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute7"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute8 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute8"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute9 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute9"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute10 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute10"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute11 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute11"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute12 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute12"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute13 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute13"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute14 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute14"' }
        If ((($ArrayAccounts | % { $_.ExtensionAttribute15 }) | Measure-Object -Maximum -Property Length).Maximum -gt 0) { $ExportHeader += ';"ExtensionAttribute15"' }
    }
    # ---------------------------------------------------------------------------------------------------------
    # Remove the file if it's present (Add-Content does not overwrite) and add the header to the empty CSV-file
    # ---------------------------------------------------------------------------------------------------------
    If ($Titel -like "*ActiveDirectory*Overview*") {
        If ($OUFilter -eq "*") {
            If ($AzureADConnected) {
                $Domain = $AzureADConnect.TenantDomain
            }
            Else {
                $Domain = (Get-WmiObject Win32_ComputerSystem).Domain
            }
            $FileCSVU = "C:\Accounts-" + $Domain + "_" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + ".csv"
        }
        Else {
            $FileCSVU = "C:\Accounts-" + $OUFilter + "_" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + ".csv"
        }
        Remove-Item ($FileCSVU) -ErrorAction SilentlyContinue
        Add-Content ($FileCSVU) $ExportHeader
        Write-Host "Exporting to CSV-file..." -ForegroundColor Magenta
    }
    # ----------------------------------------------------------------------------------
    # Iterate through each account with the current timestamp and add it to the CSV-file
    # ----------------------------------------------------------------------------------
    $Counter = 1
    ForEach ($Account in $ArrayAccounts) {
        If ($ExportType -eq "Groepnamen") {
            $SamAccountName = $Account.SamAccountName
            $AccountGroups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADPrincipalGroupMembership '$SamAccountName' | Select-Object Name | Sort-Object Name).Name"))
        }
        If ($Titel -like "*ActiveDirectory*Overview*") {
            Write-Host -NoNewLine (" - Exporting domain account $Counter of " + $ArrayAccounts.Count + ": "); Write-Host -NoNewLine $Account.Name -ForegroundColor Yellow; Write-Host "..."
        }
        If ($Titel -like "*ActiveDirectory*Delete*") {
            $ExportHeader = '"Date"'
            $SamAccountName = $Account.SamAccountName
            $AccountDN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADUser -Filter {SamAccountName -eq '$SamAccountName'} | Select-Object DistinguishedName).DistinguishedName"))
            $ADSI = [ADSI]('LDAP://{0}' -f $AccountDN)
            $CheckHomePath = Try { $ADSI.InvokeGet('TerminalServicesHomeDirectory').Length } Catch { 0 }
            If ($CheckHomePath -gt 1) {
                $HomePath = $ADSI.InvokeGet('TerminalServicesHomeDirectory')
                $HomePathRoot = $HomePath.Split('\')[0..($HomePath.Split('\').Length - 2)] -join "\"
                $Folder = $HomePath[($HomePathRoot.Length + 1)..$HomePath.Length] -join ""
                $Destination = $HomePathRoot + "\- Archive\" + ($Account.Name).Replace("|", "-")
                $FileCSVU = $Destination + "\" + $Account.SamAccountName + ".csv"
            }
            Else {
                $Destination = "C:\"
                $FileCSVU = $Destination + $Account.SamAccountName + ".csv"
            }
            Write-Host -NoNewLine " - "; Write-Host -NoNewLine $Account.Name -ForegroundColor Yellow; Write-Host -NoNewLine ": Export to CSV\Excel-file "
            Remove-Item ($FileCSVU) -ErrorAction SilentlyContinue
            Add-Content ($FileCSVU) $ExportHeader
        }
        If ($ExportMigratie) {
            $ExportValues = '"'`
                + $Account.GivenName + '";"'`
                + $Account.Surname + '";"'`
                + $Account.SamAccountName + '";"'`
                + $Account.Email + '";"'`
                + $Account.LastLogonDate + '"'`
                + ';"";"";"";"";"";"";""'
        }
        Else {
            $Groups = @()
            $LastCheck = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            If ($Account.LastLogonDate -like "*2000*") {
                $LastLogonDate = "Nog nooit aangemeld"
            }
            Else {
                $LastLogonDate = Get-Date $Account.LastLogonDate -Format "yyyy-MM-dd HH:mm:ss"
            }
            $ExportValues = '"' + $LastCheck + '"'
            If ($ExportHeader -like "*Name*") { $ExportValues += ';"' + $Account.Name + '"' }
            If ($ExportHeader -like "*GivenName*") { $ExportValues += ';"' + $Account.GivenName + '"' }
            If ($ExportHeader -like "*Surname*") { $ExportValues += ';"' + $Account.Surname + '"' }
            If ($ExportHeader -like "*Initials*") { $ExportValues += ';"' + $Account.Initials + '"' }
            If ($ExportHeader -like "*Description*") { $ExportValues += ';"' + $Account.Description + '"' }
            If ($ExportHeader -like "*SamAccountName*") { $ExportValues += ';"' + $Account.SamAccountName + '"' }
            If ($ExportHeader -like "*UserPrincipalName*") { $ExportValues += ';"' + $Account.UserPrincipalName + '"' }
            If ($ExportHeader -like "*OrganizationalUnit*") { $ExportValues += ';"' + $Account.DN + '"' }
            If ($ExportHeader -like "*LastLogonDate*") { $ExportValues += ';"' + $LastLogonDate + '"' }
            If ($ExportHeader -like "*Office365Guid*") { $ExportValues += ';"' + $Account.Office365Guid + '"' }
            If ($ExportHeader -like "*Office365License*") { $ExportValues += ';"' + $Account.License + '"' }
            If ($ExportHeader -like "*Email*") { $ExportValues += ';"' + $Account.Email + '"' }
            If ($ExportHeader -like "*Office*") { $ExportValues += ';"' + $Account.Office + '"' }
            If ($ExportHeader -like "*Title*") { $ExportValues += ';"' + $Account.Title + '"' }
            If ($ExportHeader -like "*Department*") { $ExportValues += ';"' + $Account.Department + '"' }
            If ($ExportHeader -like "*Company*") { $ExportValues += ';"' + $Account.Company + '"' }
            If ($ExportHeader -like "*HomePage*") { $ExportValues += ';"' + $Account.HomePage + '"' }
            If ($ExportHeader -like "*StreetAddress*") { $ExportValues += ';"' + $Account.StreetAddress + '"' }
            If ($ExportHeader -like "*P.O.Box*") { $ExportValues += ';"' + $Account.POBox + '"' }
            If ($ExportHeader -like "*City*") { $ExportValues += ';"' + $Account.City + '"' }
            If ($ExportHeader -like "*State*") { $ExportValues += ';"' + $Account.State + '"' }
            If ($ExportHeader -like "*PostalCode*") { $ExportValues += ';"' + $Account.PostalCode + '"' }
            If ($ExportHeader -like "*Country*") { $ExportValues += ';"' + $Account.Country + '"' }
            If ($ExportHeader -like "*OfficePhone*") { $ExportValues += ';"' + $Account.OfficePhone + '"' }
            If ($ExportHeader -like "*MobilePhone*") { $ExportValues += ';"' + $Account.MobilePhone + '"' }
            If ($ExportHeader -like "*HomePhone*") { $ExportValues += ';"' + $Account.HomePhone + '"' }
            If ($ExportHeader -like "*Fax*") { $ExportValues += ';"' + $Account.Fax + '"' }
            If ($ExportHeader -like "*HomePath*") { $ExportValues += ';"' + $Account.HomePath + '"' }
            If ($ExportHeader -like "*HomeDrive*") { $ExportValues += ';"' + $Account.HomeDrive + '"' }
            If ($ExportHeader -like "*RDPPath*") { $ExportValues += ';"' + $Account.RDPPath + '"' }
            If ($ExportHeader -like "*ExtensionAttribute1*") { $ExportValues += ';"' + $Account.ExtensionAttribute1 + '"' }
        }
        If ($ExportType -eq "Groepmarkeringen") {
            $AccountName = $Account.Name
            $MatchedGroups = ($GroupArray | Where-Object { $_.Members -like "*$AccountName*" } | Select-Object Name).Name
            ForEach ($Group in ($GroupArray | Select-Object Name).Name) {
                If ($MatchedGroups -contains $Group) {
                    $ExportValues = $ExportValues + ';"X"'
                }
                Else {
                    $ExportValues = $ExportValues + ';" "'
                }
            }
        }
        Else {
            If (!$AzureADConnected) {
                ForEach ($Group in $AccountGroups) {
                    $ExportValues = $ExportValues + ';"' + $Group + '"'
                }
            }
        }
        # -----------------------------------------------------------------------------------
        # Here we remove non-readable characters (like carriage return (10) and linefeed (13)
        # -----------------------------------------------------------------------------------
        $ExportValues = $ExportValues.Replace(([char]10).ToString(), ' ')
        $ExportValues = $ExportValues.Replace(([char]13).ToString(), ' ')
        Add-Content ($FileCSVU) $ExportValues
        If (Test-Path $FileCSVU) {
            # -----------------------------------------------------------------
            # Modifying the column names in the header with the detected groups
            # -----------------------------------------------------------------
            If ($ExportType -eq "Groepnamen") {
                $GroupColumns = $null
                For ($CounterGroupColumns = 1; $CounterGroupColumns -le $AccountGroups.Count; $CounterGroupColumns++) {
                    $GroupColumns = $GroupColumns + ';"Group ' + $CounterGroupColumns + '"'
                }
                $ExportHeader += $GroupColumns
                $FileContent = Get-Content $FileCSVU
                $FileContent[0] = $ExportHeader
                $FileContent | Out-File $FileCSVU
            }
        }
        Else {
            Write-Host " - ERROR: An error occurred adding the domain account(s) to CSV-file, please investigate!" -ForegroundColor Red
        }
        $Counter++
    }
    If ($ExportType -eq "Groepmarkeringen") {
        $GroupText = ';"Group: '
        ForEach ($Group in $GroupArray) {
            $ExportHeader += $GroupText + $Group.Name + '"'
        }
        $FileContent = Get-Content $FileCSVU
        $FileContent[0] = $ExportHeader
        $FileContent | Out-File $FileCSVU
        Remove-Item -Path "C:\_TEMP_GroupsAndMembers.csv" -Force -Confirm:$False
    }
    If (Test-Path $FileCSVU -AND $Counter -eq $ArrayAccounts.Count) {
        Write-Host -NoNewLine " - OK: Successfully exported the domain account(s) to CSV-file: "; Write-Host $FileCSVU -ForegroundColor Yellow
    }
    # ----------------------------------
    # Convert the CSV-file to Excel-file
    # ----------------------------------
    Script-Convert-CSV-to-Excel -File $FileCSVU -Category "Accounts" -Silent $true
    $FileExcel = $FileCSVU.Substring(0, $FileCSVU.Length - 4) + ".xlsx"
    If (Test-Path $FileExcel) {
        Write-Host -NoNewLine " - OK: Successfully converted the CSV-file to Excel-file:"; Write-Host $FileExcel.Split('\')[-1] -ForegroundColor Yellow
    }
    Else {
        Write-Host " - ERROR: Could not convert the CSV-file to Excel-file, please investigate!" -ForegroundColor Red
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    If ($Titel -like "*ActiveDirectory*Overview*") {
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
}