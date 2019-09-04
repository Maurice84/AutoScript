Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-9-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # -----------------------------------------
    # Iterate through each line of the CSV-file
    # -----------------------------------------
    $Counter = 1
    $CSV = Import-CSV -Path "$File"
    ForEach ($CSVLine in $CSV) {
        $AutoScript = $CSVLine.'Date'
        $CSVExport = "exported by Maurice AutoScript (full overview)"
    }
    If ($AutoScript) {
        $Count = $CSV.Length
    }
    Else {
        $CSV = Import-CSV -Delimiter ";" -Path "$File"
        $CSVExport = "exported by Maurice AutoScript (customer overview)"
        $Count = 0
        ForEach ($CSVLine in $CSV) {
            $Soort = $CSVLine.'Is this a mailbox or user account? (Mailbox/User)'
            $Aanmaken = $CSVLine.'Does this account need to be active? Yes/No'
            If ($Aanmaken -eq "Yes" -AND $Soort -eq "User") {
                $CSVExport = "exported from Excel (edited by customer)"
                $Count++
            }
        }
    } 
    ForEach ($CSVLine in $CSV) {
        If ($AutoScript) {
            $Aanmaken = "Ja"
            $Soort = "Gebruiker"
            $GivenName = $CSVLine.'GivenName'
            $Surname = $CSVLine.'Surname'
            $Initials = $CSVLine.'Initials'
            $SamAccountName = $CSVLine.'SamAccountName'
            $UserPrincipalName = $CSVLine.'UserPrincipalName'
            $Description = $CSVLine.'Description'
            $Office = $CSVLine.'Office'
            $Title = $CSVLine.'Title'
            $Department = $CSVLine.'Department'
            $Company = $CSVLine.'Company'
            $HomePage = $CSVLine.'HomePage'
            $StreetAddress = $CSVLine.'StreetAddress'
            $POBox = $CSVLine.'P.O.Box'
            $City = $CSVLine.'City'
            $State = $CSVLine.'State'
            $PostalCode = $CSVLine.'PostalCode'
            $Country = $CSVLine.'Country'
            $OfficePhone = $CSVLine.'OfficePhone'
            $MobilePhone = $CSVLine.'MobilePhone'
            $HomePhone = $CSVLine.'HomePhone'
            $Fax = $CSVLine.'Fax'
            $IPPhone = $CSVLine.'IPPhone'
            $Pager = $CSVLine.'Pager'
            $ExtensionAttribute1 = $CSVLine.'ExtensionAttribute1'
        }
        Else {
            $GivenName = $CSVLine.'GivenName'
            $Surname = $CSVLine.'Surname'
            $Email = $CSVLine.'PrimarySmtpAddress'
            $UserPrincipalName = $CSVLine.'New UserPrincipalName'
            $Soort = $CSVLine.'Is this a mailbox or user account? (Mailbox/User)'
            $Aanmaken = $CSVLine.'Does this account need to be active? (Yes/No)'
        }
        If ($Aanmaken -eq "Yes" -AND $Soort -eq "User") {
            $Username = $UserPrincipalName.Split('@')[0]
            $UPN = $Username + "@" + $UPNSuffix
            Script-Module-SetHeaders -Name $Titel
            Write-Host "The selected CSV-file is" $CSVExport
            Write-Host ("Creating domain account " + $Counter + " of " + $Count + ":")
            Write-Host
            Write-Host "Domain account info:" -ForegroundColor Magenta
            If ($Surname) { $Name = $GivenName + " " + $Surname } Else { $Name = $GivenName } Write-Host -NoNewLine " - Name: "; Write-Host $Name -ForegroundColor Yellow
            If ($UPN) { Write-Host -NoNewLine " - UserPrincipalName: "; Write-Host ($Username + " (" + $UPNSuffix + ")") -ForegroundColor Yellow }
            If ($Email) { Write-Host -NoNewLine " - PrimarySmtpAddress: "; Write-Host $Email -ForegroundColor Yellow }
            If ($Description) { Write-Host -NoNewLine " - Description: "; Write-Host $Description -ForegroundColor Yellow }
            If ($Office) { Write-Host -NoNewLine " - Office: "; Write-Host $Office -ForegroundColor Yellow }
            If ($Title) { Write-Host -NoNewLine " - Title: "; Write-Host $Title -ForegroundColor Yellow }
            If ($Department) { Write-Host -NoNewLine " - Department: "; Write-Host $Department -ForegroundColor Yellow }
            If ($HomePage) { Write-Host -NoNewLine " - HomePage: "; Write-Host $HomePage -ForegroundColor Yellow }
            If ($StreetAddress) { Write-Host -NoNewLine " - StreetAddress: "; Write-Host $StreetAddress -ForegroundColor Yellow }
            If ($POBox) { Write-Host -NoNewLine " - P.O.Box: "; Write-Host $POBox -ForegroundColor Yellow }
            If ($City) { Write-Host -NoNewLine " - City: "; Write-Host $City -ForegroundColor Yellow }
            If ($State) { Write-Host -NoNewLine " - State: "; Write-Host $State -ForegroundColor Yellow }
            If ($PostalCode) { Write-Host -NoNewLine " - PostalCode: "; Write-Host $PostalCode -ForegroundColor Yellow }
            If ($OfficePhone) { Write-Host -NoNewLine " - OfficePhone: "; Write-Host $OfficePhone -ForegroundColor Yellow }
            If ($MobilePhone) { Write-Host -NoNewLine " - MobilePhone: "; Write-Host $MobilePhone -ForegroundColor Yellow }
            If ($HomePhone) { Write-Host -NoNewLine " - HomePhone: "; Write-Host $HomePhone -ForegroundColor Yellow }
            If ($Fax) { Write-Host -NoNewLine " - Fax: "; Write-Host $Fax -ForegroundColor Yellow }
            If ($IPPhone) { Write-Host -NoNewLine " - IPPhone: "; Write-Host $IPPhone -ForegroundColor Yellow }
            If ($Pager) { Write-Host -NoNewLine " - Pager: "; Write-Host $Pager -ForegroundColor Yellow }
            If ($ExtensionAttribute1) { Write-Host -NoNewLine " - ExtensionAttribute1: "; Write-Host $ExtensionAttribute1 -ForegroundColor Yellow }
            If ($ExtensionAttribute2) { Write-Host -NoNewLine " - ExtensionAttribute2: "; Write-Host $ExtensionAttribute2 -ForegroundColor Yellow }
            If ($ExtensionAttribute3) { Write-Host -NoNewLine " - ExtensionAttribute3: "; Write-Host $ExtensionAttribute3 -ForegroundColor Yellow }
            If ($ExtensionAttribute4) { Write-Host -NoNewLine " - ExtensionAttribute4: "; Write-Host $ExtensionAttribute4 -ForegroundColor Yellow }
            If ($ExtensionAttribute5) { Write-Host -NoNewLine " - ExtensionAttribute5: "; Write-Host $ExtensionAttribute5 -ForegroundColor Yellow }
            If ($ExtensionAttribute6) { Write-Host -NoNewLine " - ExtensionAttribute6: "; Write-Host $ExtensionAttribute6 -ForegroundColor Yellow }
            If ($ExtensionAttribute7) { Write-Host -NoNewLine " - ExtensionAttribute7: "; Write-Host $ExtensionAttribute7 -ForegroundColor Yellow }
            If ($ExtensionAttribute8) { Write-Host -NoNewLine " - ExtensionAttribute8: "; Write-Host $ExtensionAttribute8 -ForegroundColor Yellow }
            If ($ExtensionAttribute9) { Write-Host -NoNewLine " - ExtensionAttribute9: "; Write-Host $ExtensionAttribute9 -ForegroundColor Yellow }
            If ($ExtensionAttribute10) { Write-Host -NoNewLine " - ExtensionAttribute10: "; Write-Host $ExtensionAttribute10 -ForegroundColor Yellow }
            If ($ExtensionAttribute11) { Write-Host -NoNewLine " - ExtensionAttribute11: "; Write-Host $ExtensionAttribute11 -ForegroundColor Yellow }
            If ($ExtensionAttribute12) { Write-Host -NoNewLine " - ExtensionAttribute12: "; Write-Host $ExtensionAttribute12 -ForegroundColor Yellow }
            If ($ExtensionAttribute13) { Write-Host -NoNewLine " - ExtensionAttribute13: "; Write-Host $ExtensionAttribute13 -ForegroundColor Yellow }
            If ($ExtensionAttribute14) { Write-Host -NoNewLine " - ExtensionAttribute14: "; Write-Host $ExtensionAttribute14 -ForegroundColor Yellow }
            If ($ExtensionAttribute15) { Write-Host -NoNewLine " - ExtensionAttribute15: "; Write-Host $ExtensionAttribute15 -ForegroundColor Yellow }
            Write-Host
            If (!$AutoScript) { Write-Host "CSV info:" -ForegroundColor Magenta }
            If (!$AutoScript) { Write-Host -NoNewLine " - Create account according to the CSV-file? "; Write-Host $Aanmaken -ForegroundColor Green }
            If (!$AutoScript) { Write-Host -NoNewLine " - Mailbox or user according to CSV-file? "; Write-Host $Soort -ForegroundColor Green }
            Write-Host
            Write-Host "Creation info:" -ForegroundColor Magenta
            Write-Host -NoNewLine " - Location: "; Write-Host $global:Subtree -ForegroundColor Yellow
            Write-Host -NoNewLine " - Profile: "; Write-Host $Profiel -ForegroundColor Yellow
            If ($Prefix) { Write-Host -NoNewLine " - Prefix: "; Write-Host $Prefix -ForegroundColor Yellow }
            Write-Host -NoNewLine " - Language: "; Write-Host $Language -ForegroundColor Yellow
            If ($global:Array) { Write-Host -NoNewLine " - Groups: "; Write-Host ($global:Array.Name -join ", ") -ForegroundColor Yellow }
            Write-Host
            $Counter++
            Start-Sleep -Seconds 2
            Identity-ActiveDirectory-Create-Accounts-10-Start
        }
    }
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    $CSV = $null
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
}