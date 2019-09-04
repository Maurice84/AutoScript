Function Messaging-Exchange-Create-Mailboxes-UsingCSV-4-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # ------------------------------------------
    # Iterate through every line of the CSV-file
    # ------------------------------------------
    $CSV = Import-CSV -Path "$global:File"
    ForEach ($CSVLine in $CSV) {
        $Exchange20xx = $CSVLine.'Date'
    }
    If (!$Exchange20xx) {
        $CSV = Import-CSV -Delimiter "`t" -Path "$global:File"
    }
    $MailboxArray = @()
    ForEach ($CSVLine in $CSV) {
        If ($Exchange20xx) {
            $Email = $CSVLine.'PrimarySmtpAddress'
            $Username = $CSVLine.'SamAccountName'
        }
        Else {
            $Email = $CSVLine.'Primary E-mail'
            $Username = $CSVLine.'User Name'
        }
        $Account = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADUser -Filter {(EmailAddress -eq '$Email') -OR (SamAccountName -eq '$Username')} | Select-Object SamAccountName")
        )
        $SamAccountName = $Account.SamAccountName
        If ($SamAccountName) {
            $ExtraAddresses = New-Object System.Collections.Generic.List[System.Object]
            If ($Exchange20xx) { 
                $DisplayName = $CSVLine.'Name'
                $LegacyExchangeDN = $CSVLine.'LegacyExchangeDN'
                $LastLogonDate = $CSVLine.'LastLogonDate'
                $SizeMB = $CSVLine.'SizeMB'
                $ItemCount = $CSVLine.'ItemCount'
                $global:Language = $CSVLine.'Language'
                For ($Count = 2; $Count -le 99; $Count++) {
                    $Column = "EmailAddress" + $Count + ":"
                    If ($CSVLine.$Column -AND $CSVLine.$Column -notlike "*.local") {
                        $ExtraAddresses.Add($CSVLine.$Column)
                    }
                }
                $Properties = @{LastLogonDate = $LastLogonDate }
            }
            Else {
                $DisplayName = $CSVLine.'Display Name'
                $X400 = $CSVLine.'X400 Address'
                $LegacyExchangeDN = $CSVLine.'X500 Address'
                $SizeMB = $CSVLine.'Size (MB)'
                $ItemCount = $CSVLine.'Total Items'
                $global:Language = $CSVLine.'Language'
                ForEach ($ExtraAddress in $CSVLine."E-mail Addresses".Split(',')) {
                    If ($ExtraAddress -AND $ExtraAddress -notlike "*.local") {
                        $ExtraAddresses.Add($ExtraAddress)
                    }
                }
                $Properties = @{X400 = $X400 }
            }   
            $Properties += @{`
                    DisplayName      = $DisplayName; `
                    SamAccountName   = $SamAccountName; `
                    Email            = $Email; `
                    Language         = $global:Language; `
                    ExtraAddresses   = $ExtraAddresses; `
                    LegacyExchangeDN = $LegacyExchangeDN; `
                    SizeMB           = $SizeMB; `
                    ItemCount        = $ItemCount`
            
            }
            $Object = New-Object PSObject -Property $Properties
            $MailboxArray += $Object
        }
    }
    $Counter = 1
    ForEach ($Mailbox in $MailboxArray) {
        $DisplayName = $Mailbox.DisplayName
        $SamAccountName = $Mailbox.SamAccountName
        $Email = $Mailbox.Email
        $ExtraAddresses = $Mailbox.ExtraAddresses
        $global:Language = $Mailbox.Language
        If ($global:Language -eq "de-DE") {
            $global:Taalkeuze = "Duits"
        }
        If ($global:Language -eq "en-US") {
            $global:Taalkeuze = "Engels"
        }
        If ($global:Language -eq "nl-NL") {
            $global:Taalkeuze = "Nederlands"
        }
        $LastLogonDate = $Mailbox.LastLogonDate
        $X400 = $Mailbox.X400
        $LegacyExchangeDN = $Mailbox.LegacyExchangeDN
        $SizeMB = $Mailbox.SizeMB
        $ItemCount = $Mailbox.ItemCount
        # ---------------------------
        # Display the mailbox details
        # ---------------------------
        Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
        Write-Host "Creating mailbox" $Counter "of" $MailboxArray.Length ":"
        Write-Host
        Write-Host "Mailbox information:" -ForegroundColor Magenta
        If ($DisplayName) { Write-Host -NoNewLine " - Name: "; Write-Host $DisplayName -ForegroundColor Yellow }
        If ($SamAccountName) { Write-Host -NoNewLine " - SamAccountName: "; Write-Host $SamAccountName -ForegroundColor Yellow }
        If ($Email) { Write-Host -NoNewLine " - PrimarySmtpAddress: "; Write-Host $Email -ForegroundColor Yellow }
        If ($ExtraAddresses) { Write-Host -NoNewLine " - Extra mail address(es): "; Write-Host ($ExtraAddresses -join ", ") -ForegroundColor Yellow }
        If ($X400) { Write-Host -NoNewLine " - Legacy address X400: "; Write-Host "Available"-f Yellow }
        If ($LegacyExchangeDN) { Write-Host -NoNewLine " - Legacy address X500: "; Write-Host "Available"-f Yellow }
        If ($LastLogonDate) { Write-Host -NoNewLine " - LastLogonDate: "; Write-Host $LastLogonDate -ForegroundColor Yellow }
        If ($SizeMB) { Write-Host -NoNewLine " - Size: "; Write-Host ($SizeMB + " MB") -ForegroundColor Yellow }
        If ($ItemCount) { Write-Host -NoNewLine " - Item Count: "; Write-Host $ItemCount -ForegroundColor Yellow }
        Write-Host
        Write-Host "Creation information:" -ForegroundColor Magenta
        Write-Host -NoNewLine " - Language: "; Write-Host ($global:Taalkeuze + " (" + $global:Language + ")") -ForegroundColor Yellow
        Write-Host
        Start-Sleep -Seconds 1
        # -----------------------------------------------------------------
        # Go to function Messaging > Create Mailboxes to create the mailbox
        # -----------------------------------------------------------------
        Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Create-Mailboxes-*-Start" })
        Write-Host
        If ($Counter -ne $MailboxArray.Count) {
            $Counter++
            Write-Host -NoNewLine "Proceeding to the next mailbox in 2 seconds..." -ForegroundColor Green
            Start-Sleep -Seconds 2
        }
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}