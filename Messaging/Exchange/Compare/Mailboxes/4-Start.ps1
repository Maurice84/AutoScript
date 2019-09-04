Function Messaging-Exchange-Compare-Mailboxes-4-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    # ------------------------------------------
    # Iterate through every line in the CSV-file
    # ------------------------------------------
    $CSV = Import-CSV -Path "$global:File"
    ForEach ($CSVLine in $CSV) {
        $Exchange20xx = $CSVLine.'Date'
    }
    $CSVArray = @()
    $Objects = @()
    ForEach ($CSVLine in $CSV) {
        If ($Exchange20xx) {
            $Email = $CSVLine.'PrimarySmtpAddress'
            $Username = $CSVLine.'SamAccountName'
        }
        Else {
            $Email = $CSVLine.'Primary E-mail'
            $Username = $CSVLine.'User Name'
        }
        $ExtraAddresses = New-Object System.Collections.Generic.List[System.Object]
        If ($Exchange20xx) { 
            $SizeMB = $CSVLine.'SizeMB'
            $ItemCount = $CSVLine.'ItemCount'
            For ($Count = 2; $Count -le 99; $Count++) {
                $Column = "EmailAddress" + $Count + ":"
                If ($CSVLine.$Column -AND $CSVLine.$Column -notlike "*.local") {
                    $ExtraAddresses.Add($CSVLine.$Column)
                }
            }
        }
        Else {
            $SizeMB = $CSVLine.'Size (MB)'
            $ItemCount = $CSVLine.'Total Items'
            ForEach ($ExtraAddress in $CSVLine."E-mail Addresses".Split(',')) {
                If ($ExtraAddress -AND $ExtraAddress -notlike "*.local") {
                    $ExtraAddresses.Add($ExtraAddress)
                }
            }
        }   
        $Properties = @{Email = $Email; ExtraAddresses = $ExtraAddresses; SizeMB = $SizeMB; ItemCount = $ItemCount }
        $Object = New-Object PSObject -Property $Properties
        $CSVArray += $Object
    }
    $Mailboxes = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -WarningAction SilentlyContinue"))
    $Mailboxes = $Mailboxes | Where-Object { $_.DisplayName -notlike "Discovery*" -AND $_.DisplayName -ne $null } | Select-Object Name, Alias, PrimarySMTPAddress, EmailAddresses | Sort-Object PrimarySMTPAddress
    ForEach ($Mailbox in $Mailboxes) {
        $EmailAddresses = ($Mailbox | Select-Object @{Name = "EmailAddresses"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress } } } | Select-Object EmailAddresses).EmailAddresses
        ForEach ($EmailAddress in $EmailAddresses) {
            If ($CSVArray.Email -eq $EmailAddress -OR $CSVArray.ExtraAddresses -contains $EmailAddress) {
                $CSVMailbox = $CSVArray | Where-Object { $_.Email -eq $EmailAddress -OR $_.ExtraAddresses -contains $EmailAddress }
                $ObjectStat = Get-MailboxStatistics -Identity $Mailbox.Alias -WarningAction SilentlyContinue
                $Property = New-Object PSObject
                $Property | Add-Member -type NoteProperty -Name 'Name' -Value $Mailbox.Name
                $Property | Add-Member -type NoteProperty -Name 'Email' -Value $Mailbox.PrimarySMTPAddress
                $Property | Add-Member -type NoteProperty -Name 'ItemsOld' -Value $CSVMailbox.ItemCount
                $Property | Add-Member -type NoteProperty -Name 'ItemsNew' -Value $ObjectStat.ItemCount
                $Property | Add-Member -type NoteProperty -Name 'ItemsDiff' -Value ($ObjectStat.ItemCount - $CSVMailbox.ItemCount)
                $Property | Add-Member -type NoteProperty -Name 'SizeOld' -Value $CSVMailbox.SizeMB
                $Property | Add-Member -type NoteProperty -Name 'SizeNew' -Value $ObjectStat.TotalItemSize.Value.ToMB()
                $Property | Add-Member -type NoteProperty -Name 'SizeDiff' -Value ($ObjectStat.TotalItemSize.Value.ToMB() - $CSVMailbox.SizeMB)
                $Objects += $Property
                BREAK
            }
        }
    }
    If (!$Objects) {
        Write-Host "No mailbox(es) detected to compare with mailbox(es) in the CSV-file." -ForegroundColor Red
        Write-Host
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    Script-Module-SetHeaders -Name $Titel
    If ($Objects) {
        ForEach ($Object in $Objects) {
            If ($Object.ItemsDiff -lt 0 -AND $Object.SizeDiff -lt 0) {
                Write-Host $Object.Name "has" ($Object.ItemsOld - $Object.ItemsNew) "items less, therefore the mailbox is" ($Object.SizeOld - $Object.SizeNew) "MB smaller!" -ForegroundColor Red
            }
        }
    }
    Else {
        Write-Host "The destination mailbox(es) are bigger and contain more items compared to the mailbox(es) in the CSV-file."
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}