Function Software-SyncBackPro-Create-Profiles-ForAccounts-5-Start {
    # ============
    # Declarations
    # ============
    [int]$Counter = 1
    [int]$Count = 0
    $CSV = Import-CSV -Delimiter ";" -Path "$global:FileCSV"
    # =========
    # Execution
    # =========
    # -------------------------------------------------
    # Check if the CSV-file has been edited by customer
    # -------------------------------------------------
    ForEach ($CSVLine in $CSV) {
        $Type = $CSVLine.'Is this a mailbox or user account? (Mailbox/User)'
        $Active = $CSVLine.'Does this account need to be active? Yes/No'
        If ($Active -eq "Yes" -AND $Type -eq "User") {
            $CSVExport = "exported from Excel (edited by customer)"
            $Count++
        }
    }
    # -----------------------------------------
    # Iterate through each line of the CSV-file
    # -----------------------------------------
    ForEach ($CSVLine in $CSV) {
        $GivenName = $CSVLine.'GivenName'
        $Surname = $CSVLine.'Surname'
        $SamAccountName = $CSVLine.'UserPrincipalName'
        $UserPrincipalName = $CSVLine.'New UserPrincipalName'
        $Type = $CSVLine.'Is this a mailbox or user account? (Mailbox/User)'
        $Active = $CSVLine.'Does this account need to be active? Yes/No'
        If ($Active -eq "Yes" -AND $Type -eq "User") {
            Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
            $Username = $UserPrincipalName.Split('@')[0]
            Write-Host "Selected CSV-file is" $CSVExport
            Write-Host "Creating SyncBackPro profile" $Counter "of" $Count ":"
            Write-Host
            Write-Host "Domain account info:" -ForegroundColor Magenta
            If ($Surname) {
                $Name = $GivenName + " " + $Surname
            }
            Else {
                $Name = $GivenName
            } Write-Host -NoNewLine " - Name: "; Write-Host $Name -ForegroundColor Yellow
            If ($SamAccountName) { Write-Host -NoNewLine " - UserPrincipalName: "; Write-Host $SamAccountName -ForegroundColor Yellow }
            If ($Username) { Write-Host -NoNewLine " - New UserPrincipalName: "; Write-Host $Username -ForegroundColor Yellow }
            Write-Host
            Write-Host "CSV info:" -ForegroundColor Magenta
            Write-Host -NoNewLine " - Create account according to the CSV-file? "; Write-Host $Active -ForegroundColor Green
            Write-Host -NoNewLine " - Mailbox or user according to CSV-file? "; Write-Host $Type -ForegroundColor Green
            Write-Host
            Write-Host "Creation info:" -ForegroundColor Magenta
            Write-Host -NoNewLine " - Selected SyncBackPro template: "; Write-Host $FileTemplate.Split("\")[-1] -ForegroundColor Yellow
            Write-Host
            $Counter++
            # ---------------------------------------------------------------------------------------
            Write-Host -NoNewLine " Step 1/2 - Copying template" -ForegroundColor Magenta; Write-Host
            # ---------------------------------------------------------------------------------------
            $Path = $FileTemplate.Replace('TEMPLATE', $Name)
            If (!(Test-Path $Path)) {
                Copy-Item $FileTemplate $Path
                If ($?) {
                    Write-Host -NoNewLine " - OK: Template successfully copied to "; Write-Host $Path.Split("\")[-1] -ForegroundColor Yellow
                }
                Else {
                    Write-Host " - ERROR: An error occurred copying the template, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: Profile already exists as" $Path.Split("\")[-1] -ForegroundColor Gray
            }
            Write-Host
            # ---------------------------------------------------------------------------------------------------------------
            Write-Host -NoNewLine " Step 2/2 - Modifying the data in the copied profile" -ForegroundColor Magenta; Write-Host
            # ---------------------------------------------------------------------------------------------------------------
            (Get-Content $Path) | ForEach-Object { $_ -replace "TEMPLATE-OLD", $SamAccountName } | Set-Content $Path
            [string]$Source = (Get-Content $Path | Select-String -pattern 'Source=')[-1]
            If ($?) {
                Write-Host -NoNewLine " - OK: Source successfully modified to "; Write-Host $Source.Split('=')[1] -ForegroundColor Yellow
            }
            Else {
                Write-Host " - ERROR: An error occurred modifying the source TEMPLATE-NEW variable, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
            (Get-Content $Path) | ForEach-Object { $_ -replace "TEMPLATE-NEW", $Username } | Set-Content $Path
            [string]$Destination = (Get-Content $Path | Select-String -pattern 'Destination=')[-1]
            If ($?) {
                Write-Host -NoNewLine " - OK: Destination successfully modified to "; Write-Host $Destination.Split('=')[1] -ForegroundColor Yellow
            }
            Else {
                Write-Host " - ERROR: An error occurred modifying the destination TEMPLATE-NEW variable, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
    }
    # ==========
    # Finalizing
    # ==========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}