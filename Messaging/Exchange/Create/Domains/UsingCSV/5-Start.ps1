Function Messaging-Exchange-Create-Domains-UsingCSV-5-Start {
    # =========
    # Execution
    # =========
    # ------------------------------------------
    # Iterate through every line in the CSV-file
    # ------------------------------------------
    $Counter = 1
    $CSV = Import-CSV -Path "$global:File"
    ForEach ($CSVLine in $CSV) {
        $Emaildomein = $CSVLine.'Name'
        $NameEmaildomein = $Emaildomein
        $Email = $Email.Split('@')[0] + "@" + $Emaildomein
        Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
        Write-Host "Creating mail domain" $Counter "of" $CSV.Length ":"
        Write-Host
        Write-Host "Mail domain information:" -ForegroundColor Magenta
        If ($Emaildomein) {
            Write-Host -NoNewLine " - Name: "; Write-Host $Emaildomein -ForegroundColor Yellow
        }
        Write-Host
        Write-Host "Creation information:" -ForegroundColor Magenta
        Write-Host -NoNewLine " - Selected mailbox: "; Write-Host $Name -ForegroundColor Yellow
        Write-Host -NoNewLine " - Applied Email Address: "; Write-Host $Email -ForegroundColor Yellow
        Write-Host
        $Counter++
        # -------------------------------------------------------------------
        # Go to function Messaging > Create Domains to create the mail domain
        # -------------------------------------------------------------------
        Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Create-Domains-*-Start" })
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}