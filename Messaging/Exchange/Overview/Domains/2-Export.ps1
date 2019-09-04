Function Messaging-Exchange-Overview-Domains-2-Export {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host "Exporting mail domain(s) to CSV-file..." -ForegroundColor Magenta
    # ------------------------------------------------------------------------------------------------------------------
    # Declare the name and path of the CSV-file and remove the CSV-file if it's present (Add-Content does not overwrite)
    # ------------------------------------------------------------------------------------------------------------------
    $Domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $FileCSVE = "C:\Maildomains-" + $Domain + "_" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + ".csv"
    Remove-Item ($FileCSVE) -ErrorAction SilentlyContinue
    # ---------------------------------------------------
    # Declare the header and add it to the empty CSV-file
    # ---------------------------------------------------
    $ExportHeader = '"Date";"Name";"DomainType";"Default"'
    Add-Content ($FileCSVE) $ExportHeader
    # ----------------------------------------------------------------------------------
    # Iterate through each account with the current timestamp and add it to the CSV-file
    # ----------------------------------------------------------------------------------
    $Counter = 1
    ForEach ($Domain in ($global:Array | Sort-Object Name)) {
        Write-Host ("Exporting mail domain $Counter of " + $global:Array.Count + ": " + $Domain.Name + "...")
        $LastCheck = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $ExportValues = '"' + $LastCheck + '";"' + $Domain.Name + '";"' + $Domain.DomainType + '";"' + $Domain.Default + '"'
        Add-Content $FileCSVE $ExportValues
        $Counter++
    }
    Write-Host
    If (Test-Path $FileCSVE) {
        Write-Host -NoNewLine "OK: Mail domain(s) successfully exported to CSV-file: "; Write-Host $FileCSVE -ForegroundColor Yellow
        Write-Host
        Script-Convert-CSV-to-Excel -File $FileCSVE -Category "Emaildomeinen"
    }
    Else {
        Write-Host "ERROR: An error occurred creating the CSV-file, please investigate!" -ForegroundColor Red
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}