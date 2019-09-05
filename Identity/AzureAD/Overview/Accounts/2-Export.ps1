Function Identity-AzureAD-Overview-Accounts-2-Export {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "Exporting Azure AD account(s) to CSV-file..." -ForegroundColor Magenta
    # --------------------------------------------------------------------------------------------------------------
    # Declare the name and path of the CSV-file and remove the file if it's present (Add-Content does not overwrite)
    # --------------------------------------------------------------------------------------------------------------
    #$FileCSVO = "C:\Office365-" + $global:Office365Connect.Account + "_" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + ".csv"
    [string]$Domain = ($global:AzureADConnect).Account
    $FileCSVO = "C:\AzureAD-" + $Domain.Split('@')[1] + "_" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv"
    Remove-Item ($FileCSVO) -ErrorAction SilentlyContinue
    # ---------------------------------------------------
    # Declare the header and add it to the empty CSV-file
    # ---------------------------------------------------
    $ExportHeader = '"Date";"Name";"UserPrincipalName";"PrimarySmtpAddress";"License";"SyncStatus"'
    Add-Content ($FileCSVO) $ExportHeader
    # ----------------------------------------------------------------------------------
    # Iterate through each account with the current timestamp and add it to the CSV-file
    # ----------------------------------------------------------------------------------
    $Counter = 1
    ForEach ($AzureADAccount in ($global:Array | Sort-Object Name)) {
        Write-Host -NoNewLine ("Exporting Azure AD account $Counter of " + $global:Array.Count + ": "); Write-Host -NoNewLine $AzureADAccount.Name -ForegroundColor Yellow; Write-Host "..."
        $LastCheck = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $ExportValues = '"'`
            + $LastCheck + '";"'`
            + $AzureADAccount.Name + '";"'`
            + $AzureADAccount.UserPrincipalName + '";"'`
            + $AzureADAccount.PrimarySmtpAddress + '";"'`
            + $AzureADAccount.License + '";"'`
            + $AzureADAccount.SyncStatus + '"'
        # -----------------------------
        # Add each line to the CSV-file
        # -----------------------------
        Add-Content $FileCSVO $ExportValues
        $Counter++
    }
    Write-Host
    If (Test-Path $FileCSVO) {
        Write-Host -NoNewLine " - OK: Successfully exported Azure AD account(s) to CSV-file: "; Write-Host $FileCSVO -ForegroundColor Yellow
        # ----------------------------------
        # Convert the CSV-file to Excel-file
        # ----------------------------------
        Script-Convert-CSV-to-Excel -File $FileCSVO -Category "Account" -Silent $true
        If (Test-Path $FileExcel) {
            Write-Host -NoNewLine " - OK: Successfully converted the CSV-file to Excel-file: "; Write-Host $FileExcel -ForegroundColor Yellow
        }
        Else {
            Write-Host " - ERROR: Could not convert the CSV-file to Excel-file, please investigate!" -ForegroundColor Red
        }
    }
    Else {
        Write-Host "ERROR: CSV-file could not be created, please investigate!" -ForegroundColor Red
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}