Function Identity-ActiveDirectory-Modify-Passwords-2-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    # ----------------------------------------------------------------
    # Remove the file if it's present (Add-Content does not overwrite)
    # ----------------------------------------------------------------
    If ($global:OUFilter -ne "*") {
        $FileCSVP = "C:\Passwords-" + $global:OUFilter + "_" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + + ".csv"
    }
    Else {
        $FileCSVP = "C:\Passwords" + (Get-Date -ForegroundColor "yyyyMMdd-HHmm") + ".csv"
    }
    Remove-Item ($FileCSVP) -ErrorAction SilentlyContinue
    # ------------------------------------------------
    # Declare and add the header to the empty CSV-file
    # ------------------------------------------------
    $ExportHeader = '"Name","Username","Password","VerificationCode"'
    Add-Content ($FileCSVP) $ExportHeader
    # -------------------------------------------------------
    # Iterate through each account and add it to the CSV-file
    # -------------------------------------------------------
    ForEach ($Account in $global:Array) {
        $Name = $Account.Name
        $Username = ($Account.UserPrincipalName).ToLower()
        $SamAccountName = $Account.SamAccountName
        # --------------------------------
        # Create and set a random password
        # --------------------------------
        Script-Module-SetPassword
        Set-ADAccountPassword -Identity $SamAccountName -NewPassword (ConvertTo-SecureString $global:Password -AsPlainText -Force) -Reset
        # --------------------------------------------
        # Add the account and password to the CSV-file
        # --------------------------------------------
        $ExportValues = '"' + $Name + '","' + $Username + '","' + $global:Password + '","' + $global:VerificationCode + '"'
        Add-Content ($FileCSVP) $ExportValues
    }
    If (Test-Path $FileCSVP) {
        Write-Host -NoNewLine " - OK: Successfully exported the domain account(s) password(s) to CSV-file: "; Write-Host $FileCSVP -ForegroundColor Yellow
        # ----------------------------------
        # Convert the CSV-file to Excel-file
        # ----------------------------------
        Script-Convert-CSV-to-Excel -File $FileCSVP -Category "Accounts" -Silent $true
        $FileExcel = $FileCSVP.Substring(0, $FileCSVP.Length - 4) + ".xlsx"
        If (Test-Path $FileExcel) {
            Write-Host -NoNewLine " - OK: Successfully converted the CSV-file to Excel-file:"; Write-Host $FileExcel.Split('\')[-1] -ForegroundColor Yellow
        }
        Else {
            Write-Host " - ERROR: Could not convert the CSV-file to Excel-file, please investigate!" -ForegroundColor Red
        }
    }
    Else {
        Write-Host " - ERROR: Could not export the domain account(s) password(s)to CSV-file, please investigate!" -ForegroundColor Red
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}