Function Software-MicrosoftOutlook-Modify-AutoComplete-5-Start {
    # ============
    # Declarations
    # ============
    [string]$AutoCompleteStream = $OutlookAppLocal + "\RoamCache\Stream_Autocomplete*.dat"
    [string]$OutlookAppLocal = "$env:userprofile\AppData\Local\Microsoft\Outlook"
    [string]$OutlookAppRoaming = "$env:userprofile\AppData\Roaming\Microsoft\Outlook"
    If ($global:Outlook -eq "2010") {
        [string]$OutlookReg = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
    }
    If ($global:Outlook -eq "2013") {
        [string]$OutlookReg = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles"
    }
    If ($global:Outlook -eq "2016") {
        [string]$OutlookReg = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    }
    [int]$Count = 1
    [array]$Files = (Get-ChildItem ($global:Drive + "AutoComplete") -Force -ErrorAction SilentlyContinue | Where-Object { ($_.Name -like "*.nk2" -OR $_.Name -like "*.dat") -AND $_.PSIsContainer -eq $false }).FullName
    # =========
    # Execution
    # =========
    # -------------------------
    # Iterate through each file
    # -------------------------
    ForEach ($File in $Files) {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        Write-Host (" Processing AutoComplete file " + $Count + " of " + $Files.Length) -ForegroundColor Magenta
        $FileName = $File.Split('\')[-1]
        $Email = $FileName.Substring(0, $FileName.Length - 4)
        $Name = $Email
        $SkipProfile = $false
        If ($Email -like "*-*") {
            If ($Email -notlike "*-1") {
                $SkipProfile = $true
            }
            $Email = $Email.Substring(0, $Email.Length - 2)
        }
        If ($SkipProfile -eq $false) {
            # ---------------------------------------------------------------------
            # If Outlook is active at this moment, then force close the application
            # ---------------------------------------------------------------------
            Stop-Process -Name OUTLOOK -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            # ------------------------------------------------
            # If an Outlook profile is present, then remove it
            # ------------------------------------------------
            If (Test-Path $OutlookReg) {
                Write-Host " - Removing current Outlook profile from the registry..."
                Remove-Item $OutlookReg -Recurse
            }
            If (Test-Path $OutlookAppLocal) {
                Write-Host " - Removing current Outlook profile from AppData\Local..."
                Remove-Item $OutlookAppLocal -Recurse -Force
            }
            If (Test-Path $OutlookAppRoaming) {
                Write-Host " - Removing current Outlook profile from AppData\Roaming..."
                Remove-Item $OutlookAppRoaming -Recurse -Force
            }
            Write-Host
            # ---------------------------------------------------------
            # Starting Outlook with the CleanAutoCompleteCache argument
            # ---------------------------------------------------------
            Write-Host (" - " + $Name + ": Starting Outlook with CleanAutoCompleteCache argument")
            Start-Process $OutlookPath -ArgumentList "/CleanAutoCompleteCache"
            # ----------------------------------------------------------------------------
            # Wait until Outlook is loaded and check the language to declare the variables
            # ----------------------------------------------------------------------------
            While (!(Get-Process | Where-Object { $_.MainWindowTitle -like "*Outlook*" })) { Start-Sleep -Seconds 3 }
            If ((Get-Process | Where-Object { $_.MainWindowTitle -like "*Openen*" })) {
                $NewEmail = "Naamloos"
                $NewProfile = "Nieuw profiel"
                $NewAccount = "Account toevoegen"
                $NewEmailAccount = "E-mailaccount toevoegen"
                $Reminder = "Herinnering"
            }       
            If ((Get-Process | Where-Object { $_.MainWindowTitle -like "*Opening*" })) {
                $NewEmail = "Untitled"
                $NewProfile = "New Profile"
                $NewAccount = "Add Account"
                $NewEmailAccount = "Add E-mail Account"
                $Reminder = "Reminder"
            }
            # ----------------------------------------------------------------
            # Waiting for the Welcome wizard to appear with New Profile window
            # ----------------------------------------------------------------
            Write-Host (" - " + $Name + ": Waiting for window: " + $NewProfile)
            While ((Select-Window -Title "*Outlook*" | Select-ChildWindow) -notlike "*$NewProfile*") { Start-Sleep -Seconds 2 }
            Select-Window -Title "*Outlook*" | Select-ChildWindow | Set-WindowActive | Send-Keys "Outlook{Enter}"
            Start-Sleep -Seconds 2
            Write-Host (" - " + $Name + ": Waiting for window: " + $NewAccount)
            If ($Outlook -eq "2010") {
                $SendKeys = "{Tab}" + $Email + "{Tab}{Enter}"
            }
            Else {
                $SendKeys = "{Tab}{Tab}" + $Email + "{Tab}{Tab}{Tab}{Enter}"
            }
            While ((Select-Window -Title "*Outlook*" | Select-ChildWindow) -notlike "*$NewAccount*") { Start-Sleep -Seconds 2 }
            Select-Window -Title "*Outlook*" | Select-ChildWindow | Set-WindowActive | Send-Keys $SendKeys
            # --------------------------------------------------------------------
            # Waiting until email address is detected, then pressing OK and Finish
            # --------------------------------------------------------------------
            Write-Host (" - " + $Name + ": Waiting for auto discovery...")
            While (((Select-Window -Title "*Outlook*" | Select-ChildWindow | Select-ChildWindow) -notlike "*$NewEmailAccount*") -OR ((Select-Window -Title "*Outlook*" | Select-ChildWindow | Select-ChildWindow) -eq $null)) { Start-Sleep -Seconds 2 }
            Select-Window -Title "*Outlook*" | Select-ChildWindow | Select-ChildWindow | Set-WindowActive | Send-Keys "{Enter}"
            Start-Sleep -Seconds 2
            Select-Window -Title "*Outlook*" | Select-ChildWindow | Set-WindowActive | Send-Keys "{Enter}"
            # -----------------------------------------------------------------
            # Opening the mailbox and checking the presence of the Stream files
            # -----------------------------------------------------------------
            Write-Host (" - " + $Name + ": Waiting until the mailbox is loaded (presence of Stream files)...")
            While (!(Test-Path ($OutlookAppLocal + "\RoamCache\Stream_*.dat"))) { Start-Sleep -Seconds 1 }
            Start-Sleep -Seconds 3
            If (Select-Window -Title "*$Reminder*") {
                Select-Window -Title "*$Reminder*" | Set-WindowActive | Send-Keys "{Esc}"
            }
            # --------------------------------------------------------------------
            # Creating and sending a test mail to prepare AutoComplete information
            # --------------------------------------------------------------------
            Write-Host (" - " + $Name + ": Creating and sending a test mail to prepare AutoComplete information")
            Select-Window -Title "*Outlook*" | Set-WindowActive | Send-Keys "^(n)"; Start-Sleep -Seconds 2
            Select-Window -Title "*$NewEmail*" | Set-WindowActive | Send-Keys $global:TestEmailAddress; Start-Sleep -Seconds 1
            Select-Window -Title "*$NewEmail*" | Set-WindowActive | Send-Keys "{Tab}{Tab}Test mail"; Start-Sleep -Seconds 1
            Select-Window -Title "*$NewEmail*" | Set-WindowActive | Send-Keys "^({Enter}){Enter}"; Start-Sleep -Seconds 1
            Start-Sleep -Seconds 3
            # -----------------------------------------------------------------------------------------------
            # Closing Outlook and waiting until the AutoComplete file is generated and at least 5 seconds old
            # -----------------------------------------------------------------------------------------------
            Write-Host (" - " + $Name + ": Closing Outlook and waiting until the AutoComplete file is generated...")
            Select-Window -Title "*Outlook*" | Set-WindowActive | Send-Keys "%({F4})"
            While (!(Resolve-Path ($OutlookAppLocal + "\RoamCache\Stream_Autocomplete*.dat"))) { Start-Sleep -Seconds 2 }
            $AutoCompleteStream = (Resolve-Path ($OutlookAppLocal + "\RoamCache\Stream_Autocomplete*.dat"))
            If (Test-Path $AutoCompleteStream) {
                Write-Host (" - " + $Name + ": OK: Generated AutoComplete file is found")
            }
            Else {
                Write-Host (" - " + $Name + ": ERROR: An error occurred locating the generated AutoComplete file, please investigate!") -ForegroundColor Red
                $Pause.Invoke()
                EXIT
            }
            While (((Get-Date) - ((Get-ChildItem $AutoCompleteStream | Select-Object LastWriteTime | Sort-Object LastWriteTime -Descending)[0]).LastWriteTime).Seconds -lt 5) { Start-Sleep -Seconds 2 }
        }
        # ----------------------------------------------------------------------------------------------------
        # Appending the contents of the old AutoComplete file to the generated AutoComplete file using NK2Edit
        # ----------------------------------------------------------------------------------------------------
        $Log = $File + ".log"
        $AutoCompleteStream = (Resolve-Path $AutoCompleteStream)
        Write-Host (" - " + $Name + ": Appending the contents of the old AutoComplete file to the generated AutoComplete file...")
        Start-Process $NK2Edit -ArgumentList "/LogFile `"$Log`" /nofirstbackup /nobackup /import_full_nk2 `"$File`""
        While ((Get-Process | Where-Object { $_.ProcessName -like "*NK2Edit*" })) { Start-Sleep -Seconds 2 }
        If ((Get-Content ($Log) | Select-String -Pattern 'Successfully saved')[-1]) {
            Write-Host (" - " + $Name + ": OK: Successfully appended to the generated AutoComplete file")
        }
        Else {
            Write-Host (" - " + $Name + ": ERROR: An error occurred appending to the generate AutoComplete file, please investigate!") -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
        # -----------------------------------------------------------------------------------------------
        # Cleaning up old Exchange entries from the AutoComplete file and import the contents to Exchange
        # -----------------------------------------------------------------------------------------------
        Write-Host (" - " + $Name + ": Cleaning up old Exchange entries from the AutoComplete file and import the contents to Exchange...")
        Start-Process $NK2Edit -ArgumentList "/AutoExportToMessageStore 1 /LogFileAppend `"$Log`" /nofirstbackup /nobackup /script_express `"$AutoCompleteStream`" 'If AddressType equal `"EX`" Delete' 'If AddressType equal `"MAPIPDL`" Delete'"
        While ((Get-Process | Where-Object { $_.ProcessName -like "*NK2Edit*" })) { Start-Sleep -Seconds 2 }
        If ((Get-Content ($Log) | Select-String -Pattern 'Successfully saved')[-1]) {
            Write-Host (" - " + $Name + ": OK: AutoComplete file successfully cleaned from old entries")
        }
        Else {
            Write-Host (" - " + $Name + ": ERROR: An error occurred cleaning the AutoComplete file, please investigate!") -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
        If ((Get-Content ($Log) | Select-String -Pattern 'Successfully copied the AutoComplete file into the message store')[-1]) {
            Write-Host (" - " + $Name + ": OK: AutoComplete file successfully imported to Exchange")
            Move-Item $File ($File + "-Imported")
            If ($?) {
                Write-Host
                Write-Host -NoNewLine " - AutoComplete file successfully processed. Proceeding to the next file in 2 seconds..." -ForegroundColor Green
                $Count++
                Start-Sleep -Seconds 2
            }
            Else {
                Write-Host (" - " + $Name + ": ERROR: An error occurred renaming $File to " + $File + "-Imported, please investigate!") -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        Else {
            Write-Host (" - " + $Name + ": ERROR: An error occurred importing the AutoComplete file to Exchange, please investigate!") -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
    }
    # ==========
    # Finalizing
    # ==========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host (" Finished AutoComplete file " + ($Count - 1) + " of " + $Files.Length) -ForegroundColor Green
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}