Function Software-MicrosoftOutlook-Modify-AutoComplete-1-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    # --------------------------------
    # Checking and downloading NK2Edit
    # --------------------------------
    Write-Host " Step 1/3 - Checking presence of NK2Edit.exe..." -ForegroundColor Magenta
    $global:NK2Edit = (Get-ChildItem "C:\" | Where-Object { $_.Name -eq "NK2Edit.exe" }).FullName
    If ($global:NK2Edit) {
        Write-Host " - OK: NK2Edit found on location" $global:NK2Edit -ForegroundColor Green
    }
    Else {
        Write-Host " - INFO: NK2Edit not found, downloading..."
        $global:Source = $URL + "Tools/nk2edit.exe"
        $global:NK2Edit = "C:\NK2Edit.exe"
        $WebClient.DownloadFile($Source, $global:NK2Edit)
        If ($?) {
            Write-Host "   - OK: NK2Edit downloaded to" $global:NK2Edit -ForegroundColor Green
        }
        Else {
            Write-Host "   - ERROR: An error occurred downloading NK2Edit, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
    }
    Write-Host
    # -----------------------------------------------------------------
    # Checking and downloading Windows Automation Snapin for Powershell
    # -----------------------------------------------------------------
    Write-Host " Step 2/3 - Checking presence of Windows Automation Snapin for Powershell..." -ForegroundColor Magenta
    If (Get-PSSnapin WASP -Registered -ErrorAction SilentlyContinue) {
        Add-PSSnapin WASP
        If ($?) {
            Write-Host " - OK: Windows Automation Snapin for Powershell loaded" -ForegroundColor Green
        }
        Else {
            Write-Host " - ERROR: An error occurred loading Windows Automation Snapin for Powershell, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
    }
    Else {
        $global:WASP = (Get-ChildItem $env:windir | Where-Object { $_.Name -eq "WASP.dll" }).FullName
        If ($global:WASP) {
            Write-Host " - OK: Windows Automation Snapin for Powershell found on location" $global:WASP -ForegroundColor Green
        }
        Else {
            Write-Host " - INFO: Windows Automation Snapin for Powershell not found, downloading..."
            $global:Source = $URL + "Modules/WASP.dll"
            $global:WASP = $env:windir + "\WASP.dll"
            $WebClient.DownloadFile($Source, $global:WASP)
            If ($?) {
                Write-Host "   - OK: Windows Automation Snapin for Powershell downloaded to location" $global:WASP -ForegroundColor Green
            }
            Else {
                Write-Host "   - ERROR: An error occurred downloading Windows Automation Snapin for Powershell, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        $global:GetRuntimeDirectory = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
        Set-Alias InstallUtil (Resolve-Path (Join-Path $GetRuntimeDirectory installutil.exe))
        InstallUtil $global:WASP | Out-Null
        If ($?) {
            Write-Host " - OK: Windows Automation Snapin for Powershell successfully installed, now loading..." -ForegroundColor Green
            Add-PSSnapin WASP
            If ($?) {
                Write-Host "   - OK: Windows Automation Snapin for Powershell loaded" -ForegroundColor Green
            }
            Else {
                Write-Host "   - ERROR: An error occurred loading Windows Automation Snapin for Powershell, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        Else {
            Write-Host " - ERROR: An error occurred installing Windows Automation Snapin for Powershell, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
    }
    Write-Host
    # --------------------------
    # Checking Microsoft Outlook
    # --------------------------
    Write-Host " Step 3/3 - Checking the presence of Microsoft Outlook..." -ForegroundColor Magenta
    $global:OutlookPath = (Resolve-Path "C:\Program Files*\Microsoft Office\Office*\OUTLOOK.EXE").Path
    If ($global:OutlookPath) {
        If ($global:OutlookPath -like "*14*") { $global:Outlook = "2010" }
        If ($global:OutlookPath -like "*15*") { $global:Outlook = "2013" }
        If ($global:OutlookPath -like "*16*") { $global:Outlook = "2016" }
        Write-Host " - OK: Microsoft Outlook" $global:Outlook "detected" -ForegroundColor Green
    }
    Else {
        Write-Host " - ERROR: An error occurred detecting Microsoft Outlook. Please install the correct version (2016 or earlier, evaluation is possible). Now returning to the previous menu" -ForegroundColor Red
        Write-Host
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}