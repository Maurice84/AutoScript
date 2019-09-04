Function Roles-IIS-Modify-OWA-LogonUPN-1-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "Checking the Exchange Server installation path..."
    if ($env:ExchangeInstallPath) {
        Write-Host "OK: Exchange Server path found: $env:ExchangeInstallPath"
        # -----------------------------------------------
        # Reading the fexppw.js file and modify if needed
        # -----------------------------------------------
        $Spaces = (" " * 20)
        $String = $Spaces + "gbid(`"username`").focus();"
        $Setting = $Spaces + "gbid(`"username`").value = `"" + $env:userdomain + "\\`" + rg[3];`r`n" + $String
        $Folders = Get-ChildItem ($env:ExchangeInstallPath + "FrontEnd\HttpProxy\owa\auth\") | Where-Object { $_.PSIsContainer -eq $true }
        ForEach ($Folder in ($Folders | Select-Object FullName).FullName) {
            $File = $Folder + "\scripts\premium\fexppw.js"
            $Content = Get-Content $File
            If ($Content | Where-Object { $_ -like "*$env:userdomain*" }) {
                Write-Host "INFO: The modification was already done in $File" -ForegroundColor Gray
                $Skipped = $true
            }
            Else {
                $Content.Replace($String, $Setting) | Set-Content $File
                If ($?) {
                    Write-Host -NoNewLine "OK: Modification successfully done in " -ForegroundColor Green; Write-Host $File -ForegroundColor Yellow
                    $Done = $true
                }
                Else {
                    Write-Host "ERROR: An error occurred modifying $File, please investigate!" -ForegroundColor Magenta
                }
            }
        }
        Write-Host
        if ($Done) {
            Write-Host -NoNewLine "Restarting Microsoft IIS Web Server..."
            Invoke-Expression "iisreset.exe /noforce" | Out-Null
            If ($?) {
                Write-Host; Write-Host
                Write-Host "OK: Microsoft IIS Web Server is restarted." -ForegroundColor Green
            }
            Else {
                Write-Host; Write-Host
                Write-Host "ERROR: An error occurred restarting Microsoft IIS Web Server, please investigate!" -ForegroundColor Red
            }
        }
        if ($Skipped) {
            Write-Host "No modification was necessary. Now returning to the previous menu."
        }
    }
    Else {
        Write-Host "ERROR: Is this not an Exchange server?" -ForegroundColor Red
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}