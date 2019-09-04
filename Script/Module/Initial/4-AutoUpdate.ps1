Function Script-Module-Initial-4-AutoUpdate {
<#
    Script-Module-SetHeaders -Name $MainTitel
    If ($InputArgs0 -eq "FTP") {
        # ===================================================================
        Write-Host "Step 1/2 - Verwijderen van oude Maurice AutoScript op FTP" -ForegroundColor Magenta
        # ===================================================================
        $Request = [System.Net.WebRequest]::Create($URL) 
        $Request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
        $Request.Credentials = $Credentials
        $Response = $Request.GetResponse()
        $Reader = New-Object IO.StreamReader $Response.GetResponseStream()
        $Objects = $Reader.ReadToEnd()
        $Reader.Close()
        $Response.Close()
        $Objects = $Objects.Replace('Web/','')
        $Objects = $Objects.Split("`r")
        $Objects = $Objects -split '\s+' -match '\S'
        $Files = $Objects | Where-Object {$_ -like "*_Maurice-AutoScript*"}
        If ($Files) {
            ForEach ($File in $Files) {
                $Request = [System.Net.WebRequest]::Create($URL+$File) 
                $Request.Method = [System.Net.WebRequestMethods+FTP]::DeleteFile
                $Request.Credentials = $Credentials
                $Response = $Request.GetResponse()
                If ($?) {
                    Write-Host "OK:" ($URL+$File) "is verwijderd"
                } Else {
                    Write-Host "FOUT:" ($URL+$File) "kon niet verwijderd worden, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    EXIT
                }
            }
        } Else {
            Write-Host "INFO: Er is geen Maurice AutoScript gevonden op de FTP server" -ForegroundColor Gray
        }
        Write-Host
        # =====================================================================
        Write-Host "Step 2/2 - Kopieren van huidige Maurice AutoScript naar FTP" -ForegroundColor Magenta
        # =====================================================================
        $VersionDate = "_Maurice-AutoScript-" + (Get-Date -format "yyyyMMdd-HHmmss") + ".ps1"
        $PS1 = Get-ChildItem $($MyInvocation.ScriptName)
        $URI = New-Object System.Uri($URL+$VersionDate)
        $WebClient.UploadFile($URI, $PS1) 
        If ($?) {
            Write-Host "OK: Maurice AutoScript is succesvol geupload naar de FTP server"
        } Else {
            Write-Host "ERROR: An error occurred Maurice AutoScript kon niet verwijderd worden van de FTP server, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
        }
        EXIT
    } ElseIf($InputArgs0 -eq "Updated") {
        ### Tijdelijk i.v.m. naamomzetting NCAS naar Maurice AutoScript
        Remove-Item $InputArgs1 -ErrorAction SilentlyContinue -Confirm:$False
        $InputArgs1 = [string](Get-ChildItem $($MyInvocation.ScriptName)).Directory+"\_Maurice-AutoScript-AutoUpdate.ps1"
        ###
        Copy-Item $MyInvocation.ScriptName $InputArgs1 -Confirm:$False
        If ($?) {
            Write-Host "OK: De huidige versie van Maurice AutoScript is bijgewerkt"
            Remove-Item $MyInvocation.ScriptName -ErrorAction SilentlyContinue -Confirm:$False
            If ($?) {
                Write-Host "OK: Tijdelijke bestanden zijn opgeruimd, nu bezig met starten van de nieuwe versie van Maurice AutoScript..."
                If ($global:CheckUAC -eq $False) {
                    Write-Host "INFO: Bezig met omzeilen van User Account Control..."
                    Copy-Item $InputArgs1 $env:temp -Confirm:$false
                    If ($?) {
                        $BypassUAC = ($env:temp + "\_Maurice-AutoScript-AutoUpdate.ps1")
                        Start-Process "powershell.exe" -ArgumentList "-File $BypassUAC" -Verb RunAs
                        EXIT
                    } Else {
                        Write-Host "ERROR: An error occurred Het bestand kon niet gekopieerd worden naar de temp folder waardoor UAC niet kan worden omzeilt, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        EXIT
                    }
                } Else {
                    $ArgumentList = '-File "' + $InputArgs1 + '"'
                    Start-Process "powershell.exe" -ArgumentList $ArgumentList
                    EXIT
                }
            } Else {
                Write-Host "ERROR: An error occurred De tijdelijke bestanden konden niet verwijderd worden, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            }
        } Else {
            Write-Host "ERROR: An error occurred De huidige versie van Maurice AutoScript kon niet verwijderd worden, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
        }
    } Else {
        If ($Shell -like "*ISE*" -OR $Shell -like "*VSCode*") {
            Write-Host "INFO: Maurice AutoScript wordt uitgevoerd in de IDE modus, de update wordt genegeerd" -ForegroundColor Yellow
            If ($global:CheckUAC -eq $false) {
                Write-Host "LET OP: PowerShell ISE heeft momenteel geen Administrator bevoegdheden. Herstart deze met Run as Administrator, anders kan je bijv geen CSV's exporteren!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        } Else {
            Script-Module-SetHeaders -Name $MainTitel
            # ====================================================
            Write-Host "Step 1/2 - Bezig met controleren van update" -ForegroundColor Magenta
            # ====================================================
            $Request = [System.Net.WebRequest]::Create($URL) 
            $Request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
            $Request.Credentials = $Credentials
            $Response = $Request.GetResponse()
            $Reader = New-Object IO.StreamReader $Response.GetResponseStream()
            $Objects = $Reader.ReadToEnd()
            $Reader.Close()
            $Response.Close()
            $Objects = $Objects.Split("`r")
            $Objects = $Objects -split '\s+' -match '\S'
            $File = $Objects | Where-Object {$_ -like "_Maurice-AutoScript*"}
            If ($File) {            
                $VersionFTP = [datetime]::parseexact($File.Replace("_Maurice-AutoScript-","").Replace(".ps1",""), 'yyyyMMdd-HHmmss', $null)
                $VersionCurrent = (Get-ChildItem $($MyInvocation.ScriptName) | Select-Object LastWriteTime).LastWriteTime
                If ($VersionFTP -gt $VersionCurrent) {
                    Write-Host "OK: Er is een nieuwere versie van Maurice AutoScript gevonden" -ForegroundColor Gray
                    $FilePath = [string](Get-ChildItem $($MyInvocation.ScriptName)).Directory+"\"+$File
                    Write-Host
                    # ====================================================
                    Write-Host "Step 2/2 - Bezig met downloaden van update" -ForegroundColor Magenta
                    # ====================================================
                    $Download = Try {$WebClient.DownloadFile($URL+$File, $FilePath)} Catch {$false}
                    If ($Download -eq $false) {
                        Write-Host "Helaas, de nieuwste versie van Maurice AutoScript kon niet worden gedownload. Het kan zijn dat de folder op read-only staat, de update wordt nu genegeerd." -ForegroundColor Yellow
                        Write-Host
                        $Pause.Invoke()
                    } Else {
                        Write-Host "OK: Nieuwe versie van Maurice AutoScript is gedownload"
                        $FileCurrent = Get-ChildItem $($MyInvocation.ScriptName)
                        . $FilePath Updated "$FileCurrent"
                        EXIT
                    }
                } Else {
                    Write-Host "INFO: Er is geen nieuwere versie van Maurice AutoScript gevonden op de FTP server" -ForegroundColor Gray
                    If ($global:CheckUAC -eq $False) {
                        Write-Host "INFO: Bezig met omzeilen van User Account Control..."
                        Copy-Item (Get-ChildItem $($MyInvocation.ScriptName)) $env:temp -Confirm:$false
                        If ($?) {
                            $BypassUAC = ($env:temp + "\_Maurice-AutoScript-AutoUpdate.ps1")
                            Start-Process "powershell.exe" -ArgumentList "-File $BypassUAC" -Verb RunAs
                            EXIT
                        } Else {
                            Write-Host "ERROR: An error occurred Het bestand kon niet gekopieerd worden naar de temp folder waardoor UAC niet kan worden omzeilt, please investigate!" -ForegroundColor Red
                            $Pause.Invoke()
                            EXIT
                        }
                    }
                }
            } Else {
                Write-Host "INFO: Er is geen Maurice AutoScript gevonden op de FTP server" -ForegroundColor Gray
                If ($global:CheckUAC -eq $False) {
                    Write-Host "INFO: Bezig met omzeilen van User Account Control..."
                    Copy-Item (Get-ChildItem $($MyInvocation.ScriptName)) $env:temp -Confirm:$false
                    If ($?) {
                        $BypassUAC = ($env:temp + "\_Maurice-AutoScript-AutoUpdate.ps1")
                        Start-Process "powershell.exe" -ArgumentList "-File $BypassUAC" -Verb RunAs
                        EXIT
                    } Else {
                        Write-Host "ERROR: An error occurred Het bestand kon niet gekopieerd worden naar de temp folder waardoor UAC niet kan worden omzeilt, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        EXIT
                    }
                }
            }
        }
    }
    Pause
#>
}
