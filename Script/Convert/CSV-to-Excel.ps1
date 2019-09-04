Function Script-Convert-CSV-to-Excel ([string]$File, [string]$Category, [string]$Silent) {
    If ($PSVersionTable.PSVersion.Major -ge 3) {
        $global:Fout = $null
        If (!$Silent) {
            Write-Host "Converting CSV to Excel:" -ForegroundColor Magenta
            Write-Host "Checking for presence of the EPPlus Snapin for Powershell..."
        }
        $global:EPPlus = (Get-ChildItem "C:\" | Where-Object { $_.Name -eq "EPPlus.dll" }).FullName
        If (!$global:EPPlus) {
            If (!$Silent) {
                Write-Host "INFO: EPPlus Snapin for Powershell not found, downloading..."
            }
            $Source = $URL + "Modules/EPPlus.dll"
            $global:EPPlus = "C:\EPPlus.dll"
            $WebClient.DownloadFile($Source, $global:EPPlus)
            If ($?) {
                If (!$Silent) {
                    Write-Host "OK: EPPlus Snapin for Powershell downloaded to " $global:EPPlus -ForegroundColor Green
                }
            }
            Else {
                Write-Host "ERROR: EPPlus Snapin for Powershell could not be downloaded! Please investigate" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        If (!$Silent) {
            Write-Host "Loading EPPlus Snapin for Powershell..."
        }
        [Reflection.Assembly]::LoadFile($global:EPPlus) | Out-Null
        If ($?) {
            # -----------------------------------------------------------------------------
            # Verwijderen van Excel als deze al bestaat (deze worden nl. niet overschreven)
            # -----------------------------------------------------------------------------
            $global:FileExcel = $File.Substring(0, $File.Length - 4) + ".xlsx"
            Remove-Item ($FileExcel) -ErrorAction SilentlyContinue
            # ---------------------
            # Variabelen declareren
            # ---------------------
            $global:ExcelPackage = New-Object OfficeOpenXml.ExcelPackage
            $global:ExcelTextFormat = New-Object OfficeOpenXml.ExcelTextFormat
            $global:Worksheet = $ExcelPackage.Workbook.Worksheets.Add($Category)
            # ---------------------------------------------------------------------------------------
            # Instellen Excel formaat o.a. Text Qualifier, deze is nodig als de cellen beginnen met "
            # ---------------------------------------------------------------------------------------
            $ExcelTextFormat.Delimiter = ";"
            $ExcelTextFormat.TextQualifier = '"'
            #$ExcelTextFormat.Encoding = [System.Text.Encoding]::UTF8
            $ExcelTextFormat.SkipLinesBeginning = '0'
            $ExcelTextFormat.SkipLinesEnd = '1'
            # -------------------------------------------------------------------------------------------
            # Instellen Excel tabelstijl (zie http://www.nudoq.org/#!/Packages/EPPlus/EPPlus/TableStyles)
            # -------------------------------------------------------------------------------------------
            $global:TableStyle = [OfficeOpenXml.Table.TableStyles]::Light10
            # ------------------------------------------------------------------------------------------
            # Converteren van de zojuist gemaakte CSV met tabelstijl, eerste regel als header en autofit
            # ------------------------------------------------------------------------------------------
            $global:null = $Worksheet.Cells.LoadFromText((Get-Item $File), $ExcelTextFormat, $TableStyle, $true)
            $global:ColumnHeaders = (Get-Content $File | Select-Object -First 1).Split(";").Replace('"', '')
            If ($Category -eq "Accounts") {
                $Worksheet.Cells[$Worksheet.Dimension.Address].Style.HorizontalAlignment = "Center"
                $global:Counter = 1
                ForEach ($global:Column in $ColumnHeaders) {
                    If ($Column -eq "GivenName" -OR $Column -eq "Surname") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Left"
                    }
                    ElseIf ($Column -eq "Date" -OR $Column -like "Groep: *" -OR $Column -like "Group *") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Center"
                    }
                    Else {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Right"
                    }
                    $Counter++
                }
            }
            ElseIf ($Category -eq "Account") {
                $Worksheet.Cells[$Worksheet.Dimension.Address].Style.HorizontalAlignment = "Right"
                $global:Counter = 1
                ForEach ($global:Column in $ColumnHeaders) {
                    If ($Column -eq "GivenName" -OR $Column -eq "Surname") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Left"
                    }
                    ElseIf ($Column -eq "Date:") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Center"
                    }
                    Else {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Right"
                    }
                    $Counter++
                }
            }
            ElseIf ($Category -like "Mailbox*") {
                $Worksheet.Cells[$Worksheet.Dimension.Address].Style.HorizontalAlignment = "Right"
                $global:Counter = 1
                ForEach ($global:Column in $ColumnHeaders) {
                    If ($Column -eq "Name") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Left"
                    }
                    ElseIf ($Column -eq "Date") {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Center"
                    }
                    Else {
                        $Worksheet.Column($Counter).Style.HorizontalAlignment = "Right"
                    }
                    $Counter++
                }
            }
            Else {
                $Worksheet.Cells[$Worksheet.Dimension.Address].Style.HorizontalAlignment = "Right"
                $Worksheet.Column(1).Style.HorizontalAlignment = "Left"
                $Worksheet.Column(2).Style.HorizontalAlignment = "Left"
                $Worksheet.Column(3).Style.HorizontalAlignment = "Left"
            }
            $Worksheet.Row(1).Style.HorizontalAlignment = "Left"
            $Worksheet.Cells[$Worksheet.Dimension.Address].AutoFitColumns()
            $ExcelPackage.SaveAs($FileExcel)
            If (Test-Path $FileExcel) {
                If (!$Silent) {
                    Write-Host -NoNewLine "OK: Successfully converted to Excel, located at: "; Write-Host $FileExcel -ForegroundColor Yellow
                }
            }
            Else {
                $global:Fout = $true
                Write-Host "ERROR: A problem occurred during conversion! Please investigate" -ForegroundColor Red
            }
        }
        Else {
            $global:Fout = $true
            Write-Host "ERROR: EPPlus Snapin for Powershell could not be loaded! Please investigate" -ForegroundColor Red
        }
    }
}