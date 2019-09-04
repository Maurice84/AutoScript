Function Storage-FileSystem-Modify-AutoComplete-5-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # -----------------------------------------
    # Iterate through each line of the CSV-file
    # -----------------------------------------
    $Total = 0
    $Counter = 1
    $CSV = Import-CSV -Path "$global:FileCSV"
    ForEach ($CSVLine in $CSV) {
        $AutoScript = $CSVLine.'Date'
        $CSVExport = "exported by Maurice AutoScript (full overview)"
        $Email = $CSVLine.'PrimarySmtpAddress'
        If ($Email) {
            $Total++
        }
    }
    ForEach ($CSVLine in $CSV) {
        $GivenName = $CSVLine.'GivenName'
        $Surname = $CSVLine.'Surname'
        $SamAccountName = $CSVLine.'SamAccountName'
        $Email = $CSVLine.'PrimarySmtpAddress'
        If ($Email) {
            Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
            Write-Host "Selected CSV is" $CSVExport
            Write-Host "Processing domain account" $Counter "of" $Total "to rename the AutoComplete file(s):"
            Write-Host
            Write-Host "Domain account info:" -ForegroundColor Magenta
            If ($Surname) {
                $Name = $GivenName + " " + $Surname
            }
            Else {
                $Name = $GivenName
            } Write-Host -NoNewLine " - Name: "; Write-Host $Name -ForegroundColor Yellow
            If ($SamAccountName) { Write-Host -NoNewLine " - UserPrincipalName: "; Write-Host $SamAccountName -ForegroundColor Yellow }
            If ($Email) { Write-Host -NoNewLine " - PrimarySmtpAddress: "; Write-Host $Email -ForegroundColor Yellow }
            Write-Host
            $Counter++
            $Objects = @()
            [array]$Objects = Get-ChildItem ($Drive + "AutoComplete\" + "$SamAccountName") -Force -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*.dat" -OR $_.Name -like "*.nk2" } | % { $_.FullName }
            If ($Objects) {
                $Count = 1
                ForEach ($Object in $Objects) {
                    If ($Objects.Count -gt 1) {
                        $FileName = $Email + "-" + $Count + $Object.Substring($Object.Length - 4)
                        $Count++
                    }
                    Else {
                        $FileName = $Email + $Object.Substring($Object.Length - 4)
                    }
                    Copy-Item $Object ($Drive + "AutoComplete\" + $FileName)
                }
            }
        }
    }
    # ==========
    # Finalizing
    # ==========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}