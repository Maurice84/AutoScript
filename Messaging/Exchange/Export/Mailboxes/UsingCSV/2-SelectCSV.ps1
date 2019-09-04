Function Messaging-Exchange-Export-Mailboxes-UsingCSV-2-SelectCSV {
    # ============
    # Declarations
    # ============
    $Task = "Select the CSV-file(s) with mailbox(es) which you want to export"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Bestand" -Functie "Selecteren" -Filter "CSV"
    $global:CSVFile = $global:Array.Name
    $global:Mailboxes = @()
    $global:CSV = Import-CSV -Path "$global:CSVFile" -Delimiter ";"
    ForEach ($CSVLine in $global:CSV) {
        $Name = $CSVLine.'Name'
        If (!$Name) {
            $Name = $CSVLine.'DisplayName'
        }
        $UserPrincipalName = $CSVLine.'UserPrincipalName'
        If (!$UserPrincipalName) {
            If ($CredentialSourceUPN -like "*\*") {
                $UserPrincipalName = $CredentialSourceUPN.Split('\')[0] + "\" + $CSVLine.'SamAccountName'
            }
            Else {
                $UserPrincipalName = $CSVLine.'SamAccountName'
            }
        }
        $Properties = @{Name = $Name; UserPrincipalName = $UserPrincipalName }
        $CSVObject = New-Object PSObject -Property $Properties
        $global:Mailboxes += $CSVObject
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ($global:CSVFile.Split('\')[-1] + ": " + ([string]$global:Mailboxes.Count + " mailbox(es)"))
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    # -------------------------------------------------------------
    # Go to function Messaging > Export Mailboxes to export mailbox
    # -------------------------------------------------------------
    Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Export-Mailboxes-*-SetDate" })
}