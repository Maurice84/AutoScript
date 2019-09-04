Function Messaging-Exchange-Move-Mailboxes-5-SelectCSV {
    # ============
    # Declarations
    # ============
    $Task = "Select the CSV-file(s) with source mailbox(es) which you want to move"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Bestand" -Functie "Selecteren" -Filter "CSV"
    $File = $global:Array.Name
    $global:Mailboxes = @()
    $global:CSV = Import-CSV -Path "$File" -Delimiter ";"
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
        $Object = New-Object PSObject -Property $Properties
        $global:Mailboxes += $Object
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ($File.Split('\')[-1] + [string]$global:Mailboxes.Count + " mailbox(es)")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 2]
}