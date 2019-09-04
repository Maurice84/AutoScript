Function Messaging-Exchange-Import-Mailboxes-5-InputArchiveFolderName {
    # ============
    # Declarations
    # ============
    $Task = "Select a target folder for the import"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($File -like "*@*") {
        $FileFormat = $File.Split('@')[0]
    } Else {
        $FileFormat = $File.Split('.')[0]
    }
    Write-Host
    Write-Host -NoNewLine "  Please enter a target folder name (for mailbox $global:Mailboxes) for the import (i.e.: "; Write-Host -NoNewLine ("Export " + $FileFormat) -ForegroundColor Yellow; Write-Host "),"
    Write-Host -NoNewLine "  or leave empty to import into the root of the mailbox: "
    $global:TargetFolder = Read-Host
    If (!$global:TargetFolder) {
        $global:TargetFolder = "n.a."
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TargetFolder
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}