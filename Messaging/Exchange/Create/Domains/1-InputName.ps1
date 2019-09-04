Function Messaging-Exchange-Create-Domains-1-InputName {
    # ============
    # Declarations
    # ============
    $Task = "Please enter the mail domain to create"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewline "  > Please enter a mail domain to create" -ForegroundColor Yellow
        $global:Emaildomein = Read-Host
        If (!$global:Emaildomein) {
            Write-Host "    Please enter a mail domain" -ForegroundColor Red; Write-Host
        }
        Else {
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Emaildomein
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}