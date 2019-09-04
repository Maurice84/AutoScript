Function Maintenance-WindowsServer-Send-Message-4-Start {
    param(
        [string]$Message,
        [array]$Servers
    )
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "Sending message to server(s):" -ForegroundColor Magenta
    ForEach ($Server in $Servers) {
        Write-Host -NoNewline (" - " + $Server + "...")
        Invoke-Command -ComputerName $Server -ArgumentList $Message -ScriptBlock {
            param(
                $Message
            );
            msg * $Message
        } -ErrorAction SilentlyContinue
        Write-Host " OK!" -ForegroundColor Green
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
}