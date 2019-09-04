Function Storage-FileSystem-Delete-Files-1-SelectObject {
    # ============
    # Declarations
    # ============
    $Task = "Select which object you would like to delete"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "  1. File(s)" -ForegroundColor Yellow
    Write-Host "  2. Folder(s)" -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "  Select an option or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
        [string]$InputChoice = Read-Host
        $InputKey = @("1", "2", "X") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "  Please use above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "1" {
            $global:Choice = "file(s)"
        }
        "2" {
            $global:Choice = "folder(s)"
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Choice
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}