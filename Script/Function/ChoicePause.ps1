Function Script-Function-ChoicePause {
    param (
        [string]$Name,
        [string]$Task
    )
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "  > Would you like to press a key when a file has been detected? (Y/N): ";
        [string]$InputChoice = Read-Host
        $InputKey = @("Y", "N") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "    Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "Y" {
            $PauseProcessing = "Yes"
        }
        "N" {
            $PauseProcessing = "No"
        }
    }
    # ==========
    # Finalizing
    # ==========
    return $PauseProcessing

}