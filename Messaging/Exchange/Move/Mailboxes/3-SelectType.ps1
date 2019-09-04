Function Messaging-Exchange-Move-Mailboxes-3-SelectType {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "  Would you like to perform a pre-stage migration by copying the ("; Write-Host -NoNewLine "B" -ForegroundColor Yellow; Write-Host -NoNewLine ")ulk of the mailbox mail-items, or would you like to move ("; Write-Host -NoNewLine "A" -ForegroundColor Yellow; Write-Host -NoNewLine ")ll mailboxes? "
        $Choice = Read-Host
        $Input = @("B","A") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "B" {
            $global:SuspendWhenReadyToComplete = $true
            $global:MigratieType = "Bulk copy mailboxes"
        }
        "A" {
            $global:SuspendWhenReadyToComplete = $false
            $global:MigratieType = "All mailboxes"
        }
    }
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Selected migration type: " -ForegroundColor Gray; Write-Host $MigratieType -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewLine "  Would you like to manually ("; Write-Host -NoNewLine "S" -ForegroundColor Yellow; Write-Host -NoNewLine ")elect mailbox(es) or like to use a CSV-file to ("; Write-Host -NoNewLine "I" -ForegroundColor Yellow; Write-Host -NoNewLine ")mport mailbox(es)? "
        $Choice = Read-Host
        $Input = @("I","S") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
        }
    } Until ($Input)
    # ==========
    # Finalizing
    # ==========
    Switch ($Choice) {
        "I" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
        "S" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 3]
        }
    }
}