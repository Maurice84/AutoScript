Function Messaging-Exchange-Export-Mailboxes-2-SetDate {
    # ============
    # Declarations
    # ============
    $Task = "Select export start-date and end-date"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "- Would you like to export all mail-items or like to select a start-date and end-date?"
    Write-Host
    Write-Host "  1. Export all mail-items" -ForegroundColor Yellow
    Write-Host "  2. Export mail-items using start-date and end-date" -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Select an option or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
        [string]$Choice = Read-Host
        $Input = @("1", "2", "X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letter/numbers from above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "1" {
            $global:StartDate = Get-Date 01-01-1900
            $global:EndDate = Get-Date
        }
        "2" {
            If (!$RemoteExchange) {
                Do {
                    Write-Host
                    Write-Host -NoNewLine "  Would you like to enter a start-date? (Use N to only export today's items) (Y/N): "
                    [string]$Choice = Read-Host
                    $Input = @("Y", "N") -contains $Choice
                    If (!$Input) {
                        Write-Host "  Please use the letters above as input" -ForegroundColor Red
                    }
                } Until ($Input)
                Switch ($Choice) {
                    "Y" {
                        Do {
                            $global:StartDate = Read-Host "  Please enter a start-date (DD-MM-YYYY)"
                            If (!($StartDate -as [datetime])) {
                                Write-Host "  Please enter correct datetime syntax" -ForegroundColor Red
                            }
                            Else {
                                If (($StartDate -as [datetime]) -gt (Get-Date)) {
                                    Write-Host "  Start-date cannot be in the future..." -ForegroundColor Red
                                }
                                Else {
                                    $global:StartDate = $StartDate -as [datetime]
                                }
                            }
                        } While ($StartDate -isnot [datetime])
                    }
                    "N" {
                        $global:StartDate = Get-Date
                    }
                }
                Do {
                    Write-Host -NoNewLine "  Would you like to enter an end-date? (Y/N): "
                    [string]$Choice = Read-Host
                    $Input = @("Y", "N") -contains $Choice
                    If (!$Input) {
                        Write-Host "  Please use the letters above as input" -ForegroundColor Red
                    }
                } Until ($Input)
                Switch ($Choice) {
                    "Y" {
                        Do {    
                            $global:EndDate = Read-Host "  Please enter an end-date (DD-MM-YYYY)"
                            If (!($EndDate -as [datetime])) {
                                Write-Host "  Please enter correct datetime syntax" -ForegroundColor Red
                            }
                            Else {
                                $global:EndDate = $EndDate -as [datetime]
                            }
                        } While ($EndDate -isnot [datetime])
                    }
                    "N" {
                        $global:EndDate = Get-Date
                    }
                }
            }
            Else {
                Write-Host "  Unfortunately it's not (yet) possible to use a custom date export using Remote Exchange. Please execute this script on the Exchange server itself and repeat the task." -ForegroundColor Red
                $Pause.Invoke()
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name

            }
        }
        "X" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
        }
    }
    If ($StartDate.ToShortDateString() -eq (Get-Date).ToShortDateString()) {
        $global:StartDateFormat = "vandaag"
    }
    Else {
        $global:StartDateFormat = $StartDate.ToString("dd-MMM-yyyy")
    }
    If ($EndDate.ToShortDateString() -eq (Get-Date).ToShortDateString()) {
        $global:EndDateFormat = "vandaag"
    }
    Else {
        $global:EndDateFormat = $EndDate.ToString("dd-MMM-yyyy")
    }

    If (($StartDateFormat -eq "vandaag") -AND ($EndDateFormat -eq "vandaag")) {
        $global:Periode = "Alleen vandaag"
    }
    ElseIf (($StartDate.ToShortDateString() -eq (Get-Date 01-01-1900).ToShortDateString()) -AND ($EndDateFormat -eq "vandaag")) {
        $global:Periode = "Alle e-mails"
    }
    Else {
        $global:Periode = $StartDateFormat + " t/m " + $EndDateFormat
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Periode
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}