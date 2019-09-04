Function Script-Function-Confirmation {
    param(
        [string]$Arguments,
        [string]$Name
    )
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name $Name
    Write-Host
    Do {
        Write-Host -NoNewLine "Please confirm if above is correct. Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N/X): "
        [string]$InputChoice = Read-Host
        $InputKey = @("Y", "N", "X") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "Y" {
            Invoke-Expression -Command ($FunctionTaskNames[[int]$FunctionTaskNames.IndexOf($Name) + 1] + $Arguments)
        }
        "N" {
            Script-Module-SetHeaders -DisplayHeaders $false -Name $Name
            $GetHeaders = Get-Variable | Where-Object { $_.Name -like "Header*" }
            $Counter = 1
            $SwitchCase = $null
            foreach ($RetrievedHeader in $GetHeaders) {
                $RetrievedHeader = [scriptblock]::Create($RetrievedHeader.Value)
                Write-Host -NoNewLine " $Counter " -ForegroundColor Red
                $RetrievedHeader.Invoke()
                $TaskName = $FunctionTaskNames | Where-Object { $_ -like "*-$Counter-*" }
                $SwitchCase += "'$Counter' {Invoke-Expression -Command '$TaskName'};"
                $Counter++
            }
            $SwitchCase += "'X' {Invoke-Expression -Command '$Name'}"
            Write-Host
            Do {
                Write-Host -NoNewLine "Please select a number to modify. Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
                [string]$InputChoice = Read-Host
                $InputKey = @()
                $InputKey += "X"
                for ($Count = 1; $Count -lt $Counter; $Count++) {
                    $InputKey += [string]$Count
                }
                if ($InputKey -notcontains $InputChoice) {
                    Write-Host "Please use the letter/numbers from above as input" -ForegroundColor Red; Write-Host
                    $InputKey = $null
                }
            } Until ($InputKey)
            # ------------------------------------------------------------------------------------------------------------------------------------
            # Creating the Switch statement by adding the start of the statement including the choice of the user, then we append the switch cases
            # and finally close the statement. After that we convert this String to a ScriptBlock and execute it.
            # ------------------------------------------------------------------------------------------------------------------------------------
            $Menu = "Switch ('$InputChoice') {" + $SwitchCase + "}"
            $Menu = [scriptblock]::Create($Menu)
            $Menu.Invoke()
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
}