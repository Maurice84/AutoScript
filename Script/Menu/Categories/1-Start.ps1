Function Script-Menu-Categories-1-Start {
    # ===============
    # Initializations
    # ===============
    If ($InputArgs0 -eq "Updated") { EXIT }
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name $global:MainTitel
    Script-Module-ClearVariables
    $global:MenuNameStart = ($MyInvocation.MyCommand).Name
    $MenuOptions = $global:Functions | Where-Object { $_.Subcategory -eq "Menu" -AND $_.Task -eq "Subcategories" }
    Write-Host "$global:Greetings Please select 1 of the categories beneath to execute tasks like create, modify, remove or export/import:"
    Write-Host
    # -------------------------------------------
    # Here we iterate through every menu function
    # -------------------------------------------
    $Counter = 1
    $SwitchCase = $null
    foreach ($MenuOption in $MenuOptions) {
        $Name = $MenuOption.Name
        $Subcategory = $MenuOption.SubTask
        Write-Host ("  " + $Counter + ". " + $Subcategory) -ForegroundColor Yellow
        $Subcategories = ($global:Functions | Where-Object { $_.Category -eq $Subcategory } | Select-Object Subcategory -Unique).Subcategory
        # ------------------------------------------------------------------------------------------------------------
        # If a Subcategory description (subcategory converted with spaces) is found we then use this as output instead
        # ------------------------------------------------------------------------------------------------------------
        foreach ($Subcategory in $Subcategories) {
            Write-Host ("       - " + $Subcategory) -ForegroundColor Gray
        }
        # -------------------------------------------------------------------------------------------------------
        # Here we append a Switch statement line to automate the values for the statement after the foreach loop,
        # we use the counter as input value and add the function name to execute with.
        # -------------------------------------------------------------------------------------------------------
        $SwitchCase += "'$Counter' {Invoke-Expression -Command '$Name'};"
        $Counter++
        Write-Host
    }
    Write-Host "  X. Exit script" -ForegroundColor Red
    $SwitchCase += "'X' {EXIT}"
    Write-Host
    Do {
        [string]$Choice = Read-Host "Please select a category"
        [array]$MenuInput += "X"
        for ($Count = 1; $Count -le $Counter; $Count++) {
            $MenuInput += [string]$Count
        }
        if ($MenuInput -notcontains $Choice) {
            Write-Host "Please use the letter/numbers from above as input" -ForegroundColor Red; Write-Host
            $MenuInput = $null
        }
    } Until ($MenuInput)
    # ------------------------------------------------------------------------------------------------------------------------------------
    # Creating the Switch statement by adding the start of the statement including the choice of the user, then we append the switch cases
    # and finally close the statement. After that we convert this String to a ScriptBlock and execute it.
    # ------------------------------------------------------------------------------------------------------------------------------------
    $Menu = "Switch ('$Choice') {" + $SwitchCase + "}"
    $Menu = [scriptblock]::Create($Menu)
    $Menu.Invoke()
}