Function Script-Function-MenuCategory {
    param (
        [string]$Category
    )
    # =========
    # Execution
    # =========
    Write-Host ("Alright " + $global:CurrentUser + ", please select a task for the category " + $Category + ":")
    Write-Host
    # -------------------------------------------
    # Here we iterate through every menu function
    # -------------------------------------------
    $Counter = 1
    $SwitchCase = $null
    $Subcategories = ($global:Functions | Where-Object { $_.Category -eq $Category } | Select-Object Subcategory -Unique).Subcategory
    foreach ($Subcategory in $Subcategories) {
        Write-Host ("  " + $Subcategory) -ForegroundColor Cyan
        $SubcategoryFunctions = $global:Functions | Where-Object { $_.Category -eq $Category -AND $_.Subcategory -eq $Subcategory }
        $MenuOptions = $SubcategoryFunctions | Where-Object { $_.StepCount -eq "1" }
        foreach ($MenuOption in $MenuOptions) {
            $Name = $MenuOption.Name
            $Subject = $MenuOption.Subject
            $Task = $MenuOption.Task
            $SubTask = $MenuOption.SubTask
            $DisplayName = $Task + " " + $Subject
            if ($SubTask) {
                $DisplayName += " " + $SubTask
            }
            # ------------------------------------------------------------------------------------
            # Output the line with the correct outlining of spaces with a menu list larger than 10
            # ------------------------------------------------------------------------------------
            if ($Counter -lt 10) {
                $Spaces = 5
            }
            else {
                $Spaces = 4
            }
            Write-Host ((" " * $Spaces) + $Counter + ". " + $DisplayName) -ForegroundColor Yellow
            # -------------------------------------------------------------------------------------------------------
            # Here we append a Switch statement line to automate the values for the statement after the foreach loop,
            # we use the counter as input value and add the function name to execute with.
            # -------------------------------------------------------------------------------------------------------
            $TaskNames = ($SubcategoryFunctions | Where-Object { $_.Subject -eq $Subject -AND $_.Task -eq $Task }).Name -join "','"
            $SwitchCase += "'$Counter' {`$global:FunctionTaskNames = @(); `$global:FunctionTaskNames = '$TaskNames'; Invoke-Expression -Command '$Name'};"
            $Counter++
        }
        Write-Host
    }
    $SwitchCase += "'X' {Get-Variable -Exclude '$global:StartupVariables' | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command '$global:MenuNameStart'}"
    Do {
        Write-Host -NoNewline "Please select a task or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the main menu: "
        [string]$Choice = Read-Host
        [array]$MenuInput += "X"
        for ($Count = 1; $Count -lt $Counter; $Count++) {
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
