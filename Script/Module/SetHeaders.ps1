Function Script-Module-SetHeaders {
    param (
        [boolean]$DisplayHeaders = $true,
        [string]$Name,
        [string]$CurrentTask
    )
    [int]$StepCount = $null
    [string]$Task = $null
    [string]$SubTask = $null
    [string]$Category = $Name.Split('-')[0]
    [string]$Subcategory = $Name.Split('-')[1]
    [string]$Name = $Name.Replace($Category + '-', '')
    [string]$Name = $Name.Replace($Subcategory + '-', '')
    [array]$TaskElements = $Name.Split('-')
    # =========
    # Execution
    # =========
    # ------------------------------------------------------------------------------------------------------
    # Clearing the screen by using 'Clear' instead of 'Clear-Host' due to PowerShell backwards compatibility
    # ------------------------------------------------------------------------------------------------------
    Clear
    # -----------------------------------------------------
    # Iterate through the elements of the Name (split by -)
    # -----------------------------------------------------
    if ($TaskElements.Count -gt 1) {
        [int]$Counter = 1
        foreach ($Part in $TaskElements) {
            if ([string]$Part[0] -as [int]) {
                $StepCount = [int]$Part
            }
            if ($StepCount) {
                $SubTask = [string]$Part
            }
            else {
                if ($Counter -ne $TaskElements.Count) {
                    $Task += [string]$Part + " "
                    $Counter++
                }
                else {
                    $SubTask = [string]$Part
                }
            }
        }
        $global:Title = $Category + " > " + $Subcategory + " > " + $Task
        if ($StepCount) {
            $global:Title += "> Step " + [string]$StepCount + ": " + $SubTask
        }
        else {
            $global:Title += "> " + $SubTask
        }
    }
    else {
        $global:Title = $Name
    }
    # -----------------
    # Display the title
    # -----------------
    $global:Title = "Maurice AutoScript - " + $global:Title
    Write-Host ("=" * $global:Title.Length) -ForegroundColor Gray
    Write-Host $global:Title -ForegroundColor Gray
    Write-Host ("=" * $global:Title.Length) -ForegroundColor Gray
    Write-Host
    # --------------------------------------------
    # Display Headers when going through the tasks
    # --------------------------------------------
    if ($DisplayHeaders -eq $true) {
        $global:VarHeaderName = "Header" + $Name.Split("-")[-2] + $Name.Split("-")[-1]
        $GetHeaders = Get-Variable | Where-Object { $_.Name -like "Header*" }
        # Write-Host $GetHeaders.Name -ForegroundColor Yellow
        if ($StepCount -gt 1) {
            $GetHeaders = $GetHeaders | Select-Object -First ($StepCount - 1)
        }
        else {
            $GetHeaders = $null
        }
        if ($GetHeaders) {
            foreach ($RetrievedHeader in $GetHeaders) {
                $RetrievedHeader = [scriptblock]::Create($RetrievedHeader.Value)
                $RetrievedHeader.Invoke()
            }
        }
    }
    # -------------------------------------------------------------
    # If a Task has been given, then display this after the headers
    # -------------------------------------------------------------
    if ($CurrentTask) {
        Write-Host ("- " + $CurrentTask + ":") -ForegroundColor Gray
        Write-Host
    }
}