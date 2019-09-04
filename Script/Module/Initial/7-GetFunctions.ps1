function Script-Module-Initial-7-GetFunctions {
    # ============
    # Declarations
    # ============
    $LoadedFunctions = Get-ChildItem function:\ | Where-Object { $SystemFunctions -notcontains $_ }
    $ExcludedWordsWithSpaces = @()
    $ExcludedWordsWithSpaces = "AutoComplete", "SyncBackPro"
    $Descriptions = @()
    $global:Functions = @()
    # =========
    # Execution
    # =========
    $Item = New-Object PSObject
    $Item | Add-Member -type NoteProperty -Name "IIS" -Value "Internet Information Services"
    $Item | Add-Member -type NoteProperty -Name "OWA" -Value "Outlook Web Access"
    $Item | Add-Member -type NoteProperty -Name "RDSH" -Value "Remote Desktop Session Host"
    $Descriptions += $Item
    # -------------------------------------
    # Iterate through every loaded function
    # -------------------------------------
    foreach ($Function in $LoadedFunctions) {
        $Name = $Function.Name
        if ($Name -like "Script-*" -OR $Name -like "StartFile_*") {
            $Type = "Script"
        }
        else {
            $Type = "Function"
        }
        $Elements = $Name.Split('-')
        $ElementsWithSpaces = @()
        foreach ($Word in $Elements) {
            if ($ExcludedWordsWithSpaces -notcontains $Word) {
                $WordWithSpaces = $null
                $Counter = 0
                foreach ($Letter in $Word.toCharArray()) {
                    $NextLetter = $Word[$Counter + 1]
                    if ($Letter -cmatch "[a-z]") {
                        $WordWithSpaces += $Letter
                        if ($NextLetter -cmatch "[A-Z]") {
                            $WordWithSpaces += " "
                        }
                    }
                    else {
                        $WordWithSpaces += $Letter
                        if ($NextLetter -cmatch "[A-Z]") {
                            $NextNextLetter = $Word[$Counter + 2]
                            if ($NextNextLetter -cmatch "[a-z]") {
                                $WordWithSpaces += " "
                            }
                        }
                    }
                    $Counter++
                }
                $ElementsWithSpaces += $WordWithSpaces
            }
            else {
                $ElementsWithSpaces += $Word
            }
        }
        $Category = $ElementsWithSpaces[0]
        $Subcategory = $ElementsWithSpaces[1]
        if ($Descriptions.$Subcategory) {
            $Subcategory = $Descriptions.$Subcategory
        }
        $Task = $ElementsWithSpaces[2]
        $Subject = $ElementsWithSpaces[3]
        if ($Descriptions.$Subject) {
            $Subject = $Descriptions.$Subject
        }
        $SubTask = $null
        $StepCount = $null
        $StepTask = $null
        $Counter = 1
        foreach ($Element in $ElementsWithSpaces[4..($ElementsWithSpaces.Count)]) {
            if ([string]$Element[0] -as [int]) {
                $StepCount = $Element
            }
            if ($StepCount) {
                $StepTask = [string]$Element
            }
            else {
                if ($Counter -ne $Elements.Count) {
                    $SubTask += [string]$Element + " "
                    $Counter++
                }
                else {
                    $StepTask = [string]$Element
                }
            }
        }
        if ($SubTask) {
            $SubTask = $SubTask.TrimEnd()
        }
        $Properties = @{`
            Name        = $Name; `
            Type        = $Type; `
            Category    = $Category; `
            Subcategory = $Subcategory; `
            Task        = $Task; `
            Subject     = $Subject; `
            SubTask     = $SubTask; `
            StepCount   = $StepCount; `
            StepTask    = $StepTask;
        }
        $Object = New-Object PSObject -Property $Properties
        $global:Functions += $Object
    }
}