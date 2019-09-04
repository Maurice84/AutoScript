Function Script-Menu-Subcategories-5-Software {
    # ============
    # Declarations
    # ============
    $global:MenuNameCategory = ($MyInvocation.MyCommand).Name
    $Category = ($global:Functions | Where-Object { $_.Name -eq $global:MenuNameCategory }).SubTask
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name $global:MainTitel
    Script-Module-ClearVariables
    Script-Function-MenuCategory -Category $Category
}