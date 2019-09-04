Function Script-Module-ClearVariables {
    Get-Variable | Where-Object { $global:StartupVariables -notcontains $_.Name } | Clear-Variable -Scope Global -ErrorAction SilentlyContinue
    Get-Variable | Where-Object { $global:StartupVariables -notcontains $_.Name } | Clear-Variable -Scope Local -ErrorAction SilentlyContinue
    Get-Variable | Where-Object { $global:StartupVariables -notcontains $_.Name } | Remove-Variable -Scope Global -ErrorAction SilentlyContinue
    Get-Variable | Where-Object { $global:StartupVariables -notcontains $_.Name } | Remove-Variable -Scope Local -ErrorAction SilentlyContinue
}