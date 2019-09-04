Function Storage-FileSystem-Modify-FilePermissions-2-InputPath {
    # ============
    # Declarations
    # ============
    $global:Applications = @()
    $Task = "Select file of folder path to set group permissions"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "  > Please enter the file or folder path to use for the group to set permissions: "
        $global:Path = Read-Host
        $InputKey = $global:Path
        If (!$InputKey -OR $InputKey -notlike "*:\*") {
            Write-Host "    Please enter the file path" -ForegroundColor Red
            $InputKey = $null
        }
    } Until ($InputKey)
    $global:Applications = New-Object PSObject -Property @{Group = $global:Group; Path = $global:Path}
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Path
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}