Function Storage-FileSystem-Delete-Files-6-Start {
    # ============
    # Declarations
    # ============
    [int]$Counter = 0
    [int]$CountFailed = 0
    [int]$CountSuccess = 0
    [string]$Exclude1 = $global:Drive + "System Volume Information"
    [string]$Exclude2 = $global:Drive + "Windows"
    [string]$ExportHeader = "Status;DeletedFile"
    [string]$Log = "C:\FileDeleteLog-" + (Get-Date -ForegroundColor yyyyMMdd_HHmmss) + ".csv"
    [string]$Task = "Indexing through $global:Choice with search filter: $global:SearchFilter"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Add-Content ($Log) $ExportHeader
    If ($global:Server -ne $env:computername) {
        $global:Drive = "\\" + $global:Server.ToUpper() + "\" + $global:Drive[0] + "$"
    }
    # ----------------------------
    # Iterate through every object
    # ----------------------------
    Get-ChildItem $global:Drive -Force -Recurse -ErrorAction SilentlyContinue | Where-Object {`
            $_.FullName -notlike "$Exclude1*" -AND `
            $_.FullName -notlike "$Exclude2*" -AND `
            $_.PSIsContainer -eq $global:Folder } | `
        ForEach-Object {
        If ((($_.FullName).Length + 15) -gt $global:MaxRowLength) {
            $Line = ($_.FullName).Substring(0, $global:MaxRowLength - 15) + "..."
        }
        Else {
            $Line = $_.FullName
        }
        If ($Counter -eq $global:MaxRowCount + 10) {
            $Counter = 0
            Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        }
        If ($_.FullName -like "*$global:SearchFilter*") {
            Write-Host -NoNewLine " - FOUND: " -ForegroundColor Red; Write-Host $Line -ForegroundColor Yellow
            If (Test-Path $_.FullName) {
                If ($global:Choice -eq "folders") {
                    & cmd.exe /c RD /S /Q $_.FullName
                    If (!$?) {
                        $Problem = $true
                    }
                }
                Else {
                    Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                    If (!$?) {
                        $Problem = $true
                    }
                }
                If ($Problem) {
                    Write-Host -NoNewLine " - ERROR: An error occurred deleting, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    $CountFailed++
                    $Problem = $null
                }
                Else {
                    Write-Host " - OK: Successfully deleted" -ForegroundColor Green
                    $CountSuccess++
                }
                If ($global:PauseProcessing -eq "Yes") {
                    Write-Host -NoNewLine " - PAUSE: Use S to skip this message in the future. If not, press Enter to continue: " -ForegroundColor Yellow
                    $InputKey = Read-Host
                    If ($InputKey -eq "s") {
                        $global:PauseProcessing = $null
                    }
                    Write-Host
                }
            }
        }
        Else {
            Write-Host -NoNewLine " -" $Line -ForegroundColor Gray; Write-Host
        }
        $Counter++
    }
    Write-Host
    # --------------------------------------------
    # Displaying the successful and failed objects
    # --------------------------------------------
    if ($CountSuccess -ge 1) {
        Write-Host -NoNewLine "Status: $CountSuccess $global:Choice restored and can be found in." -ForegroundColor Green; Write-Host $Log -ForegroundColor Yellow
    }
    If ($CountFailed -ge 1) {
        Write-Host -NoNewLine "Please Note: $CountFailed $global:Choice could not be deleted, please investigate " -ForegroundColor Red; Write-Host $Log -ForegroundColor Yellow
    }
    If (($CountFailed + $CountSuccess) -eq 0) {
        Write-Host "There are no $global:Choice detected with entered search filter" -ForegroundColor Green
    }
    # ---------------------------
    # Repeat process if requested
    # ---------------------------
    Do {
        Write-Host
        Write-Host -NoNewLine "Would you like to search on another (network)drive? Use "; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N): "
        [string]$InputChoice = Read-Host
        $InputKey = @("Y", "N") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "Y" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 3]
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}