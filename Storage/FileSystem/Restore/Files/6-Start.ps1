Function Storage-FileSystem-Restore-5-Start {
    # ============
    # Declarations
    # ============
    [int]$Counter = 0
    [int]$CountFailed = 0
    [int]$CountSuccess = 0
    [string]$Exclude1 = $global:DriveSource + "System Volume Information"
    [string]$Exclude2 = $global:DriveSource + "Windows"
    [string]$ExportHeader = "Status;OldFile;RecoveredFile"
    [string]$Log = "C:\FileRestoreLog-" + (Get-Date -ForegroundColor yyyyMMdd_HHmmss) + ".csv"
    [string]$Task = "Indexing through files with search filter: $global:SearchFilter"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Add-Content ($Log) $ExportHeader
    # ----------------------------
    # Iterate through every object
    # ----------------------------
    Get-ChildItem $DriveSource -Force -Recurse -ErrorAction SilentlyContinue | Where-Object { `
            $_.FullName -notlike "$Exclude1*" -AND `
            $_.FullName -notlike "$Exclude2*" } | `
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
        If ($_.FullName -like "*$global:SearchFilter") {
            Write-Host -NoNewLine " - FOUND: " -ForegroundColor Red; Write-Host $_.FullName -ForegroundColor Yellow
            $Destination = $_.FullName.Replace($global:DriveSource, $global:DriveTarget)
            If (!(Test-Path $Destination)) {
                $SourceBacktick = $Source.Replace('[', '`[')
                $SourceBacktick = $SourceBacktick.Replace(']', '`]')
                Copy-Item $SourceBacktick $Destination
                If ($?) {
                    Write-Host -NoNewLine " - RESTORED: " -ForegroundColor Green; Write-Host $_.FullName -ForegroundColor Yellow
                    $ExportLine = "Restored;" + $_.FullName + ";" + $Source
                    Add-Content ($Log) $ExportLine
                    $CountSuccess++
                }
                Else {
                    $Problem = $true
                    $CountFailed++
                }
            } else {
                    Write-Host -NoNewLine " - ALREADY EXISTS: " -ForegroundColor Green; Write-Host $_.FullName -ForegroundColor Yellow
                    $ExportLine = "Already Exists;" + $_.FullName + ";" + $Source
                    Add-Content ($Log) $ExportLine
            }
            If ($Problem) {
                Write-Host -NoNewLine " - ERROR: An error occurred deleting, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                $CountFailed++
                $Problem = $null
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
        Else {
            Write-Host -NoNewLine " -" $Line -ForegroundColor Gray; Write-Host
        }
        $Counter++
    }
    # --------------------------------------------
    # Displaying the successful and failed objects
    # --------------------------------------------
    if ($CountSuccess -ge 1) {
        Write-Host -NoNewLine "Status: $CountSuccess file(s) restored and can be found in." -ForegroundColor Green; Write-Host $Log -ForegroundColor Yellow
    }
    If ($CountFailed -ge 1) {
        Write-Host -NoNewLine "Please Note: $CountFailed file(s) could not be deleted, please investigate " -ForegroundColor Red; Write-Host $Log -ForegroundColor Yellow
    }
    If (($CountFailed + $CountSuccess) -eq 0) {
        Write-Host "There are no files detected with entered search filter" -ForegroundColor Green
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}