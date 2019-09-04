Function Maintenance-WindowsServer-Overview-Applications-2-InputApplications {
    # ============
    # Declarations
    # ============
    $global:Applications = @()
    $Task = "Manual entered application file paths to scan"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "  The Uninstall registry-keys of the server name(s) will be scanned using PowerShell Remoting." -ForegroundColor Cyan
    Write-Host "  Applications that were placed manually (like portable apps) on the server are therefore not detected." -ForegroundColor Cyan
    Write-Host "  To circumvent this you can optionally enter the name of the application and it's file path." -ForegroundColor Cyan
    Write-Host
    Do {
        Write-Host -NoNewline "  Would you like to enter manual file paths to scan? (Y/N): " -ForegroundColor Yellow
        $InputChoice = Read-Host
        $InputKey = @("Y"; "N") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "  Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "Y" {
            Do {
                Do {
                    Write-Host -NoNewLine "  > Please enter a name for the application (i.e. Notepad): "
                    $AppName = Read-Host
                } Until ($AppName)
                Do {
                    Write-Host -NoNewline "  > Please enter the file path (i.e. C:\Windows\notepad.exe): "
                    $AppFilePath = Read-Host
                    $InputKey = $AppFilePath
                    If (!$InputKey -OR $InputKey -notlike "*:\*") {
                        Write-Host "    This is not a valid file path" -ForegroundColor Red
                        $InputKey = $null
                    }
                } Until ($InputKey)
                # ------------------------------------------------------
                # Adding the application name and file path to the array
                # ------------------------------------------------------
                $Application = New-Object -TypeName System.Object
                $Application | Add-Member -MemberType NoteProperty -Name "Name" -Value $AppName
                $Application | Add-Member -MemberType NoteProperty -Name "FilePath" -Value $AppFilePath
                $global:Applications += $Application
                # ------------------
                # Repeat if required
                # ------------------
                Write-Host
                Do {
                    Write-Host -NoNewLine "  Would you like to add another application and file path? (Y/N): " -ForegroundColor Yellow
                    $InputChoice = Read-Host
                    $InputKey = @("Y"; "N") -contains $InputChoice
                    If (!$InputKey) {
                        Write-Host "  Please use the letters above as input" -ForegroundColor Red
                    }
                } Until ($InputKey)
                Switch ($InputChoice) {
                    "N" {
                        $InputDone = $true
                    }
                }
            } Until ($InputDone)
        }
    }
    # ==========
    # Finalizing
    # ==========
    if ($global:Applications) {
        Set-Variable $global:VarHeaderName -Value ([string]$global:Applications.Count + " application(s)")
    }
    else {
        Set-Variable $global:VarHeaderName -Value "None"
    }
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}