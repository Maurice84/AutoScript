Function Maintenance-WindowsServer-Overview-Applications-4-Start {
    param (
        [array]$Applications,
        [array]$Servers
    )
    # ============
    # Declarations
    # ============
    $Counter = 1
    $ExportCSVFiles = @()
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host
    Write-Host "Scanning servers for applications" -ForegroundColor Magenta
    Write-Host " - Starting server jobs:" -ForegroundColor Yellow
    # ---------------------------
    # Iterate through each server
    # ---------------------------
    foreach ($Server in $Servers) {
        Write-Host -NoNewline ("   > Adding job to server " + $Counter + "/" + $Servers.Count + ": " + $Server + "...")
        $ExportCSVFile = "C:\AppScan-" + $Server + "_" + (Get-Date -f "yyyyMMdd-HHmm") + ".csv"
        $ExportCSVFiles += $ExportCSVFile
        Start-Job -Name $Server -ArgumentList $Applications, $ExportCSVFile, $Server -ScriptBlock {
            param(
                [array]$Applications,
                [string]$ExportCSVFile,
                [string]$Server
            )
            $Result = @()
            # -----------------------------------------------
            # Connecting to the remote registry of the server
            # -----------------------------------------------
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)
            # ----------------------------------
            # Indexing the Uninstall information
            # ----------------------------------
            $SubBranch = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            $RegistryKey = $Registry.OpenSubKey($SubBranch)
            foreach ($Key in $RegistryKey.GetSubKeyNames()) {
                $NewSubKey = $SubBranch + "\\" + $Key
                $ReadUninstall = $Registry.OpenSubKey($NewSubKey)
                if ($ReadUninstall.GetValue("DisplayName")) {
                    $Value = New-Object -TypeName System.Object
                    $Value | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
                    $Value | Add-Member -MemberType NoteProperty -Name "Name" -Value $ReadUninstall.GetValue("DisplayName")
                    if ($ReadUninstall.GetValue("DisplayVersion")) {
                        $Version = $ReadUninstall.GetValue("DisplayVersion")
                    }
                    else {
                        $Version = "None"
                    }
                    $Value | Add-Member -MemberType NoteProperty -Name "Version" -Value $Version
                    $Result += $Value
                }
            }
            $SubBranch = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            $RegistryKey = $Registry.OpenSubKey($SubBranch)
            foreach ($Key in $RegistryKey.GetSubKeyNames()) {
                $NewSubKey = $SubBranch + "\\" + $Key
                $ReadUninstall = $Registry.OpenSubKey($NewSubKey)
                if ($ReadUninstall.GetValue("DisplayName")) {
                    $Value = New-Object -TypeName System.Object
                    $Value | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
                    $Value | Add-Member -MemberType NoteProperty -Name "Name" -Value $ReadUninstall.GetValue("DisplayName")
                    if ($ReadUninstall.GetValue("DisplayVersion")) {
                        $Version = $ReadUninstall.GetValue("DisplayVersion")
                    }
                    else {
                        $Version = "None"
                    }
                    $Value | Add-Member -MemberType NoteProperty -Name "Version" -Value $Version
                    $Result += $Value
                }
            }
            # -------------------
            # Sorting unique Name
            # -------------------
            $Result = $Result | Sort-Object Name -Unique
            # ---------------------------------------------------
            # Iterate through every manual application if entered
            # ---------------------------------------------------
            foreach ($Application in $Applications) {
                $UNC = "\\" + $Server + "\" + ($Application.FilePath).Replace(":", "$")
                if (Test-Path $UNC) {
                    $Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($UNC).ProductVersion
                    if (!$Version) {
                        $Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($UNC).FileVersion
                        if (!$Version) {
                            $Version = "None"
                        }
                    }
                    $Value = New-Object -TypeName System.Object
                    $Value | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
                    $Value | Add-Member -MemberType NoteProperty -Name "Name" -Value $Application.Name
                    $Value | Add-Member -MemberType NoteProperty -Name "Version" -Value $Version
                    $Result += $Value
                }
            }
            # ------------------------
            # Indexing Windows Updates
            # ------------------------
            $Updates = Get-HotFix -ComputerName $Server | Select-Object HotFixID, Description | Sort-Object HotFixID
            foreach ($Update in $Updates) {
                $Value = New-Object -TypeName System.Object
                $Value | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
                $Value | Add-Member -MemberType NoteProperty -Name "Name" -Value ("Windows " + $Update.Description)
                $Value | Add-Member -MemberType NoteProperty -Name "Version" -Value $Update.HotFixID
                $Result += $Value
            }
            # ----------------------------------
            # Get the Windows Server information
            # ----------------------------------
            $OSInfo = Get-CimInstance Win32_OperatingSystem -ComputerName $Server | Select-Object Name, Version, ServicePackMajorVersion, BuildNumber, OSArchitecture, OperatingSystemSKU
            $OS = (($OSInfo).Name -Split '\|')[0]
            if ($OSInfo.ServicePackMajorVersion -ge 1) {
                $OS += " SP" + $OSInfo.ServicePackMajorVersion
            }
            $OS += " " + $OSInfo.OSArchitecture
            $Value = New-Object -TypeName System.Object
            $Value | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
            $Value | Add-Member -MemberType NoteProperty -Name "Name" -Value $OS
            $Value | Add-Member -MemberType NoteProperty -Name "Version" -Value $OSInfo.Version
            $Result += $Value
            # ----------------
            # Exporting to CSV
            # ----------------
            Remove-Item -Path ("C:\AppScan-" + $Server + "*.csv") -ErrorAction SilentlyContinue -Force -Confirm:$false
            $Result | Sort-Object Name, Version | Export-Csv -Path $ExportCSVFile -NoTypeInformation -Delimiter ";" -Force
        } | Out-Null
        # ---------------------------
        # Check if the job is running
        # ---------------------------
        if ((Get-Job -Name $Server).State -eq "Running") {
            Write-Host " OK!" -ForegroundColor Green
        }
        $Counter++
    }
    # ---------------------------------
    # Wait until all jobs are completed
    # ---------------------------------
    Write-Host
    Write-Host -NoNewLine " - Waiting for jobs to finish:" -ForegroundColor Yellow
    $PreviousMessage = $null
    do {
        $Jobs = Get-Job
        $Completed = $Jobs | Where-Object { $_.State -eq "Completed" }
        $Running = $Jobs | Where-Object { $_.State -eq "Running" }
        $Failed = $Jobs | Where-Object { $_.State -ne "Running" -AND $_.State -ne "Completed" }
        $MessagePrefix = "   > " + (Get-Date -Format "HH:mm:ss") + ": "
        if (!$Running) {
            $Message = "Done!"
            Write-Host
            Write-Host ($MessagePrefix + $Message) -ForegroundColor Green
        }
        else {
            $Message = ("In Progress: " + [string]$Running.Count)
            if ($Completed) {
                $Message += (", Successful: " + [string]$Completed.Count)
            }
            if ($Failed) {
                $Message += (", Failed: " + [string]$Failed.Count)
            }
            $Message += "..."
            if ($Message -ne $PreviousMessage) {
                Write-Host
                Write-Host -NoNewLine ($MessagePrefix + $Message) -ForegroundColor Cyan
                $PreviousMessage = $Message
            }
            Start-Sleep -Seconds 3
        }
    } until (!$Running)
    # -------
    # Cleanup
    # -------
    Get-Job | Remove-Job
    # -----------------------------------------------------
    # Combine the CSV-files if there's more than 1 CSV-file
    # -----------------------------------------------------
    if ($ExportCSVFiles.Count -gt 1) {
        Write-Host
        Write-Host "Combine CSV-files" -ForegroundColor Magenta
        # -------------------
        # Index all CSV-files
        # -------------------
        Write-Host -NoNewLine " - Indexing all CSV-files..."
        $Array = @()
        foreach ($CSVFile in $ExportCSVFiles) {
            $Array += Import-CSV -Path $CSVFile -Delimiter ";"
        }
        if ($Array) {
            $Array = $Array | Sort-Object Name, Version, Server
            Write-Host " OK!" -ForegroundColor Green
        }
        else {
            Write-Host " Error!" -ForegroundColor Red
            Pause
            Exit
        }
        # ---------------------------------------
        # Index all unique servers for the header
        # ---------------------------------------
        Write-Host -NoNewLine " - Indexing all servers..."
        $AllServers = ($Array | Select-Object Server -Unique | Sort-Object Server).Server
        $ExportHeader = '"Name";"Version"'
        foreach ($Server in $AllServers) {
            $ExportHeader += ';"' + $Server + '"'
        }
        Remove-Item -Path ("C:\AppScanMatrix_*.csv") -ErrorAction SilentlyContinue -Force -Confirm:$false
        $ExportCSVFile = ("C:\AppScanMatrix_" + (Get-Date -f "yyyyMMdd-HHmm") + ".csv")
        Add-Content $ExportCSVFile $ExportHeader
        if (Test-Path $ExportCSVFile) {
            Write-Host " OK!" -ForegroundColor Green
        }
        else {
            Write-Host " Error!" -ForegroundColor Red
            Pause
            Exit
        }
        # --------------------------------------------------
        # Go through all rows from the combined CSVs (Array)
        # --------------------------------------------------
        $Name = $null
        $Version = $null
        # -----------------------------------
        # Loop through every row in the Array
        # -----------------------------------
        Write-Host -NoNewLine " - Indexing and exporting to a combined matrix CSV-file..."
        foreach ($Row in $Array) {
            $RowVersion = ($Row.Version).Trim()
            # ---------------------------------------------------------------------------------------------------------------------
            # Check if the application has already been parsed by first checking the Name. If this is the same as the previous row,
            # then check of the Version is the same. If this is not the same, then proceed. If not then skip.
            # ---------------------------------------------------------------------------------------------------------------------
            if ($Name -eq $Row.Name) {
                if ($Version -eq $RowVersion) {
                    $Skip = $true
                }
            }
            if (!$Skip) {
                $Name = $Row.Name
                $Version = ($Row.Version).Trim()
                $ExportValues = '"' + $Name + '";"' + $Version + '"'
                $Servers = ($Array | Where-Object { $_.Name -eq $Name -AND $_.Version -eq $Version } | Select-Object Server | Sort-Object Server).Server
                foreach ($Server in $AllServers) {
                    If ($Servers -contains $Server) {
                        $ExportValues = $ExportValues + ';"X"'
                    }
                    Else {
                        $ExportValues = $ExportValues + ';" "'
                    }
                }
                Add-Content $ExportCSVFile $ExportValues
            }
            $Skip = $null
        }
        if (Test-Path $ExportCSVFile) {
            Write-Host " OK!" -ForegroundColor Green
        }
        else {
            Write-Host " Error!" -ForegroundColor Red
            Pause
            Exit
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}