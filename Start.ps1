<#
.SYNOPSIS
    Execute or Merge the AutoScript using the separate function files
.DESCRIPTION
    This script will execute or merge the AutoScript using the separate function files in the subfolders
.INPUTS
    Leave empty or use 'Merge' as parameter to merge all separate function files to an all-in-one script file
.OUTPUTS
    None
.NOTES  
    Version:        1.0
    Author:         Maurice Heikens
    Creation Date:  2019-04-18
    Last Edited:    2019-04-18
    Purpose/Change: Initial script development
.EXAMPLE
    .\Start.ps1
#>

# ===============
# Initializations
# ===============
$InputMerge = $args[0]
if (!$InputMerge) {
    $InputMerge = "None"
}
#$InputMerge = "Merge"
$BypassUAC = $args[1]
# ------------------------------
# Run as Administrator if needed
# ------------------------------
if (!$BypassUAC) {
    Start-Process "powershell.exe" -ArgumentList "-File ""$PSScriptRoot\Start.ps1""","None","Admin" -Verb RunAs
    EXIT
}
# --------------------------------------------
# Get and execute the required startup content
# --------------------------------------------
if ($InputMerge -ne "Merge") {
    Invoke-Expression -Command ". '$PSScriptRoot\Script\Init.ps1'"
}

# ============
# Declarations
# ============
# ----------------------
# Setting the Basefolder
# ----------------------
$BaseFolder = $PSScriptRoot
Set-Location -Path "$BaseFolder"
if ($InputMerge -eq "Merge") {
    $AutoScriptName = "Maurice-AutoScript-Merged"
    $AutoScriptFile = "$BaseFolder" + "\_" + $AutoScriptName + "_" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".ps1"
}

# =========
# Functions
# =========
function StartFile_Index-FunctionFiles {
    param (
        [string]$Folder
    )
    # ----------------------------------------------------------------------------------------
    # First we index all separate function files. Folders named like 'Exclude' will be skipped
    # ----------------------------------------------------------------------------------------
    Write-Host -NoNewLine " - Now loading all separate function files..."
    $Files = (Get-ChildItem "$Folder" -Filter "*.ps1" -Recurse | Where-Object { $_.DirectoryName -ne "$Folder" -AND $_.DirectoryName -notlike "*Exclude*" }).FullName
    $Array = @()
    foreach ($File in $Files) {
        $Name = (($File.Replace("$Folder" + "\", "")).Split(".")[0]).Split("\") -join "-"
        # -------------------------------------------------------------------------------------
        # Detect if script file is a required startup function and in which order it should run
        # -------------------------------------------------------------------------------------
        if ($Name -like "Script*Menu*Start" -OR $Name -like "Script*Initial*") {
            $Type = "Startup"
        }
        else {
            $Type = "Normal"
        }
        $Properties = @{Name = $Name; FilePath = $File; Type = $Type }
        $Object = New-Object PSObject -Property $Properties
        $Array += $Object
    }
    if ($Array) {
        Write-Host " OK!" -ForegroundColor Green
        return $Array | Sort-Object Name
    }
    else {
        Write-Host " Error! Files could not be found" -ForegroundColor Red
    }
}
function StartFile_Merge-RemoveAndCreateFile {
    param (
        [string]$Folder,
        [string]$Name,
        [string]$File,
        [string]$InitFile
    )
    # --------------------------------------------------------------
    # First we remove all previous generated AutoScript Merged files
    # --------------------------------------------------------------    
    $OldMergedFiles = (Get-ChildItem $Folder | Where-Object { $_.Name -like "*$Name*" }).FullName
    foreach ($OldMergedFile in $OldMergedFiles) {
        Write-Host -NoNewLine (" - Remove old AutoScript merged file " + $OldMergedFile.Split("\")[-1] + "... ")
        Remove-Item $OldMergedFile -Force -Confirm:$false
        if ($?) {
            Write-Host "OK!" -ForegroundColor Green
        }
        else {
            Write-Host "Error! Old Merged AutoScript file could not be removed. Script will be terminated."
            Exit
        }
    }
    # -------------------------------------------------
    # Then we create a new empty AutoScript merged file
    # -------------------------------------------------
    Write-Host -NoNewLine " - Creating new AutoScript merged file... "
    New-Item -Path ($File) -ItemType File -Confirm:$false | Out-Null
    if (Test-Path $File) {
        Write-Host "OK!" -ForegroundColor Green
    }
    else {
        Write-Host "Error! AutoScript merged file could not be created. Script will be terminated."
        Exit
    }
    # ------------------------------------------------------------------------------------------------------
    # After that we retrieve required startup content and add it to the newly created AutoScript merged file
    # ------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine " - Adding required startup content to AutoScript merged file... "
    $Content = Get-Content $InitFile
    Add-Content -Path "$File" -Value $Content
    if ($?) {
        Write-Host "OK!" -ForegroundColor Green
    }
    else {
        Write-Host "Error! Could not add required startup content to AutoScript merged file. Script will be terminated."
        Exit
    }
}
function StartFile_Merge-AddContentToMergedFile {
    param (
        [string]$Name,
        [string]$Content,
        [string]$MergedFile,
        [string]$Type
    )
    # -------------------------------------------------------------------------------------------
    # Get the content of the detected script file and add it to the newly created AutoScript file
    # -------------------------------------------------------------------------------------------
    if ($Type -eq "File") {
        Write-Host -NoNewLine (" - Adding $Name... ")
        $Value = Get-Content $Content
    }
    if ($Type -eq "Startup") {
        Write-Host -NoNewLine (" - Adding startup function: $StartupFunction... ")
        $Value = $Content
    }
    Add-Content -Path "$MergedFile" -Value $Value
    if ($?) {
        Write-Host "OK!" -ForegroundColor Green
    }
    else {
        Write-Host "Error! Could not add content to AutoScript merged file. Script will be terminated."
        Exit
    }
}

# =========
# Execution
# =========
Clear-Host
Write-Host "Maurice AutoScript" -ForegroundColor Magenta
Write-Host (" - Script started at " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
# -----------------------------------------------------------------------------------------------------------
# Index all script files and gather the fullpath and directory name to declare a functionname for the script.
# -----------------------------------------------------------------------------------------------------------
$Functions = StartFile_Index-FunctionFiles -Folder "$BaseFolder"
# ----------------------------------------------------------------------------------------------
# If Merge has been set: Remove all AutoScript Merged files and create a new one, otherwise skip
# ----------------------------------------------------------------------------------------------
if ($InputMerge -eq "Merge") {
    $InitCode = ($Functions | Where-Object { $_.Name -eq "Script-Init" }).FilePath
    StartFile_Merge-RemoveAndCreateFile -Folder "$BaseFolder" -Name $AutoScriptName -File $AutoScriptFile -Init $InitCode
}
# ------------------------------------------------
# Loop through every detected function script file
# ------------------------------------------------
foreach ($Function in $Functions) {
    $Name = $Function.Name
    $FilePath = $Function.FilePath
    # -------------------------------------------------------------------------------------------------------------------
    # If Merge has been set: Get the content of the detected script file and add it to the newly created AutoScript file.
    # Otherwise load the function of the detected script file
    # -------------------------------------------------------------------------------------------------------------------
    if ($FilePath -notlike "*\Init.ps1" -AND $FilePath -notlike "*\Start.ps1") {
        if ($InputMerge -eq "Merge") {
            StartFile_Merge-AddContentToMergedFile -Name $Name -Content "$FilePath" -Type "File" -MergedFile $AutoScriptFile
        }
        else {
            Invoke-Expression -Command ". '$FilePath'"
        }
    }
}
# ----------------------------------------------------
# Add execution of required startup functions in order
# ----------------------------------------------------
$StartupFunctions = @()
$StartupFunctions += ($Functions | Where-Object { $_.Type -eq "Startup" -AND $_.Name -like "Script*Module*" }).Name
$StartupFunctions += ($Functions | Where-Object { $_.Type -eq "Startup" -AND $_.Name -like "Script-Menu-Categories-1-Start" }).Name
foreach ($StartupFunction in $StartupFunctions) {
    if ($InputMerge -eq "Merge") {
        StartFile_Merge-AddContentToMergedFile -Name $StartupFunction -Content $StartupFunction -Type "Startup" -MergedFile $AutoScriptFile
    }
    else {
        Invoke-Expression -Command $StartupFunction
    }
}
# ------------------------
# Starting the Merged file
# ------------------------
if ($InputMerge -eq "Merge") {
    Write-Host " - Merge successful completed, press Enter to execute the merged file or close this window to exit" -ForegroundColor Cyan
    Pause
    Invoke-Expression -Command ". $AutoScriptFile"
}
# -------------------------
# Cleaning up when crashing
# -------------------------
if ($Error) {
    if ($global:Office365Exchange) {
        Remove-PSSession $global:Office365Exchange
    }
    Get-Variable | Where-Object { $_.Name -notlike "InputArgs*" -AND $_.Name -ne "profile" } | Remove-Variable -ErrorAction SilentlyContinue
    Write-Host "The script has been terminated due to the crash above, please investigate and rerun this script" -ForegroundColor Magenta
    Pause
}