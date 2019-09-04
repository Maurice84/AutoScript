Function Script-Module-Initial-2-GetPreRequirement {
    # ============
    # Declarations
    # ============
    $global:CheckUAC = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    $global:PowerShell = $PSVersionTable.PSVersion.Major
    $global:OS = (Get-WmiObject -class Win32_OperatingSystem | Select-Object Caption).Caption
    if ($($MyInvocation.ScriptName)) {
        $global:MainTitel = "Merged " + (Get-ChildItem $($MyInvocation.ScriptName) | Select-Object LastWriteTime).LastWriteTime
    }
    else {
        $global:MainTitel = "NonMerged"
    }
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name $global:MainTitel
    if ($global:PowerShell -eq 1) {
        Write-Host "Warning: PowerShell 1.0 has been detected and cannot proceed, please upgrade to the newest version and restart." -ForegroundColor Red
        Write-Host
        $Pause.Invoke()
        EXIT
    }
    if ($global:PowerShell -le 4) {
        Write-Host -NoNewLine "Please note: PowerShell $global:PowerShell has been detected. " -ForegroundColor Cyan
        if ($global:PowerShell -le 3) {
            Write-Host "Only minimal functions can be used like exporting account/mailbox overviews" -ForegroundColor Cyan
        }
        else {
            Write-Host
        }
        Write-Host "You can proceed, but it's highly advisable to upgrade to the latest version to have a fully functioning AutoScript." -ForegroundColor Cyan
        Write-Host "The most recent version is Windows PowerShell 5.1. Compatible with Windows Server 2008 R2 SP1, .NET 4.5 is pre-required:" -ForegroundColor Cyan
        Write-Host
        Write-Host "Microsoft .NET Framework 4.7.2 (reboot needed after install)" -ForegroundColor Yellow
        Write-Host " - https://www.microsoft.com/en-us/download/details.aspx?id=55167" -ForegroundColor Cyan
        Write-Host
        Write-Host "Windows Management Framework 5.1:" -ForegroundColor Yellow
        Write-Host " - https://www.microsoft.com/en-us/download/details.aspx?id=54616" -ForegroundColor Cyan
        Write-Host
        $Pause.Invoke()
    }
    $global:FakeNames = @()
    $global:FakeNames = "Tony Stark", "Matt Murdog", "Danny Rand", "Jessica Jones", `
        "Bruce Banner", "Peter Parker", "Sue Storm", "Hulk", "Iron Man", "Captain America", `
        "Black Panther", "Spider-Man", "Silver Surfer", "Doctor Strange", "Deadpool", "Ant-Man", `
        "Batman", "Aquaman", "Wonder Woman", "Superman", "Bruce Wayne", "Steve Rogers", "Captain Marvel", `
        "Wolverine", "Cyclops", "Phantom", "Green Hornet", "Zorro", "Green Lantern", "The Flash", "Magneto", `
        "Thor", "Gemini", "Professor X", "Iceman", "Beast", "Storm", "Rogue", "Gambit", "Jubilee", "Cable"
}