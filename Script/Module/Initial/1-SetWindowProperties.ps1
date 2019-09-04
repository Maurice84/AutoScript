Function Script-Module-Initial-1-SetWindowProperties {
    # ============
    # Declarations
    # ============
    $global:Shell = (Get-Variable -Name profile).Value
    If ($Shell -notlike "*ISE*" -AND $Shell -notlike "*VSCode*") {
        $PSHost = Get-Host
        $PSWindow = $PSHost.UI.RawUI
        $PSTitle = $PSWindow.WindowTitle
        $PSTitle = $MainTitel
        $PSWindow.BackgroundColor = "Black"
        $Newsize = $PSWindow.BufferSize
        $Newsize.Height = 250
        $Newsize.Width = 155
        $PSWindow.Buffersize = $Newsize
        $Newsize = $PSWindow.WindowSize
        $Newsize.Height = 50
        $Newsize.Width = 155
        $PSWindow.WindowSize = $Newsize
        $global:MaxRowCount = 30
        $global:Pause = { Write-Host; Write-Host -NoNewLine "Press a key to continue..."; $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") }
    }
    Else {
        $global:MaxRowCount = 10
        $global:Pause = { Write-Host; pause }
    }
    $global:MaxRowLength = 155
}