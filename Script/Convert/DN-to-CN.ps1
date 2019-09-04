function Script-Convert-DN-to-CN {
    param (
        [string]$Debug,
        [string]$Path
    )
    # -----------------------------------------------------------------------------
    # Converting DistinguishedName (OU=..,OU=...,DC=...) to CanonicalName (\OU\OU\)
    # -----------------------------------------------------------------------------
    [string]$Subtree = $null
    If ($Path -like "*CN=Users,DC*") {
        $Path = $Path.Replace('CN=Users,DC', 'OU=Users,DC')
    }
    if ($Debug) {
        Write-Host "Path is" $Path -ForegroundColor Yellow  
    }
    ForEach ($Item in ($Path.Replace('\,', '~').Split(","))) {
        Switch -Regex ($Item.TrimStart().Substring(0, 3)) {
            "OU=" {
                [array]$TempOU += $Item.Replace("OU=", "")
                [array]$TempOU += '\'
            }
            "DC=" {
                [array]$TempDC += $Item.Replace("DC=", "")
                [array]$TempDC += '.'
            }
        }
    }
    if ($Debug) {
        Write-Host "TempOU is" $TempOU -ForegroundColor Yellow
        Write-Host "TempOU.Count is" $TempOU.Count -ForegroundColor Yellow
        Write-Host "TempOU.Length is" $TempOU.Length -ForegroundColor Yellow
    }
    If ($TempOU.Count -ge 1) {
        For ($i = $TempOU.Count; $i -ge 0; $i--) {
            $TempOUs += $TempOU[$i]
            if ($Debug) {
                Write-Host "TempOUs +=" $TempOU[$i] -ForegroundColor Yellow
            }
        }
        $Subtree += [string]$TempOUs.Substring(1)
    }
    if (!$Subtree) {
        $Subtree = "\"
    }
    if ($Debug) {
        Write-Host "Subtree is" $Subtree -ForegroundColor Yellow
        $Pause.Invoke()
    }
    return $Subtree
}