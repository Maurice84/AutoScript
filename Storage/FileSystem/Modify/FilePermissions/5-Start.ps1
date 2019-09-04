Function Storage-FileSystem-Modify-FilePermissions-5-Start {
    param (
        [array]$Applications,
        [array]$Servers
    )
    # ============
    # Declarations
    # ============
    [array]$Apps = $Applications | Select-Object Path -Unique | Sort-Object Path
    $ServerCount = 1
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host
    Write-Host "Setting permissions on server(s)" -ForegroundColor Magenta
    # ---------------------------
    # Iterate through each server
    # ---------------------------
    ForEach ($Server in $Servers) {
        $AppCount = 1
        ForEach ($Path in $Apps.Path) {
            Write-Host (" - Processing permissions on " + $Server + ":") -ForegroundColor Yellow
            Write-Host -NoNewLine ("   > Processing file " + $AppCount + " of " + $Apps.Count + ": " + $Path)
            $Path = "\\" + $Server + "\" + $Path[0] + "$\" + $Path.Substring(3)
            If ((Get-Item $Path) -isnot [System.IO.DirectoryInfo]) {
                $AppDesc = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path).FileDescription
                If ($AppDesc) {
                    Write-Host -NoNewLine (" (" + $AppDesc + ")") -ForegroundColor Yellow
                }
            }
            Write-Host -NoNewline "... "
            # ---------------------------------------
            # Remove Inheritance and copy permissions
            # ---------------------------------------
            $ACL = (Get-Item $Path).GetAccessControl('Access')
            If (($ACL.Access).IsInherited -eq $true) {
                $ACL.SetAccessRuleProtection($True, $True)
                Set-ACL -Path "$Path" -ACLObject $ACL
            }
            # ---------------------------------
            # Remove BUILTIN\Users if it exists
            # ---------------------------------
            $ACLGroups = $ACL.Access | Where-Object { $_.IdentityReference -like "*\Users" }
            If ($ACLGroups) {
                Do {
                    $ACLGroups | ForEach-Object { $ACL.RemoveAccessRule($_) | Out-Null }
                    Set-ACL -Path "$Path" -ACLObject $ACL
                    $ACL = (Get-Item $Path).GetAccessControl('Access')
                    $ACLGroups = $ACL.Access | Where-Object { $_.IdentityReference -like "*\Users" }
                } While ($ACLGroups)
            }
            # ---------------------------------------
            # Set the group permissions on the object
            # ---------------------------------------
            ForEach ($App in $Applications) {
                If ($Path -eq ("\\" + $Server + "\" + ($App.Path)[0] + "$\" + ($App.Path).Substring(3))) {
                    $Group = $App.Group
                    $ACL = (Get-Item "$Path").GetAccessControl('Access')
                    $AppACL = $ACL.Access | Where-Object { $_.IdentityReference -like "*$Group" }
                    If (!($AppACL)) {
                        If ((Get-Item $Path) -is [System.IO.DirectoryInfo]) {
                            $InheritanceFlags = "ContainerInherit, ObjectInherit"
                        }
                        Else {
                            $InheritanceFlags = "None"
                        }
                        $ReadAndExecute = New-Object System.Security.AccessControl.FileSystemAccessRule("$Group", "ReadAndExecute", $InheritanceFlags, "None", "Allow")
                        $ACL.SetAccessRule($ReadAndExecute)
                        Set-ACL -Path "$Path" -ACLObject $ACL
                        Write-Host " OK!" -ForegroundColor Green
                    }
                    Else {
                        Write-Host " Already set" -ForegroundColor Gray
                    }
                }
            }
            If ($AppCount -lt $Apps.Count) {
                $AppCount++
            }
        }
        If ($ServerCount -lt $Servers.Count) {
            $ServerCount++
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}