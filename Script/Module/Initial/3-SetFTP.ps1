Function Script-Module-Initial-3-SetFTP {
<#
    If (Test-Connection 192.168.0.250 -Count 1 -ErrorAction SilentlyContinue) {
        $global:URL = "ftp://192.168.0.250/Web/"
    } Else {
        $global:URL = "ftp://maurice84.myqnapcloud.com/Web/"
    }
    $User = "update-script"
    $Pass = "Soxu39_7"
    $global:Credentials = New-Object System.Net.NetworkCredential($User, $Pass)
    $global:WebClient = New-Object System.Net.WebClient 
    $global:WebClient.Credentials = New-Object System.Net.NetworkCredential($User, $Pass)
#>
}
