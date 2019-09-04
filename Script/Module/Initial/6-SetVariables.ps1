Function Script-Module-Initial-6-SetVariables {
    # ============
    # Declarations
    # ============
    # ------------------
    # Customer Variables
    # ------------------
    $global:Customer = $null
    $global:CustomerFS = $null
    $global:CustomerWithSpaces = $null
    $global:CustomerUPNSuffix = $null
    # --------------
    # Menu Variables
    # --------------
    $global:CurrentUser = (([ADSI]"WinNT://$env:userdomain/$env:username,user").FullName -Split (" "))[0]
    If (!$global:CurrentUser -AND !$global:CurrentUser.Length -le 1) {
        $global:CurrentUser = ([ADSI]"WinNT://$env:userdomain/$env:username,user").Name
        If (!$global:CurrentUser) {
            $global:CurrentUser = "Administrator"
        }
    }
    If ((Get-Date).Hour -ge 0) { $global:Greetings = "Good night" }
    If ((Get-Date).Hour -ge 5) { $global:Greetings = "Good morning" }
    If ((Get-Date).Hour -ge 12) { $global:Greetings = "Good afternoon" }
    If ((Get-Date).Hour -ge 17) { $global:Greetings = "Good evening" }
    $global:Greetings += " " + $global:CurrentUser + "!"
    # ----------------
    # System Variables
    # ----------------
    $global:Actief = $null
    $global:ActiveDirectory = $null
    $global:CredManager = $null
    $global:Exchange = $null
    $global:ExchangeServer = $null
    $global:ExchangeVersion = $null
    $global:Functions = $null
    $global:MenuNameCategory = $null
    $global:MenuNameStart = $null
    $global:MenuNameTask = $null
    $global:PDC = $null
    $global:Office365 = $null
    $global:Office365Connect = $null
    $global:RemoteActiveDirectory = $null
    $global:RemoteExchange = $null
    $global:SessionAD = $null
    $global:StartupVariables = $null
    New-Variable -Name StartupVariables -Value ((Get-Variable | Select-Object Name) | % { $_.Name }) -Force -Scope Global
}
