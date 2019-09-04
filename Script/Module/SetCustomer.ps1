Function Script-Module-SetCustomer {
    param (
        [string]$Name
    )
    If (!$global:Customer) {
        Script-Module-SetHeaders -Name $Name
        Write-Host " You have selected to deploy an environment for a new customer. This script will ask a few one-time questions." -ForegroundColor Magenta
        Write-Host " As long as this script is not aborted, the script will remember and re-use the customer details." -ForegroundColor Magenta
        Write-Host
        Do {
            Write-Host -NoNewLine " Question 1/3: Please enter the customer name including capitals as stated in the project: "
            $InputChoice = Read-Host
            $InputKey = $InputChoice
            If (!$InputKey) {
                Write-Host "Please use above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($InputKey)
        $global:CustomerWithSpaces = $InputKey
        $global:Customer = $InputKey.Replace(' ', '')
        $InputKey = $null
        Do {
            Write-Host -NoNewLine " Question 2/3: Please enter a domain (the suffix after @, like customer.nl): "
            $InputChoice = Read-Host
            $InputKey = $InputChoice
            If (!$InputKey -OR $InputKey -notlike "*.*") {
                Write-Host " Please use above as input" -ForegroundColor Red
                $InputKey = $null
            }
        } Until ($InputKey)
        $global:CustomerUPNSuffix = $InputKey
        Do {
            Write-Host -NoNewLine " Question 3/3: Please enter a reachable file-server name: "
            $InputChoice = Read-Host
            $InputKey = $InputChoice
            If ($InputKey) {
                $CheckFS = Invoke-Command -ComputerName $InputKey -ErrorAction SilentlyContinue -ScriptBlock { $true }
                If ($CheckFS -ne $true) {
                    Write-Host " The server is unreachable, please re-enter the server" -ForegroundColor Red
                    $InputKey = $null
                }
            }
            Else {
                Write-Host " Please use above as input" -ForegroundColor Red
            }
        } Until ($InputKey)
        $global:CustomerFS = $InputKey.ToUpper()
    }
}