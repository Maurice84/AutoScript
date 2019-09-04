Function Messaging-Exchange-Modify-Mailboxes-4-SetType {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $global:WeergaveNaam -ForegroundColor Yellow
    Write-Host -NoNewLine "- Current type mailbox: "; Write-Host $global:SoortMailbox -ForegroundColor Yellow
    Write-Host
    Write-Host "  1. Equipment" -ForegroundColor Yellow
    Write-Host "  2. Room" -ForegroundColor Yellow
    Write-Host "  3. Shared" -ForegroundColor Yellow
    Write-Host "  4. User" -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Please select a category or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
        $Choice = Read-Host
        $Input = @("1"; "2"; "3"; "4"; "X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "1" {
            $Type = "Equipment"
            $RecipientTypeDetails = "EquipmentMailbox"
        }
        "2" {
            $Type = "Room"
            $RecipientTypeDetails = "RoomMailbox"
        }
        "3" {
            $Type = "Shared"
            $RecipientTypeDetails = "SharedMailbox"
        }
        "4" {
            $Type = "Regular"
            $RecipientTypeDetails = "UserMailbox"
        }
        "X" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
        }
    }
    $SetType = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -Type '$Type' -DomainController '$PDC'"))
    $CheckType = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
    If (($CheckType | Select-Object RecipientTypeDetails).RecipientTypeDetails -ne $RecipientTypeDetails) {
        Write-Host "ERROR: An error occurred setting the mailbox to type $Type, please investigate!" -ForegroundColor Red
    }
    Else {
        If ($Type = "Regular") {
            $global:SoortMailbox = "User"
        }
        Else {
            $global:SoortMailbox = $Type
        }
        Write-Host "OK: Mailbox successfully set to $global:SoortMailbox" -ForegroundColor Green
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
}