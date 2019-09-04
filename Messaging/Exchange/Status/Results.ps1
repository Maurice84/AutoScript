Function Messaging-Exchange-Status-Results([string]$Selectie, [string]$Actief, [string]$SortObject = "Name") {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If ($global:ExchangeVersion -eq "2007" -AND $Selectie -eq "MailboxExport" -AND $Actief -ne "Ja") {
        Write-Host -NoNewLine "Exchange 2007 detected! " -ForegroundColor Magenta; Write-Host -NoNewLine "This version uses XML-files to retrieve the mailbox request status."; Write-Host
        Write-Host "Therefore you have to select a (network)drive where the PST and XML-file(s) are located:"
        Write-Host
        Functie-Selectie-Drive -Filter "XML" -Selectie "Import"
    }
    If ($Selectie -eq "Export") {
        $Type = "export request(s)"
        Write-Host -NoNewLine "Indexing the mailbox $Type statistics..."
        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange) {
            $MBStatus = @()
            $GetMailboxExportRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxExportRequest -DomainController '$PDC'"))
            $GetMailboxExportRequest = ($GetMailboxExportRequest | Select-Object Name, RequestGuid | Sort-Object $SortObject).RequestGuid
            ForEach ($MB in $GetMailboxExportRequest) {
                $MBStats = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxExportRequestStatistics '$MB' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
                If ($MBStats) {
                    $MBStatus += $MBStats
                }
                Else {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-MailboxExportRequestStatistics '$MB' -Confirm:0 -DomainController '$PDC'"))
                }
            }
        }
        If ($global:ExchangeVersion -eq "2007") {
            $MBStatus = Get-ChildItem $Drive -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*XML*" }
        }
    }
    If ($Selectie -eq "Import") {
        $Type = "import request(s)"
        Write-Host -NoNewLine "Indexing the mailbox $Type statistics..."
        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange) {
            $MBStatus = @()
            $GetMailboxImportRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxImportRequest -DomainController '$PDC'"))
            $GetMailboxImportRequest = ($GetMailboxImportRequest | Select-Object Name, RequestGuid | Sort-Object $SortObject).RequestGuid
            ForEach ($MB in $GetMailboxImportRequest) {
                $MBStats = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxImportRequestStatistics '$MB' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
                If ($MBStats) {
                    $MBStatus += $MBStats
                }
                Else {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-MailboxImportRequestStatistics '$MB' -Confirm:0 -DomainController '$PDC'"))
                }
            }
        }
        If ($global:ExchangeVersion -eq "2007") {
            Write-Host "Unfortunately mailbox $ExportImportType does not work with Exchange Server 2007." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    If ($Selectie -eq "Move") {
        $Type = "move request(s)"
        $MBStatus = @()
        Do {
            Write-Host -NoNewline "Would you like to have the "; Write-Host -NoNewLine "copied data size and percentage" -ForegroundColor Yellow; Write-Host -NoNewLine " indexed? This can take longer (Y/N): "
            $Choice = Read-Host
            $Input = @("Y", "N") -contains $Choice
            If (!$Input) {
                Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($Input)
        Switch ($Choice) {
            "Y" {
                $DetectSizeAndPercentage = $true
            }
        }
        Script-Module-SetHeaders -Name $Titel
        Write-Host -NoNewLine "Indexing the mailbox move-request(s), this can take a few moments..."
        If ($DetectSizeAndPercentage) {
            If (!$Office365Exchange) {
                $GetMoveRequests = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MoveRequest -ResultSize Unlimited | Select-Object DisplayName, Identity"))
            }
            Else {
                $GetMoveRequests = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MoveRequest -ResultSize Unlimited | Select-Object DisplayName, Identity"))
            }
            $Counter = 1
            ForEach ($MoveRequest in $GetMoveRequests) {
                Script-Module-SetHeaders -Name $Titel
                Write-Host -NoNewLine ("Indexing mailbox move-request(s) statistics, this can take a while... (" + $Counter + "/" + [string]$GetMoveRequests.Count + ")")
                $MoveRequestIdentity = $MoveRequest.Identity
                If (!$Office365Exchange) {
                    $GetMoveRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MoveRequest '$MoveRequestIdentity' | Get-MoveRequestStatistics -ErrorAction SilentlyContinue | Select-Object DisplayName, Status, PercentComplete, LastUpdateTimestamp, Message, BytesTransferred")
                    )
                }
                Else {
                    $GetMoveRequest = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create(`
                                "Get-MoveRequest '$MoveRequestIdentity' | Get-MoveRequestStatistics -ErrorAction SilentlyContinue | Select-Object DisplayName, Status, PercentComplete, LastUpdateTimestamp, Message, BytesTransferred")
                    )
                }
                $Properties = @{`
                        DisplayName         = $GetMoveRequest.DisplayName; `
                        Status              = $GetMoveRequest.Status; `
                        PercentComplete     = $GetMoveRequest.PercentComplete; `
                        LastUpdateTimestamp = $GetMoveRequest.LastUpdateTimestamp; `
                        Message             = $GetMoveRequest.Message; `
                        BytesTransferred    = $GetMoveRequest.BytesTransferred
                }
                $Object = New-Object PSObject -Property $Properties
                $MBStatus += $Object
                $Counter++
            }
        }
        Else {
            If (!$Office365Exchange) {
                $MBStats = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MoveRequest -ResultSize Unlimited | Select-Object DisplayName, Status"))
            }
            Else {
                $MBStats = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MoveRequest -ResultSize Unlimited | Select-Object DisplayName, Status"))
            }
            If ($MBStats) {
                $MBStatus += $MBStats
            }
        }
        $MBStatus = $MBStatus | Sort-Object DisplayName
    }
    If ($Selectie -eq "Restore") {
        $Type = "restore request(s)"
        Write-Host -NoNewLine "Indexing the mailbox $Type..."
        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange) {
            $MBStatus = @()
            $GetMailboxRestoreRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxRestoreRequest -DomainController '$PDC'"))
            $GetMailboxRestoreRequest = ($GetMailboxRestoreRequest | Select-Object Name, RequestGuid | Sort-Object $SortObject).RequestGuid
            ForEach ($MB in $GetMailboxRestoreRequest) {
                $MBStats = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxRestoreRequestStatistics '$MB' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
                If ($MBStats) {
                    $MBStatus += $MBStats
                }
                Else {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-MailboxRestoreRequestStatistics '$MB' -Confirm:0 -DomainController '$PDC'"))
                }
            }
        }
        If ($global:ExchangeVersion -eq "2007") {
            Write-Host "Unfortunately mailbox $ExportImportType does not work with Exchange Server 2007." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    $global:Array = @()
    $ColumnNameLength = 0
    $ColumnStatusLength = 0
    $ColumnDateTimeLength = 0
    $ColumnSizeLength = 0
    $Counter = 1
    ForEach ($MailboxStatus in $MBStatus) {
        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -OR $global:Office365Exchange) {
            $Name = $MailboxStatus.DisplayName
            If (!$Name) {
                $Name = $MailboxStatus.Name
            }
            $Status = ($MailboxStatus.Status).Value
            If (!$Status) {
                $Status = $MailboxStatus.Status
            }
            If ($MailboxStatus.PercentComplete) {
                $Status = [string]$MailboxStatus.PercentComplete + "% " + $Status
            }
            $DateTime = $MailboxStatus.LastUpdateTimestamp
            If (!$DateTime) {
                $DateTime = "n.a."
            }
            $Message = $MailboxStatus.Message
            If (!$Message -OR $Message.Length -le 1) {
                $Message = "n.a."
            }
            If ($MailboxStatus.BytesTransferred) {
                $Size = [string]$MailboxStatus.BytesTransferred
                $Size = [double]($Size).Split(' ')[2].Replace('(', '').Replace(',', '')
            }
            Else {
                $Size = $null
            }
        }
        If ($global:ExchangeVersion -eq "2007") {
            $FileXML = $MailboxStatus.FullName
            $XML = [xml](Get-Content $FileXML)
            $Name = ($XML."export-Mailbox".TaskDetails | Select-Object Item).Item.Source.DisplayName
            $DateTime = ($XML."export-Mailbox".TaskFooter | Select-Object EndTime).EndTime
            $Message = ($XML."export-Mailbox".TaskDetails | Select-Object Item).Item.Result."#text"
            $Size = (Get-ChildItem ($XML."export-Mailbox".TaskDetails | Select-Object Item).Item.Target.PSTFilePath | Select-Object Length).Length
            $StatusCompleted = ($XML."export-Mailbox".TaskFooter | Select-Object Result).Result.CompletedCount
            $StatusFailed = ($XML."export-Mailbox".TaskFooter | Select-Object Result).Result.ErrorCount
            If ($StatusCompleted -eq 1) { $Status = "Completed" }
            If ($StatusFailed -eq 1) { $Status = "Failed" }
        }
        If ($Name.Length -gt $ColumnNameLength) { $ColumnNameLength = $Name.Length }
        If ($Status.Length -gt $ColumnStatusLength) { $ColumnStatusLength = $Status.Length }
        If (([string]$DateTime).Length -gt $ColumnDateTimeLength) { $ColumnDateTimeLength = ([string]$DateTime).Length }
        If (([string]$Size).Length) {
            If (([string]$Size).Length -ge 10) {
                $SizeType = "GB"
            }
            Else {
                $SizeType = "MB"
            }
            $Size = "{0:F2}" -f ($Size / ("1" + $SizeType)) + (" " + $SizeType)
        }
        Else {
            $Size = "n.a."
        }
        If ($Size.Length -gt $ColumnSizeLength) {
            $ColumnSizeLength = $Size.Length
        }
        If (!$FileXML.Length) {
            $FileXML = "n.a."
        }
        # ----------------
        # Scrambling time!
        # ----------------
        #If ($Name -notlike "Test*") {
        #    $FakeName = ($FakeNames | Get-Random) + " " + (Get-Random -Maximum 9999)
        #    $Properties = @{RealName=$Name;Name=$FakeName;}
        #} Else {
        $Properties = @{RealName = $Name; Name = $Name; }
        #}
        # ----------------
        $Properties += @{Nr = $Counter; Status = $Status; DateTime = $DateTime; Message = $Message; Size = $Size; Log = $FileXML }
        $Object = New-Object PSObject -Property $Properties
        $global:Array += $Object
        $Counter++
    }
    $Counter = 0
    [System.Collections.ArrayList]$global:ArrayList = $global:Array
    $global:ArrayList = $global:ArrayList | Sort-Object Name
    If ($global:ArrayList.Count) {
        $CountArrayList = $global:ArrayList.Count
    }
    Else {
        $CountArrayList = 1
    }
    While ($global:ArrayList) {
        $Input = $null
        $Keuze = $null
        $Sorting = "A-Z"
        Do {
            Script-Module-SetHeaders -Name $Titel
            Write-Host -NoNewLine "The following "; Write-Host -NoNewLine $global:ArrayList.Count -ForegroundColor Yellow; Write-Host " mailbox $Type are initiated with the following status:"
            Write-Host
            Write-Host (" " * 5) "[N]ame:" (" " * ($ColumnNameLength - 6)) "[S]tatus:" (" " * ($ColumnStatusLength - 8)) "[D]atum:" (" " * ($ColumnDateTimeLength - 6)) "Grootte:"
            $KolomCounter = 0
            Do {
                $Nr = ($Counter + 1)
                $Name = $global:ArrayList[$Counter].Name
                $Status = [string]$global:ArrayList[$Counter].Status
                $DateTime = [string]$global:ArrayList[$Counter].DateTime
                $Size = [string]$global:ArrayList[$Counter].Size
                $Message = [string]$global:ArrayList[$Counter].Message
                $ColumnNameSpaties = (" " * ($ColumnNameLength - $Name.Length + 3))
                $ColumnStatusSpaties = (" " * ($ColumnStatusLength - $Status.Length + 3))
                $ColumnDateTimeSpaties = (" " * ($ColumnDateTimeLength - $DateTime.Length + 3))
                $ColumnSizeSpaties = (" " * ($ColumnSizeLength - $Size.Length))
                If ($Size.Length -le 6) { $Size = (" " * 3) + $Size }
                If ($Size.Length -le 7) { $Size = (" " * 2) + $Size }
                If ($Size.Length -le 8) { $Size = (" " * 1) + $Size }
                If ($Message.Length -eq 0) { $Message = "n.a." }
                If ($Nr -ge 100) { $Nr = " " + $Nr + ". " }
                ElseIf ($Nr -ge 10) { $Nr = "  " + $Nr + ". " }
                Else { $Nr = "   " + $Nr + ". " }
                $KolomTekst = [string]$Nr + $Name + $ColumnNameSpaties + $Status + $ColumnStatusSpaties + $DateTime + $ColumnDateTimeSpaties + $Size
                If ($Status -like "*Completed") { Write-Host $KolomTekst -ForegroundColor Green }
                If ($Status -like "*InProgress") { Write-Host $KolomTekst -ForegroundColor Yellow }
                If ($Status -like "*Suspended" -AND $Status -notlike "*Auto*") { Write-Host $KolomTekst -ForegroundColor Magenta }
                If ($Status -like "*AutoSuspended") { Write-Host $KolomTekst -ForegroundColor Cyan }
                If ($Status -like "*Failed") { Write-Host $KolomTekst -ForegroundColor Red }
                $Counter++
                $KolomCounter++
            } Until (($KolomCounter -eq $global:MaxRowCount) -OR ($Counter -eq $global:ArrayList.Count))
            Write-Host
            $Completed = $global:ArrayList | Select-Object Status | Where-Object { $_.Status -like "*Completed" }
            If ($Completed) {
                If ($Completed.Count) {
                    $CountCompleted = $Completed.Count
                }
                Else {
                    $CountCompleted = 1
                }
            }
            Else {
                $CountCompleted = 0
            }
            $InProgress = $global:ArrayList | Select-Object Status | Where-Object { $_.Status -like "*InProgress" }
            If ($InProgress) {
                If ($InProgress.Count) {
                    $CountInProgress = $InProgress.Count
                }
                Else {
                    $CountInProgress = 1
                }
            }
            Else {
                $CountInProgress = 0
            }
            $Suspended = $global:ArrayList | Select-Object Status | Where-Object { $_.Status -like "*Suspended" -AND $_.Status -notlike "*Auto*" }
            If ($Suspended) {
                If ($Suspended.Count) {
                    $CountSuspended = $Suspended.Count
                }
                Else {
                    $CountSuspended = 1
                }
            }
            Else {
                $CountSuspended = 0
            }
            $AutoSuspended = $global:ArrayList | Select-Object Status | Where-Object { $_.Status -like "*AutoSuspended" }
            If ($AutoSuspended) {
                If ($AutoSuspended.Count) {
                    $CountAutoSuspended = $AutoSuspended.Count
                }
                Else {
                    $CountAutoSuspended = 1
                }
            }
            Else {
                $CountAutoSuspended = 0
            }
            $Failed = $global:ArrayList | Select-Object Status | Where-Object { $_.Status -like "*Failed" }
            If ($Failed) {
                If ($Failed.Count) {
                    $CountFailed = $Failed.Count
                }
                Else {
                    $CountFailed = 1
                }
            }
            Else {
                $CountFailed = 0
            }
            Write-Host -NoNewLine "  Status:"; Write-Host -NoNewLine " InProgress: " -ForegroundColor Gray; Write-Host -NoNewLine $CountInProgress -ForegroundColor Yellow
            Write-Host -NoNewLine " - Completed: " -ForegroundColor Gray; Write-Host -NoNewLine $CountCompleted -ForegroundColor Yellow
            Write-Host -NoNewLine " - Suspended: " -ForegroundColor Gray; Write-Host -NoNewLine $CountSuspended -ForegroundColor Yellow
            Write-Host -NoNewLine " - AutoSuspended: " -ForegroundColor Gray; Write-Host -NoNewLine $CountAutoSuspended -ForegroundColor Yellow
            Write-Host -NoNewLine " - Failed: " -ForegroundColor Gray; Write-Host $CountFailed -ForegroundColor Yellow
            Write-Host
            # ----------------
            # Keuze pagina('s)
            # ----------------
            If (!($Keuze) -AND ($KolomCounter -eq $global:MaxRowCount) -AND ($Counter -lt $global:ArrayList.Count) -AND ($Counter -le $global:MaxRowCount)) {
                Write-Host ">> [V]olgende Pagina"
                Write-Host
                $Segment = "Eerste"
            }
            If (!($Keuze) -AND ($KolomCounter -eq $global:MaxRowCount) -AND ($Counter -lt $global:ArrayList.Count) -AND ($Counter -gt $global:MaxRowCount)) {
                Write-Host "<< [T]erug naar vorige pagina  |  [V]olgende Pagina >>"
                Write-Host
                $Segment = "Midden"
            }
            If (!($Keuze) -AND ($KolomCounter -le $global:MaxRowCount) -AND ($Counter -eq $global:ArrayList.Count) -AND ($Counter -gt $global:MaxRowCount)) {
                Write-Host "<< [T]erug naar vorige pagina"
                Write-Host
                $Segment = "Eind"
            }
            If (!($Keuze) -AND ($Counter -le $global:MaxRowCount) -AND ($Counter -eq $global:ArrayList.Count)) {
                $Segment = "Alles"
            }
            Do {
                If ($Selectie -eq "Move") {
                    If (!$global:MoveChangeMenu -AND !$global:MoveDeleteMenu -AND !$global:NewMoveStatus) {
                        Write-Host -NoNewLine "Would you like to ("; Write-Host -NoNewLine "C" -ForegroundColor Yellow; Write-Host -NoNewLine ")hange or ("; Write-Host -NoNewLine "R" -ForegroundColor Yellow; Write-Host ")emove 1 or more mailbox move-requests?"
                        $global:MoveMenu = $true
                    }
                    If ($global:MoveChangeMenu) {
                        Write-Host -NoNewLine "Please select a category: ("; Write-Host -NoNewLine "R" -ForegroundColor Yellow; Write-Host -NoNewLine ")esume, ("; Write-Host -NoNewLine "P" -ForegroundColor Yellow; Write-Host -NoNewLine ")ause or ("; Write-Host -NoNewLine "M" -ForegroundColor Yellow; Write-Host ")igrate (relevant with status AutoSuspended)"
                    }
                }
                If ($Selectie -ne "Move" -OR $global:NewMoveStatus -OR $global:MoveDeleteMenu) {
                    Write-Host -NoNewLine "Please select a mailbox request to  "
                    If ($NewMoveStatus) {
                        Write-Host -NoNewLine $NewMoveStatus -ForegroundColor Cyan
                    }
                    Else {
                        Write-Host -NoNewLine "delete" -ForegroundColor Cyan
                    }
                    Write-Host -NoNewLine " or use "; Write-Host -NoNewLine "A" -ForegroundColor Yellow; Write-Host " to select all."
                }
                Write-Host -NoNewLine "Use "; Write-Host -NoNewLine "R" -ForegroundColor Yellow; Write-Host -NoNewLine " to refresh the status or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous "
                If ($global:MoveChangeMenu -OR $global:MoveDeleteMenu) {
                    Write-Host -NoNewLine "question: "
                }
                Else {
                    Write-Host -NoNewLine "menu: "
                }
                [string]$Choice = Read-Host
                If ($Choice -as [int]) {
                    If ([int]$Choice -eq "0" -OR [int]$Choice -gt $CountArrayList) {
                        If ($Selectie -ne "Move" -OR ($Selectie -eq "Move" -AND ($global:NewMoveStatus -OR $global:MoveDeleteMenu))) {
                            Write-Host "Please use above as input" -ForegroundColor Red; Write-Host
                        }
                    }
                    Else {
                        $Input = $Choice
                        $Keuze = $global:ArrayList[$Choice - 1].RealName
                        If ($Segment -eq "Eerste" -OR $Segment -eq "Alles") {
                            $Counter = 0
                            BREAK
                        }
                        If ($Segment -eq "Midden") {
                            $Counter -= $global:MaxRowCount
                            BREAK
                        }
                        If ($Segment -eq "Eind") {
                            If ($KolomCounter -ge 2) {
                                $Counter -= $KolomCounter
                                BREAK
                            }
                            Else {
                                $Counter -= ($KolomCounter + $global:MaxRowCount)
                                BREAK
                            }
                        }
                    }
                }
                Else {
                    If ($Segment -eq "Midden" -OR $Segment -eq "Eind") {
                        If ($Choice -eq "T") {
                            $Counter -= ($KolomCounter + $global:MaxRowCount)
                            BREAK
                        }
                    }
                    If ($Segment -eq "Eerste" -OR $Segment -eq "Midden") {
                        If ($Choice -eq "V") {
                            BREAK
                        }
                    }
                    If ($global:MoveChangeMenu) {
                        If ($Choice -eq "H") {
                            $global:NewMoveStatus = "resume"
                        }
                        If ($Choice -eq "P") {
                            $global:NewMoveStatus = "pause"
                        }
                        If ($Choice -eq "M") {
                            $global:NewMoveStatus = "migrate"
                        }
                        $Input = $Choice
                        $global:MoveChangeMenu = $null
                    }
                    If ($Choice -eq "W") {
                        $Input = $Choice
                        $global:MoveMenu = $false
                        $global:MoveDeleteMenu = $true
                    }
                    If ($Choice -eq "X") {
                        If ($global:MoveChangeMenu -OR $global:MoveDeleteMenu) {
                            If ($global:MoveChangeMenu) {
                                $global:MoveMenu = $true
                                $global:MoveChangeMenu = $null
                            }
                            If ($global:MoveDeleteMenu) {
                                $global:MoveMenu = $true
                                $global:MoveDeleteMenu = $null
                            }
                        }
                        Else {
                            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; $global:Exchange = $null; Invoke-Expression -Command $global:MenuNameCategory
                        }
                    }
                    ElseIf ($Choice -eq "R") {
                        Messaging-Exchange-Status-Results -Selectie $Selectie -Actief $Actief -SortObject $SortObject
                    }
                    ElseIf ($Choice -eq "A") {
                        If ($Selectie -ne "Move" -OR ($Selectie -eq "Move" -AND ($global:MoveMenu -OR $global:MoveChangeMenu -OR $global:MoveDeleteMenu))) {
                            $Input = $Choice
                            If ($global:MoveMenu) {
                                $global:MoveChangeMenu = $true
                                $global:MoveMenu = $null
                            }
                            Else {
                                $Keuze = "Alles"
                            }
                        }
                    }
                    ElseIf ($Choice -eq "N") {
                        If ($CountArrayList -ge 2) {
                            $SortObject = "Name"
                            $SortObjectFormat = "Naam"
                            $Input = $Choice
                        }
                    }
                    ElseIf ($Choice -eq "S") {
                        If ($CountArrayList -ge 2) {
                            $SortObject = "Status"
                            $SortObjectFormat = "Status"
                            $Input = $Choice
                        }
                    }
                    ElseIf ($Choice -eq "D") {
                        If ($CountArrayList -ge 2) {
                            $SortObject = "DateTime"
                            $SortObjectFormat = "Datum"
                            $Input = $Choice
                        }
                    }
                    Else {
                        Write-Host "Please use above as input" -ForegroundColor Red; Write-Host
                    }
                    If ($SortObjectFormat -OR $global:MoveMenu -OR $global:MoveChangeMenu -OR $global:MoveDeleteMenu -OR $global:NewMoveStatus) {
                        If ($SortObjectFormat) {
                            If ($Sorting -eq "Z-A") {
                                $global:ArrayList = $global:ArrayList | Sort-Object $SortObject
                                $Sorting = "A-Z"
                            }
                            ElseIf ($Sorting -eq "A-Z") {
                                $global:ArrayList = $global:ArrayList | Sort-Object $SortObject -Descending
                                $Sorting = "Z-A"
                            }
                            Else {
                                $Sorting = $null
                            }
                        }
                        If ($Counter -ge $global:MaxRowCount) {
                            $Counter = $Counter - $global:MaxRowCount
                        }
                        Else {
                            $Counter = 0
                        }
                        BREAK
                    }
                }
            } Until ($Input)
        } Until ($Keuze)
        If ($Selectie -eq "Export") {
            If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange) {
                If ($Keuze -eq "Alles") {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxExportRequest -DomainController '$PDC' | Remove-MailboxExportRequest -Confirm:0 -DomainController '$PDC'")
                    )
                }
                Else {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxExportRequest -Name '$Keuze' -DomainController '$PDC' | Remove-MailboxExportRequest -Confirm:0 -DomainController '$PDC'")
                    )
                }
            }
            If ($global:ExchangeVersion -eq "2007") {
                If ($Keuze -eq "Alles") {
                    ForEach ($Item in $global:ArrayList) {
                        Write-Host $Item.Name $Item.Log
                        Remove-Item ($Item.Log) -ErrorAction SilentlyContinue
                    }
                }
                Else {
                    Remove-Item ($global:ArrayList[$Choice - 1].Log) -ErrorAction SilentlyContinue
                }
            }
        }
        If ($Selectie -eq "Import") {
            If ($Keuze -eq "Alles") {
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxImportRequest -DomainController '$PDC' | Remove-MailboxImportRequest -Confirm:0 -DomainController '$PDC'")
                )
            }
            Else {
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxImportRequest -Name '$Keuze' -DomainController '$PDC' | Remove-MailboxImportRequest -Confirm:0 -DomainController '$PDC'")
                )
            }
        }
        If ($Selectie -eq "Move" -AND $Keuze) {
            If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -OR $global:Office365Exchange) {
                If (!$global:PSSessionExchange) {
                    If ($Office365Exchange) {
                        Import-PSSession $Office365Exchange -WarningAction SilentlyContinue | Out-Null
                    }
                    Else {
                        Import-PSSession -Session (Get-PSSession -Name 'Exchange') -DisableNameChecking -WarningAction SilentlyContinue | Out-Null
                    }
                    $global:PSSessionExchange = $True
                }
                If ($Keuze -eq "Alles") {
                    If ($global:MoveDeleteMenu) {
                        If (!$Office365Exchange) {
                            Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue -DomainController "$PDC" | Remove-MoveRequest -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue | Remove-MoveRequest -Confirm:$false
                        }
                    }
                    If ($global:NewMoveStatus -eq "resume") {
                        If (!$Office365Exchange) {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus Suspended -ErrorAction SilentlyContinue -DomainController "$PDC" | Resume-MoveRequest -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus Suspended -ErrorAction SilentlyContinue | Resume-MoveRequest -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "InProgress"
                        }
                    }
                    If ($global:NewMoveStatus -eq "migrate") {
                        If (!$Office365Exchange) {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus AutoSuspended -ErrorAction SilentlyContinue -DomainController "$PDC" | Resume-MoveRequest -SuspendWhenReadyToComplete:$false -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus AutoSuspended -ErrorAction SilentlyContinue | Resume-MoveRequest -SuspendWhenReadyToComplete:$false -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "InProgress"
                        }
                    }
                    If ($global:NewMoveStatus -eq "pause") {
                        If (!$Office365Exchange) {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus InProgress -ErrorAction SilentlyContinue -DomainController "$PDC" | Suspend-MoveRequest -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Get-MoveRequest -ResultSize Unlimited -MoveStatus InProgress -ErrorAction SilentlyContinue | Suspend-MoveRequest -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "Suspended"
                        }
                    }
                }
                Else {
                    If ($global:MoveDeleteMenu) {
                        If (!$Office365Exchange) {
                            Remove-MoveRequest -Identity "$Keuze" -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Remove-MoveRequest -Identity "$Keuze" -Confirm:$false
                        }
                    }
                    If ($global:NewMoveStatus -eq "resume") {
                        If (!$Office365Exchange) {
                            Resume-MoveRequest -Identity "$Keuze" -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Resume-MoveRequest -Identity "$Keuze" -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "InProgress"
                        }
                    }
                    If ($global:NewMoveStatus -eq "migrate") {
                        If (!$Office365Exchange) {
                            Resume-MoveRequest -Identity "$Keuze" -SuspendWhenReadyToComplete:$false -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Resume-MoveRequest -Identity "$Keuze" -SuspendWhenReadyToComplete:$false -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "InProgress"
                        }
                    }
                    If ($global:NewMoveStatus -eq "pause") {
                        If (!$Office365Exchange) {
                            Suspend-MoveRequest -Identity "$Keuze" -Confirm:$false -DomainController "$PDC"
                        }
                        Else {
                            Suspend-MoveRequest -Identity "$Keuze" -Confirm:$false
                        }
                        If ($?) {
                            $ChangedMoveStatus = "Suspended"
                        }
                    }
                }
            }
        }
        If ($Selectie -eq "Restore") {
            If ($Keuze -eq "Alles") {
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxRestoreRequest -DomainController '$PDC' | Remove-MailboxRestoreRequest -Confirm:0 -DomainController '$PDC'")
                )
            }
            Else {
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                            "Get-MailboxRestoreRequest -Name '$Keuze' -DomainController '$PDC' | Remove-MailboxRestoreRequest -Confirm:0 -DomainController '$PDC'")
                )
            }
        }
        If ($?) {
            If ($Keuze -eq "Alles") {
                If ($Selectie -ne "Move" -OR ($Selectie -eq "Move" -AND !$global:NewMoveStatus)) {
                    $global:ArrayList = $null
                }
                If ($Selectie -eq "Move" -AND $global:NewMoveStatus) {
                    $TempArrayList = @()
                    ForEach ($MoveRequest In $global:ArrayList) {
                        If ($global:NewMoveStatus -eq "resume") {
                            If ($MoveRequest.Status -eq "Suspended") {
                                $MoveRequestStatus = $ChangedMoveStatus
                            }
                            Else {
                                $MoveRequestStatus = $MoveRequest.Status
                            }
                        }
                        If ($global:NewMoveStatus -eq "pause") {
                            If ($MoveRequest.Status -eq "InProgress") {
                                $MoveRequestStatus = $ChangedMoveStatus
                            }
                            Else {
                                $MoveRequestStatus = $MoveRequest.Status
                            }
                        }
                        If ($global:NewMoveStatus -eq "migrate") {
                            If ($MoveRequest.Status -eq "AutoSuspended") {
                                $MoveRequestStatus = $ChangedMoveStatus
                            }
                            Else {
                                $MoveRequestStatus = $MoveRequest.Status
                            }
                        }
                        $Properties = @{`
                                Nr       = $MoveRequest.Nr; `
                                RealName = $MoveRequest.RealName; `
                                Name     = $MoveRequest.Name; `
                                Status   = $MoveRequestStatus; `
                                DateTime = $MoveRequest.DateTime; `
                                Message  = $MoveRequest.Message; `
                                Size     = $MoveRequest.Size; `
                                Log      = $MoveRequest.FileXML`
                        
                        }
                        $Object = New-Object PSObject -Property $Properties
                        $TempArrayList += $Object
                    }
                    $global:ArrayList = $TempArrayList
                }
            }
            Else {
                If ($Selectie -ne "Move" -OR ($Selectie -eq "Move" -AND $Keuze -AND !$global:NewMoveStatus)) {
                    $global:ArrayList.RemoveAt(($Choice - 1))
                }
                If ($Selectie -eq "Move" -AND $Keuze -AND $global:NewMoveStatus) {
                    $TempArrayList = @()
                    ForEach ($MoveRequest In $global:ArrayList) {
                        If ($MoveRequest.RealName -eq $Keuze) {
                            $MoveRequestStatus = $ChangedMoveStatus
                        }
                        Else {
                            $MoveRequestStatus = $MoveRequest.Status
                        }
                        $Properties = @{`
                                Nr       = $MoveRequest.Nr; `
                                RealName = $MoveRequest.RealName; `
                                Name     = $MoveRequest.Name; `
                                Status   = $MoveRequestStatus; `
                                DateTime = $MoveRequest.DateTime; `
                                Message  = $MoveRequest.Message; `
                                Size     = $MoveRequest.Size; `
                                Log      = $MoveRequest.FileXML`
                        
                        }
                        $Object = New-Object PSObject -Property $Properties
                        $TempArrayList += $Object
                    }
                    $global:ArrayList = $TempArrayList
                }
            }
        }
        Else {
            Write-Host "ERROR: An error occurred, please investigate!" -ForegroundColor Red
            Write-Host
            $Pause.Invoke()
            EXIT
        }
    }
    # ==========
    # Finalizing
    # ==========
    If (!$global:ArrayList) {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        Write-Host "There are no mailbox $Type found, now returning to previous menu" -ForegroundColor Yellow
        Write-Host
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; $global:Exchange = $null; Invoke-Expression -Command $global:MenuNameCategory
    }
}