Function Script-Module-Initial-5-SetSystemLocale ([string]$Locale = "nl-NL") {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Threading")
    [System.Reflection.Assembly]::LoadWithPartialName("System.Globalization")
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::CreateSpecificCulture($Locale)
}