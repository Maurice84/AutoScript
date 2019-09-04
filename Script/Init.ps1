$InputArgs0 = $args[0]
$InputArgs1 = $args[1]
Get-Variable | Where-Object { $_.Name -notlike "InputArgs*" -AND $_.Name -ne "profile" } | Remove-Variable -ErrorAction SilentlyContinue
$global:SystemFunctions = (Get-ChildItem function:\ | Where-Object { $_.Noun -ne $null -AND $_.Verb -ne $null })