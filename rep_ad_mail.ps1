$a = "<style>"
$a = $a + "BODY{background-color:White;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#eef2f3}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:White}"
$a = $a + "</style>"

Get-ADUser -Filter {(mail -notlike "Health*") -and (mail -notlike "*test*") -and (mail -notlike "MSME*") -and (mail -notlike "oitt*") -and (mail -ne "null") -and (Enabled -eq "true")} -Properties displayName,Mail,Department | Select-Object DisplayName,Mail,Department | ConvertTo-HTML -head $a | Out-File .\"Mail.html"

Invoke-Expression ".\Mail.html"
