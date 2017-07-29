<#
.Synopsis
   A script to list all used numbers in a Lync deployment
.DESCRIPTION
    A script to list all used numbers in a Lync deployment
    Created by Lasse Nordvik Wedø - All rights reserved
    Http://tech.rundtomrundt.com

    V 1.1 - April 2012 - Added a function to count the number of users as well
    V 1.2 - June 2012 - Added a function to count and the number of users not enabled for EVas well
    V 1.3 - March 2013 - BugFIX + changing folder and filepath to a static value. Run this script without interaction. Signed
    V 1.4 (unofficial) - October 2013 - Added loop to pull additional information about the user from Active Directory, redid page CSS, fixed typo
					   - http://jackstromberg.com/2013/10/export-a-list-of-numbers-used-in-lync-server-2013/
.EXAMPLE
   Assigned_numbers.ps1
#>

<#
 Setting folder, file and finding date
#>


$filepath = "C:\inetpub\lync-numbers\"
$date = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$file = $filepath + "index.html"
New-Item $filepath -type directory -force -Verbose

<#
 Creating a style for the htm output, coulors and sizes may be adjusted as fit.
 Building HTML Construct for $file
 the "out-file" simply writes whatever is in between the " " into the file.
#>

"<!DOCTYPE html PUBLIC &quot;-//W3C//DTD XHTML 1.0 Transitional//EN&quot; &quot;http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd&quot;>" | Out-File $file
"<html xmlns=&quot;http://www.w3.org/1999/xhtml&quot;>" | Out-File $file -append
"<head>" | Out-File $file -append
"<title>Assigned Users</title>" | Out-File $file -append
"<meta http-equiv=&quot;Content-Type&quot; content=&quot;text/html; charset=ISO-8859-1&quot;>" | Out-File $file -append
"<meta name=&quot;author&quot; content=&quot;Lasse Nordvik Wedø &quot;>" | Out-File $file -append
"<meta name=&quot;copyright&quot; content=&quot;Lasse Nordvik Wedø &quot;>" | Out-File $file -append
"<!--This is a documentation done based on a script by Lasse Nordvik Wedø at Datametrix, all rights reserved -->" | Out-File $file -append
"<style type=`"text/css`">body{background-color:#FFF;font-family:`"wf_SegoeUI`",`"Segoe UI`",`"Segoe`",`"Segoe WP`",`"Tahoma`",`"Verdana`",`"Arial`",`"sans-serif`"}table{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}th{color:#FFF;border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#1570a6}td{width:200px;border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:white}" | Out-File $file -append
"</style>" | Out-File $file -append
"</head>" | Out-File $file -append
"<body>" | Out-File $file -append

<#
 Counting the number of users enabled for Lync
#>

[system.Console]::ForegroundColor = [System.ConsoleColor]::Gray
write-host "Fetching users enabled of Lync"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($sipaddress in (get-csuser -filter {LineURI -eq $Null})) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

<#first creating a header#>
"<H3>There are $Showcounter users enabled for Lync without a lineURI</H3>" | Out-File $file -append
<#writing result#>
Get-CsUser -Filter {LineURI -eq $Null} | sort -Property sipaddress | Select-Object Name,sipaddress | ConvertTo-HTML | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of users with a lineURI and listing them
#>

[system.Console]::ForegroundColor = [System.ConsoleColor]::Gray
write-host "Fetching users with LineURI"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($lineuri in (get-csuser -filter {LineURI -ne $Null})) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter users with a primary LineURI</H3>" | Out-File $file -append
$UserList = Get-CsUser -Filter {LineURI -ne $Null} | sort -Property LineURI | Select-Object SamAccountName,LineURI,Name,SipAddress
$userInfo = @()
foreach ($LyncUser in $UserList)
{
$ADUser = Get-ADUser -Identity $LyncUser.SAMAccountName -Properties Department, Title, Mobile
$tableObj = new-object System.Object
$tableObj | Add-Member -type NoteProperty -Name "LineURI" -Value $LyncUser.LineURI
$tableObj | Add-Member -type NoteProperty -Name "Name" -Value $LyncUser.Name
$tableObj | Add-Member -type NoteProperty -Name "Title" -Value $ADUser.Title
$tableObj | Add-Member -type NoteProperty -Name "Department" -Value $ADUser.Department
$tableObj | Add-Member -type NoteProperty -Name "Mobile Number" -Value $ADUser.Mobile
$tableObj | Add-Member -type NoteProperty -Name "SIPAddress" -Value $LyncUser.SIPAddress
$userInfo += $tableObj
}

$userInfo | ConvertTo-HTML | Out-File $file -append

# Cleaning up the results
$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of users with a private line and listing them
#>

write-host "Fetching users with a private line"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($privateline in (get-csuser -Filter {privateline -ne $Null})) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter users with a private line</H3>" | Out-File $file -append
Get-CsUser -Filter {privateline -ne $Null} | sort -Property privateline | Select-Object privateline,Name,sipaddress | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of analog devices and listing them
#>

write-host "Fetching analog devices"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsAnalogDevice) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter analog devices</H3>" | Out-File $file -append
Get-CsAnalogDevice -Filter {LineURI -ne $Null}  | Sort -property lineuri | Select-Object lineuri,displayname  | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of common area phones and listing them
#>

write-host "Fetching common area phones"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsCommonAreaPhone) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter common area phones devices</H3>" | Out-File $file -append
Get-CsCommonAreaPhone -Filter {LineURI -ne $Null}   | sort -property lineuri | Select-object lineuri, displaynumber, displayname | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of RGS workflows and listing them
#>

write-host "Fetching RGS workflows"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsRgsWorkflow) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter RGS workflows with a LineURI</H3>" | Out-File $file -append
Get-CsRgsWorkflow | sort -Property lineuri | Select-object lineuri,displaynumber,name,active | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of voice conference numbers and listing them
#>

write-host "Fetching voice conference numbers"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsDialInConferencingAccessNumber) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter dialin access numbers</H3>" | Out-File $file -append
Get-CsDialInConferencingAccessNumber -Filter {LineURI -ne $Null} | sort -Property lineuri | Select-object lineuri,displayname | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of Exchange objects and listing them
#>

write-host "Fetching voice Exchange objects"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsExUmContact) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter Exchange objects</H3>" | Out-File $file -append
Get-CsExUmContact -Filter {LineURI -ne $Null} | sort -Property lineuri | Select-object Lineuri,displayname,displayuri | ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

<#
Counting the number of trusted application points and listing them
#>

write-host "Fetching voice trusted application points"
$Showcounter = 0
function countmy-users {
$counting = 0
foreach ($LineURI in Get-CsTrustedApplicationEndpoint) {
$counting++
}
write-output $counting
}
$Showcounter += countmy-users

"<H3>There are $Showcounter trusted application endpoints with a LineURI</H3>" | Out-File $file -append
Get-CsTrustedApplicationEndpoint -Filter {LineURI -ne $Null} | sort -Property lineuri | Select-object -Property lineuri, displayname, displaynumber| ConvertTo-HTML -fragment | Out-File $file -append

$Showcounter = $NULL
$counting = $NULL
$Showcounter = $NULL

[system.Console]::ForegroundColor = [System.ConsoleColor]::Green
write-host "Creating output file $file"
write-host "Done!"
[system.Console]::ForegroundColor = [System.ConsoleColor]::Gray
# SIG # Begin signature block
# MIIP9QYJKoZIhvcNAQcCoIIP5jCCD+ICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbpWOrLzNonLx4lRB0mtZRraM
# 0rWggg06MIIGjzCCBXegAwIBAgIQCdVQr1MHEqo87VeUdjqdqTANBgkqhkiG9w0B
# AQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTEzMDMwNDAwMDAwMFoXDTE0MDMxMjEy
# MDAwMFowWzELMAkGA1UEBhMCTk8xEDAOBgNVBAcTB05lc3R0dW4xHDAaBgNVBAoM
# E0xhc3NlIE5vcmR2aWsgV2Vkw7gxHDAaBgNVBAMME0xhc3NlIE5vcmR2aWsgV2Vk
# w7gwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCmEeGD7Q+uIGlBY0Dk
# RaQVA6Ruytap2iFG0d2wOq/TZoXBWJ/OVw2RJGorbbDplHK0A8av+OcWNS9yt9pd
# uUwszpEF37SNMr7BJM9zwv5vAQ+JrmqjDwQq6aXxGJEpN+iGO0EV2h6mqMuCQArg
# ku5Q3vltoaTZi5DMwMGHRTwXhSWen7uri5UM/PtgThla6WVPN8oH2MDtGpMBZQut
# 73EwWz4MwKKoMuIwyZiaH11gPfPNR815MGYWsBqY5QS62wcBhWd4qPbp53PWqbwE
# BYGpi9UWo7VYpWfiPFGCLrnQ02fJLMUAfMu/HlkcZkqbyNVKkYLo3PxB/Z78Fc6F
# /CLpAgMBAAGjggM5MIIDNTAfBgNVHSMEGDAWgBR7aM4pqsAXvkl64eU/1qf3RY81
# MjAdBgNVHQ4EFgQUqPyBB5j5AsCEFifaVyJGKIcX6AgwDgYDVR0PAQH/BAQDAgeA
# MBMGA1UdJQQMMAoGCCsGAQUFBwMDMHMGA1UdHwRsMGowM6AxoC+GLWh0dHA6Ly9j
# cmwzLmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFhLmNybDAzoDGgL4YtaHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL2Fzc3VyZWQtY3MtMjAxMWEuY3JsMIIBxAYD
# VR0gBIIBuzCCAbcwggGzBglghkgBhv1sAwEwggGkMDoGCCsGAQUFBwIBFi5odHRw
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYI
# KwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAg
# AEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAg
# AGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBl
# AHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBu
# AGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAg
# AGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAg
# AGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIABy
# AGUAZgBlAHIAZQBuAGMAZQAuMIGCBggrBgEFBQcBAQR2MHQwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBMBggrBgEFBQcwAoZAaHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ29kZVNpZ25pbmdD
# QS0xLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBBQUAA4IBAQBQvhNA+IYp
# rQSK2o73IfP/I9RJsV3gude+53pKR8iC4myKuIWEbWpT87xkl3Ca52cTM1htwkv9
# TldywiFyVmIcbq4wT9BsfghI793F1OG7spZ0qCZ30PaEmFXXFVIeDSdL/JPjT3k+
# nBWPNkmHfeAGHaCcL7n1lhtfBFbyfeera60IEUv/zmUwyIkL0oCnLYfFinYnwnOI
# bjQNSZDpURU5OLgyqFmkAD20fULcewvlLl2vnmXaVZZYl/YAB4gfptVXt+vM3Nft
# m3e8T8osoC4PHwoAkrls1n57dDASdbY80bw2h4J71apQOvVwmplnMx52wuKko8Ra
# 6sF8Bt8ubMPiMIIGozCCBYugAwIBAgIQD6hJBhXXAKC+IXb9xextvTANBgkqhkiG
# 9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1
# cmVkIElEIFJvb3QgQ0EwHhcNMTEwMjExMTIwMDAwWhcNMjYwMjEwMTIwMDAwWjBv
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENv
# ZGUgU2lnbmluZyBDQS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# nHz5oI8KyolLU5o87BkifwzL90hE0D8ibppP+s7fxtMkkf+oUpPncvjxRoaUxasX
# 9Hh/y3q+kCYcfFMv5YPnu2oFKMygFxFLGCDzt73y3Mu4hkBFH0/5OZjTO+tvaaRc
# AS6xZummuNwG3q6NYv5EJ4KpA8P+5iYLk0lx5ThtTv6AXGd3tdVvZmSUa7uISWjY
# 0fR+IcHmxR7J4Ja4CZX5S56uzDG9alpCp8QFR31gK9mhXb37VpPvG/xy+d8+Mv3d
# KiwyRtpeY7zQuMtMEDX8UF+sQ0R8/oREULSMKj10DPR6i3JL4Fa1E7Zj6T9OSSPn
# BhbwJasB+ChB5sfUZDtdqwIDAQABo4IDQzCCAz8wDgYDVR0PAQH/BAQDAgGGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMIIBwwYDVR0gBIIBujCCAbYwggGyBghghkgBhv1s
# AzCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1j
# cHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAA
# dQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAA
# YwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8A
# ZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4A
# ZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUA
# ZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwA
# aQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQA
# IABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wEgYDVR0T
# AQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0f
# BHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUe2jOKarAF75J
# euHlP9an90WPNTIwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJ
# KoZIhvcNAQEFBQADggEBAHtyHWT/iMg6wbfp56nEh7vblJLXkFkz+iuH3qhbgCU/
# E4+bgxt8Q8TmjN85PsMV7LDaOyEleyTBcl24R5GBE0b6nD9qUTjetCXL8KvfxSgB
# VHkQRiTROA8moWGQTbq9KOY/8cSqm/baNVNPyfI902zcI+2qoE1nCfM6gD08+zZM
# kOd2pN3yOr9WNS+iTGXo4NTa0cfIkWotI083OxmUGNTVnBA81bEcGf+PyGubnviu
# nJmWeNHNnFEVW0ImclqNCkojkkDoht4iwpM61Jtopt8pfwa5PA69n8SGnIJHQnEy
# hgmZcgl5S51xafVB/385d2TxhI2+ix6yfWijpZCxDP8xggIlMIICIQIBATCBgzBv
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENv
# ZGUgU2lnbmluZyBDQS0xAhAJ1VCvUwcSqjztV5R2Op2pMAkGBSsOAwIaBQCgeDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBTi2STm2bYirJBZLlYSnyewYZz3ajANBgkqhkiG9w0BAQEFAASCAQAaGL7eIfag
# GjLDYUWE768mqXNwOd0EuZACP5D5s6+MHuQBX9HV1IRjIKBus8MX8R+rilIywMEb
# HPuzwyJT90puGZNm59ZJjYzD8z0izHOzR4lg0uEVv9oY7Ko0+kh8jFblxt/UJfiA
# mEFne+fyM0xSO8WjP0K1oe4axQFtxfuF6is+sI7GOyZsBsWrhRRHFMniOgCCAnHq
# jzXoee8n0URCFHcaHUcPbt/E8/K0J8FpgjNoX+8eNeeYLMpXCgGqvv4WEQDB3fK8
# 5VTZS61OLapTtN6ULxUJQ2YCEtjsDl1hztN3/F/Y2n6qIFFNPqNg969gzKVePB0f
# oU/sk7/jyMFo
# SIG # End signature block