#############
# This script will auto enroll to a specific lesson of sportbit. Example for CrossFit Amersfoort
#
# This script needs to Run in PowerShell Core 7 and not in the Windows Build in Powershell
# Download powershell here: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.2#msi
#
# By scheduling this at 00:01 on the next day, your automatically subscribed to the lesson. So if I want to schedule for next weeks mondays WOD. I need to run this
# script at 00:01
#
# This script is beta and the following bug detected:
# - Subscribing to sunday lessons is not possible by default. Since Sunday is not on Next week when run on monday 00:01
#############

$sportbit_username = "SportBitUsername" #your sportbit username here
$sportbit_password = "SuperSecretPassword" #your sportbit password here
$sportbit_boxurl = "https://crossfitamersfoort.sportbitapp.nl" #you can grab this url from the emails of Sprotbit waitinglist urls
$sportbit_locatie = "1" # When loging in with the browser, click a calendar. The URL will show a GET value with ?locatie=1 or ?locatie=7. Specify the location number on which schedule the lesson is.
$logfile = "C:\Temp\Aanmelden_SportBit_tent.txt" # Output of logfile. If you don't want a logfile, remove all the Add-Content lines in script for now
$lessontime = "19:30" #Start time of the workout
$lessonname = "WOD" #Name of the workout. I.e. "OnRamp" of "Gewichtheffen (OLY)"
$teams_webhook = "URL to teams webhook" #URL to you Teams webhook, so you can receive a message. If you don't use teams, remove last line "Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $teamsmessage_json -Uri "$($teams_webhook)"" for now

#Generated VARS
$lesmomenten_url = "$($sportbit_boxurl)/cbm/account/lesmomenten/?locatie=$($sportbit_locatie)"
$sportbit_inlogurl = "$($sportbit_boxurl)/cbm/account/inloggen/"
$sportbit_logouturl = "$($sportbit_boxurl)/cbm/account/logout/"

#daynummers
#0 = Monday
#5 = Sathurday
[Int]$trainingdaynumber = (Get-Date).DayOfWeek - 1 #Today minus 1 to subscribe to lesson

Add-Content -path $logfile -value "================================="
Add-Content -path $logfile -value "New Run $(Get-Date)"
Add-Content -path $logfile -value "================================="
#Weeknummer

[int]$thisweek = Get-Date -UFormat %V
[int]$nextweek = $thisweek + 1
[int]$thisyear = Get-Date -UFormat %Y

Add-Content -path $logfile -value "This week: $($thisweek)"
Add-Content -path $logfile -value "Next week: $($nextweek)"
Add-Content -path $logfile -value "This year: $($thisyear)"

#TeamsCard
$teamscard = "This week: $($thisweek) <br>"
$teamscard += "Next week: $($nextweek) <br>"
$teamscard += "This year: $($thisyear)<br>"

$SportBit_nextweekscheduleurl = "$($sportbit_boxurl)/cbm/account/lesmomenten/$($thisyear)/$($nextweek)/?locatie=$($sportbit_locatie)"

Add-Content -path $logfile -value "SportBit Nextweek schedule URL: $($SportBit_nextweekscheduleurl)"

#TeamsCard
$teamscard += "SportBit Nextweek schedule URL: <a href=""$($SportBit_nextweekscheduleurl)"">$($SportBit_nextweekscheduleurl)</a><br>"

#######
# Login to sportbit
#######

#On the sportbit page these are the login attributes in HTML
#<form action="account/inloggen/?post=1" method="POST">
#<input type="text" placeholder="Jouw gebruikersnaam / e-mailadres" name="username" value="" />
#<input type="password" class="left" placeholder="Jouw wachtwoord" name="password" value="" />
#<input type="submit" class="" value="Inloggen"/>

#Create headers to login
$SportBit_LoginHeader = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$SportBit_LoginHeader.Add("User-Agent", 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:68.0) Gecko/20100101 Firefox/68.0')

$SportBit_PostURL = "$($sportbit_boxurl)/cbm/account/inloggen/?post=1"

Add-Content -path $logfile -value "Start loging in to Sportbit: $($SportBit_PostURL)"

#TeamsCard
$teamscard += "Start loging in to Sportbit: <a href=""$($SportBit_PostURL)"">$($SportBit_PostURL)</a><br>"

Write-Host "Sportbit: Fetching Cookie"
Add-Content -path $logfile -value "Sportbit: Fetching Cookie"

#TeamsCard
$teamscard += "Sportbit: Fetching Cookie<br>"

$SportBit_LoginPage = Invoke-WebRequest -Uri $SportBit_PostURL -SessionVariable pbo_session -Method Get -Headers $SportBit_LoginHeader

#GenerateLoginform
$formdata = @{
    username = $sportbit_username
    password = $sportbit_password
}

#Send the login request to SportBit
Write-Host "Sportbit: Logging in with user $($sportbit_username)"
Add-Content -path $logfile -value "Sportbit: Logging in with user $($sportbit_username)"

#TeamsCard
$teamscard += "Sportbit: Logging in with user $($sportbit_username)<br>"

$SportBit_Login = Invoke-WebRequest -Uri "$($sportbit_boxurl)/cbm/account/inloggen/?post=1" -WebSession $pbo_session -Headers $SportBit_LoginHeader -Method Post -Form $formdata


########
# Subscribe to lesson
########
$lesmomenten_dezeweek = Invoke-WebRequest -Uri $SportBit_nextweekscheduleurl -WebSession $pbo_session -Method Get -Headers $SportBit_LoginHeader

$nextweekdate = (Get-Date).AddDays(6).ToString('dd-MM-yyyy')
Add-Content -path $logfile -value "Next week lesson date : $($nextweekdate)"

#TeamsCard
$teamscard += "Next week lesson date : $($nextweekdate)<br>"
$lessonURL = $lesmomenten_dezeweek | Select-Object -ExpandProperty links | Where-Object href -Match "/$($nextweekdate)/$($lessontime)" | Select-Object -ExpandProperty href

#example URL https://crossfitamersfoort.sportbitapp.nl/cbm/training-info/12-02-2022/13:00/30878/aanmelden
#output training-info/16-02-2022/19:30/31017/

$register_url = "$($sportbit_boxurl)/cbm/$($lessonURL)aanmelden"

#Regsiter for Training
if ($null -ne $register_url) {
    Add-Content -path $logfile -value "Register for Training: $($lessonname)"

    #TeamsCard
    $teamscard += "Register for Training: $($lessonname)<br>"

    Write-Host "Sportbit: Register Traing URL: $($register_url)"
    Add-Content -path $logfile -value "Sportbit: Register Traing URL: $($register_url)"

    #TeamsCard
    $teamscard += "Sportbit: Register Traing URL: <a href=""$($register_url)"">$($register_url)</a><br>"

    $register_training = Invoke-WebRequest -Uri $register_url -WebSession $pbo_session -Headers $SportBit_LoginHeader -Method Get
    $register_training.Content | Out-File $logfile
    Write-Host "Sportbit: Registered for training $($lessonname)"
    Add-Content -path $logfile -value "Sportbit: Registered for training $($lessonname)"

    #TeamsCard
    $teamscard += "Sportbit: <b>Registered for training $($lessonname)</b><br>"
} else {
    Write-Host "Error fetching training URL $($register_url)"
    Add-Content -path $logfile -value "Error fetching training URL $($register_url)"

    #TeamsCard
    $teamscard += "Error fetching training URL $($register_url)<br>"
}
########
# Logout from Sportbit
########
Write-Host "Sportbit: Logout user $($sportbit_username)"
Add-Content -path $logfile -value "Sportbit: Logout user $($sportbit_username)"
#TeamsCard
$teamscard += "Sportbit: Logout user $($sportbit_username)<br>"
$SportBit_logout = Invoke-WebRequest -Uri $sportbit_logouturl -WebSession $pbo_session -Headers $SportBit_LoginHeader -Method Get

<##
$teamsmessage = @{
    "type" = "message"
    "attachments" = @( 
        @{
            "contentType" = "application/vnd.microsoft.card.adaptive"
            "contentUrl" = $null
            "content" = @{
               '$schema' =  "http://adaptivecards.io/schemas/adaptive-card.json"
               "type" = "AdaptiveCard"
               "version" = "1.2"
               "body" = @(
                        @{
                        "type"= "TextBlock"
                        "text" = $teamscard
                   }
                )
            }
        }
    )
}
##>

$teamsmessage = @{
    "text" = $teamscard
}

$teamsmessage_json = ConvertTo-Json -InputObject $teamsmessage -Depth 5

Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $teamsmessage_json -Uri "$($teams_webhook)"