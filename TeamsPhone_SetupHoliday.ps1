# Teams-SetupHoliday.ps1
# -----
# Setzt im Teams den Urblaub und richtet diesen in der automatischen Telefonzentrale ein

$Modules = Get-Module | Select -ExpandProperty Name
If ("MicrosoftTeams" -notin $Modules) { 
   Connect-MicrosoftTeams }


# Urlaub setzten

#Neujahr 01.01.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-01-01T00:00:00" -End "2023-01-01T23:45:00"
New-CsOnlineSchedule -Name "Neujahr" -FixedSchedule -DateTimeRanges @($DateRange)

#Berchtoldstag 02.01.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-01-02T00:00:00" -End "2023-01-02T23:45:00"
New-CsOnlineSchedule -Name "Berchtoldstag" -FixedSchedule -DateTimeRanges @($DateRange)

#Karfreitag 07.04.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-04-07T00:00:00" -End "2023-04-07T23:45:00"
New-CsOnlineSchedule -Name "Karfreitag" -FixedSchedule -DateTimeRanges @($DateRange)

#Ostersonntag 09.04.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-04-09T00:00:00" -End "2023-04-9T23:45:00"
New-CsOnlineSchedule -Name "Ostersonntag" -FixedSchedule -DateTimeRanges @($DateRange)

#Auffahrt 18.05.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-05-18T00:00:00" -End "2023-05-18T23:45:00"
New-CsOnlineSchedule -Name "Auffahrt" -FixedSchedule -DateTimeRanges @($DateRange)

#Pfingstmontag 29.05.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-05-29T00:00:00" -End "2023-05-29T23:45:00"
New-CsOnlineSchedule -Name "Pfingstmontag" -FixedSchedule -DateTimeRanges @($DateRange)

#Bundesfeier 01.08.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-08-01T00:00:00" -End "2023-08-01T23:45:00"
New-CsOnlineSchedule -Name "Bundesfeier" -FixedSchedule -DateTimeRanges @($DateRange)

#Weihnachten 25.12.2023
$DateRange = New-CsOnlineDateTimeRange -Start "2023-12-25T00:00:00" -End "2023-12-26T23:45:00"
New-CsOnlineSchedule -Name "Weihnachten" -FixedSchedule -DateTimeRanges @($DateRange)



# Ferien / Feiertage in Automatische Telefonzentrale setzen

# Neujahr
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Neujahr"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Berchtoldstag
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Berchtoldstag"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Karfreitag
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Karfreitag"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Ostersonntag
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Ostersonntag"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Auffahrt
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Auffahrt"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Pfingstmontag
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Pfingstmontag"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Bundesfeier
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Bundesfeier"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant

# Weihnachten
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"
$HolidayName = "Weihnachten"
$HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$HolidayScheduleName = $HolidaySchedule.Name
$GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "TEXT ANSAGE HIER"
$MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$Menu = New-CsAutoAttendantMenu -Name $HolidayScheduleName -MenuOptions @($MenuOption)
$CallFlow = New-CsAutoAttendantCallFlow -Name $HolidayScheduleName -Menu $Menu -Greetings $GreetingPrompt
$CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
$AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
$AutoAttendant.CallFlows += @($CallFlow)
$AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
Set-CsAutoAttendant -Instance $AutoAttendant