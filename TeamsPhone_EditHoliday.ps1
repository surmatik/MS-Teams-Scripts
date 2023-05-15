# Teams-EditHoliday.ps1
# -----
# Skript ersetzt die Feiertagen mit den neuen

$Modules = Get-Module | Select -ExpandProperty Name
If ("MicrosoftTeams" -notin $Modules) { 
   Connect-MicrosoftTeams }

# Neujahr
$HolidayName = "Neujahr"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-01-01T00:00:00" -End "2023-01-02T23:45:00"
$Neujahr = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Neujahr.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Neujahr

# Karfreitag
$HolidayName = "Karfreitag"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-04-07T00:00:00" -End "2023-04-07T23:45:00"
$Karfreitag = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Karfreitag.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Karfreitag

# Ostermontag
$HolidayName = "Ostermontag"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-04-10T00:00:00" -End "2023-04-10T23:45:00"
$Ostermontag = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Ostermontag.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Ostermontag

# Auffahrt
$HolidayName = "Auffahrt"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-05-18T00:00:00" -End "2023-05-18T23:45:00"
$Auffahrt = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Auffahrt.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Auffahrt

# Pfingstmontag
$HolidayName = "Pfingstmontag"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-05-29T00:00:00" -End "2023-05-29T23:45:00"
$Pfingstmontag = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Pfingstmontag.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Pfingstmontag

# Bundesfeier
$HolidayName = "Bundesfeier"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-08-01T00:00:00" -End "2023-08-01T23:45:00"
$Bundesfeier = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Bundesfeier.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Bundesfeier

# Weihnachten
$HolidayName = "Weihnachten"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-12-25T00:00:00" -End "2023-12-25T23:45:00"
$Weihnachten = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Weihnachten.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Weihnachten

# Stephanstag
$HolidayName = "Stephanstag"
$DateRange = New-CsOnlineDateTimeRange -Start "2023-12-26T00:00:00" -End "2023-12-26T23:45:00"
$Stephanstag = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
$Stephanstag.FixedSchedule.DateTimeRanges = $DateRange
Set-CsOnlineSchedule -Instance $Stephanstag