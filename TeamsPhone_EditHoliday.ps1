# Teams-EditHoliday.ps1
# Skript ersetzt die Feiertagen mit den neuen

# Lino Bertschinger, 12.06.2023


$Modules = Get-Module | Select -ExpandProperty Name
If ("MicrosoftTeams" -notin $Modules) { 
   Connect-MicrosoftTeams }

# Feiertage hier anpassen
$Holidays = @(
   @{
         Name = "Neujahr"
         Start = "2023-01-01T00:00:00"
         End = "2023-01-02T23:45:00"
   },
   @{
      Name = "Karfreitag"
      Start = "2023-04-07T00:00:00"
      End = "2023-04-07T23:45:00"
   },
   @{
      Name = "Ostermontag"
      Start = "2023-04-10T00:00:00"
      End = "2023-04-10T23:45:00"
   },
   @{
      Name = "Auffahrt"
      Start = "2023-05-18T00:00:00"
      End = "2023-05-18T23:45:00"
   },
   @{
      Name = "Pfingstmontag"
      Start = "2023-05-29T00:00:00"
      End = "2023-05-29T23:45:00"
   },
   @{
      Name = "Bundesfeier"
      Start = "2023-08-01T00:00:00"
      End = "2023-08-01T23:45:00"
   },
   @{
      Name = "Weihnachten"
      Start = "2023-12-25T00:00:00"
      End = "2023-12-25T23:45:00"
   },
   @{
      Name = "Stephanstag"
      Start = "2023-12-26T00:00:00"
      End = "2023-12-26T23:45:00"
   }
)

# Konfiguration
foreach ($Holiday in $Holidays) {
   $HolidayName = $Holiday.Name
   $DateRange = New-CsOnlineDateTimeRange -Start $Holiday.Start -End $Holiday.End
   $HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $HolidayName}
   $HolidaySchedule.FixedSchedule.DateTimeRanges = $DateRange
   Set-CsOnlineSchedule -Instance $HolidaySchedule
}
