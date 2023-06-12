# Teams-SetupHoliday.ps1
# Setzt im Teams den Urblaub und richtet diesen in der automatischen Telefonzentrale ein

# Lino Bertschinger, 12.06.2023


# Microsoft Teams Modul & Anmeldung
$Modules = Get-Module | Select -ExpandProperty Name
If ("MicrosoftTeams" -notin $Modules) { 
   Connect-MicrosoftTeams }

# Telefonzentralen Name hier anpassen
$AutoAttendantName = "AutoAttendantNAMEHIERNAPSSSEN"

# Feiertage hier anpassen
$Holidays = @(
   @{
         Name = "Neujahr"
         Start = "2023-01-01T00:00:00"
         End = "2023-01-01T23:45:00"
         Announcement = "TEXT ANSAGE HIER"
   },
   @{
         Name = "Berchtoldstag"
         Start = "2023-01-02T00:00:00"
         End = "2023-01-02T23:45:00"
         Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Karfreitag"
      Start = "2023-04-07T00:00:00"
      End = "2023-04-07T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Ostermontag"
      Start = "2023-04-10T00:00:00"
      End = "2023-04-10T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Auffahrt"
      Start = "2023-05-18T00:00:00"
      End = "2023-05-18T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Pfingstmontag"
      Start = "2023-05-29T00:00:00"
      End = "2023-05-29T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Bundesfeier"
      Start = "2023-08-01T00:00:00"
      End = "2023-08-01T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   },
   @{
      Name = "Weihnachten"
      Start = "2023-12-25T00:00:00"
      End = "2023-12-26T23:45:00"
      Announcement = "TEXT ANSAGE HIER"
   }
)

# Konfiguration
foreach ($Holiday in $Holidays) {
   # Setzt Urlaub
   $DateRange = New-CsOnlineDateTimeRange -Start $Holiday.Start -End $Holiday.End
   New-CsOnlineSchedule -Name $Holiday.Name -FixedSchedule -DateTimeRanges @($DateRange)

   # Setzt den Urlaub in der Telefonzentrale
   $HolidaySchedule = Get-CsOnlineSchedule | Where-Object {$_.Name -eq $Holiday.Name}
   $GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $Holiday.Announcement
   $MenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
   $Menu = New-CsAutoAttendantMenu -Name $Holiday.Name -MenuOptions @($MenuOption)
   $CallFlow = New-CsAutoAttendantCallFlow -Name $Holiday.Name -Menu $Menu -Greetings $GreetingPrompt
   $CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $HolidaySchedule.Id -CallFlowId $CallFlow.Id
   $AutoAttendant = Get-CsAutoAttendant -NameFilter $AutoAttendantName
   $AutoAttendant.CallFlows += @($CallFlow)
   $AutoAttendant.CallHandlingAssociations += @($CallHandlingAssociation)
   Set-CsAutoAttendant -Instance $AutoAttendant
}