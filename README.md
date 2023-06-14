# MS-Teams-Scripts

## Authentifizierung <br>
Beim Starten der Skripte wird `Connect-MicrosoftTeams` ausgeführt, welches ein Fenster öffnet indem man sich mit seinem Microsoft 365  Account authentifizieren kann.
Soll der Skript in einem anderen Tenant, in welches man via Microsoft Partner Programm zugriff hat, so kann man die Tenant ID Des Kunden Tenants in die Variable `$TenantID` einfügen. Wenn es der eigene Tenant ist, kann man die Variable leer lassen.

## Telefonnummer Zuweissung
Script: [Teams-PhoneNumberAssigment.ps1](Teams-PhoneNumberAssigment.ps1) <br>
Der Skript weisst den Benutzern die im Script angegebn werden, die entsprechende Telefonnummer und die entsprechende Richtline zu.

In der Variable `$UsersAndPhoneNumbers` muss mit folgendem Syntax die Mail Adresse des Microsoft 365 Accounts und die Telefnnummer des Benutzers angegeben werden.
```Powershell
$UsersAndPhoneNumbers = @{
    "user1@surmatik.ch" = "+41441234501"
    "user2@surmatik.ch" = "+41441234502"
}
```

## Ferien / Feiertage für Telefonzentrale
###  Ferien / Feiertage neu erstellen <br>
Script: [TeamsPhone_SetupHoliday.ps1](TeamsPhone_SetupHoliday.ps1) <br>
Der Skript hinterlegt im Tenant die angegeben Urlaube und fügt diese mit der Ansage zu einer Telefonzentrale hinzu.

In der Variable `$AutoAttendantName` muss der Name der Telefonzentrale angegeben werden. <br>
Danach kann in der Variable `$Holidays` mit folgendem Syntax die Ferien / Feiertage, mit Datum und der Text Ansage angegeben werden
```Powershell
@{
    Name = "Neujahr"
    Start = "2023-01-01T00:00:00"
    End = "2023-01-01T23:45:00"
    Announcement = "TEXT ANSAGE HIER"
},
```

### Ferien / Feiertage Datum bearbeiten
Script: [TeamsPhone_EditHoliday.ps1](TeamsPhone_EditHoliday.ps1) <br>
Der Skript bearbeitet die bereits definierten Urlaube im Teams Tenant mit neuen Daten (zB. für das neue Jahr). <br>
In der Variable `$Holidays` mit folgendem Syntax die Ferien / Feiertage, mit Datum und der Text Ansage angegeben werden.
```Powershell
@{
    Name = "Neujahr"
    Start = "2023-01-01T00:00:00"
    End = "2023-01-02T23:45:00"
},
```