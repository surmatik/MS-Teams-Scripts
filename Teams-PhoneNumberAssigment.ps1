# Teams-NumberAssigment.ps1
# Liste die alle Telefonnummern den Usern und die Sunrise Richtlinie zuweisst

# Lino Bertschinger, 14.06.2023

# TenantID des Kundens hier anpassen
# Wenn es der eigene Tenant ist kann man die Variable leer lassen
$TenantID = ""

# Name der Routing Policy für die User angeben
$RoutingPolicyName = "Sunrise Unlimited"

# Liste mit allen Benutzern
# Syntax: "MAILADRESSE" = "+41TELEFONNUMMER"
$UsersAndPhoneNumbers = @{
    "user1@surmatik.ch" = "+41441234501"
    "user2@surmatik.ch" = "+41441234502"
    "user3@surmatik.ch" = "+41441234503"
    "user4@surmatik.ch" = "+41441234504"
}


# Microsoft Teams Modul & Anmeldung
$Modules = Get-Module | Select -ExpandProperty Name
If ("MicrosoftTeams" -notin $Modules) { 
   If ($TenantID) {
      Connect-MicrosoftTeams -TenantId $TenantID
   } else {
      Connect-MicrosoftTeams
   }
}

# Schleife für alle Benutzer und Telefonnummern
foreach ($user in $UsersAndPhoneNumbers.Keys) {
    $phoneNr = $UsersAndPhoneNumbers[$user]
    
    # Zuweisen der Telefonnummer
    Set-CsPhoneNumberAssignment -Identity $user -PhoneNumber $phoneNr -PhoneNumberType DirectRouting

    # Gewähren der Online-Sprachroutingrichtlinie
    Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName $RoutingPolicyName
}