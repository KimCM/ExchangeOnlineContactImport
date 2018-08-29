# Exchange 365 - Import von Kontakten

Die E-Mail-Adressen der eigenen Active Directory Benutzer werden bei Exchange 365 automatisch dem Adressbuch hinzugefügt, so dass z.B. beim Versand einer E-Mail der Empfänger automatisch vorgeblendet wird.

Existieren jedoch mehrere Schwesterunternehmen, und sind deren Active Directories weder Teil eines gemeinsamen Forests noch gibt es directory synchronization oder ähnliche Mechanismen, dann bleibt zum Austausch von geschäftlichen Kontaktdaten manchmal nur die Möglichkeit, Dateien mit Kontaktdaten auszutauschen und diese Kontaktdaten dann als Kontakt zu importieren. Einen Vorschlag für den automatischen Export der eigenen Benutzerdaten aus dem Active Directory gibt es [hier](https://github.com/KimCM/ActiveDirectoryUserExport).

Das Script [user-import.ps1](user-import.ps1) importiert die Benutzerdaten aus der CSV-Datei als Exchange 365 Kontakte.

## Ausführung 

> Bei Verwendung von Exchange Online musst du einen [Connect mit PowerShell zu Exchange Online](https://technet.microsoft.com/library/jj984289.aspx) 
> herstellen. Bei Verwendung von Exchange on-premises kannst du zur Ausführung z.B. die Exchange Management Shell verwenden.

## Zeitsteuerung

Meiner Meinung macht es Sinn dieses Script zeitgesteuert z.B. alle 15 Minuten durch den Task Scheduler ausführen zu lassen. Die Dateien werden von einem UNC-Pfad gelesen, so dass andere Scripte (z.B. in den Schwesterfirmen) diese Dateien als Ziel ihres Exports nutzen können.

# Datenschutz

Du solltest vor der Bereitstellung die geltenden Bestimmungen zum Datenschutz beachten und das Script bei Bedarf anpassen.

## Customizing

Belege im Script `user-import.ps1` folgende Variablen mit sinnvollen Werten:

```ps1
$organizations = @("schwesterfirma1", "schwesterfirma2")
$share = "share"
$importFolder = "import"
```
z.B so:

```ps1
$organizations = @("schwesterfirma1", "schwesterfirma2")
$share = "\\fileserver\share"
$importFolder = "adexport"
```

## Optional: Erzeugen einer Adressliste pro Organisation

Werden viele Kontakte in Exchange importiert geht der Überblick leicht verloren. Es kann daher hilfreich sein, die Kontakte einer Organisation in Adresslisten zu gruppieren. Die Adressliste wird z.B. in Outlook vorgeblendet.

Eine solche Adressliste kann einmalig angelegt werden:

```ps
New-AddressList -Name 'schwesterfirma1' -RecipientFilter { Alias -ne $null -and ObjectCategory -like 'person' -and Company -like 'schwesterfirma1*' }
New-AddressList -Name 'schwesterfirma2' -RecipientFilter { Alias -ne $null -and ObjectCategory -like 'person' -and Company -like 'schwesterfirma2*' }
New-AddressList -Name 'schwesterfirma3' -RecipientFilter { Alias -ne $null -and ObjectCategory -like 'person' -and Company -like 'schwesterfirma3*' }
```

> Neu angelegte Adresslisten sind leer, selbst wenn du den RecipientFilter richtig konfiguriert hast. Exchange on-premises kennt das cmdlet 
> [Update-AddressList](https://docs.microsoft.com/en-us/powershell/module/exchange/email-addresses-and-address-books/update-addresslist?view=exchange-ps). Das gibt es in Exchange Online nicht mehr, s. auch [KB 2955640](https://support.microsoft.com/en-us/help/2955640/new-address-lists-that-you-create-in-exchange-online-don-t-contain-all).
> Daher kann es notwendig sein, dass du im Exchange Online nach jedem Import diese Adresslisten im Script löschen und neu erzeugen musst.

## Datensatzbeschreibung

| Feldname | Beschreibung des Inhalts | Beispiel |
| --- | --- | --- |
| `GivenName` | Vorname | Kim  |
| `Surname` | Nachname | Meiser |
| `DisplayName` | Anzeigename | Kim Meiser |
| `Initials` | Initialen / Kürzel | KimCM |
| `ObjectGUID` | Eindeutige Identifikation dieses Benutzers, wird idealerweise als Identifikation dieses Datensatzes verwendet | 5B4BE7E1-0F40-4DC7-98DE-07F6BF9CFDBE |
| `Title` | Stellenbezeichnung, Berufsbezeichnung | Chief Architect |
| `Department` | Abteilung | IT |
| `Company` | Organisation, Firma | Mega Business Inc. |
| `StreetAddress` | Straße inkl. Hausnummer | Innovation Street 2a |
| `PostalCode` | Postleitzahl | 12345 |
| `City` | Stadt | Springfield |
| `Country` | Land | DE |
| `Office` | Ort des Büros oder Bürobezeichnung, z.B. Zürich, Saarbrücken oder Frankfurt am Main | Saarbrücken West
| `EmailAddress` | E-Mail-Adresse | github@kimcm.de |
| `wwwHomePage` | Benutzer-Seite im Internet oder Intranet | http://kimcm.de
| `OfficePhone` | Telefon-Nummer Büro Festnetz | +49 681 555-555 |
| `MobilePhone` | Telefon-Nummer mobil | +49 555-56789 |
| `Fax` | Telefax-Nummer | +49 555-56780 |
| `ipPhone` | Skype ID oder SIP-Konto | Kim.Meiser@firmenname.de |