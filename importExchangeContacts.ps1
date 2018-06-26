$organizations = @("schwesterfirma1", "schwesterfirma2")
$share = "share"
$importFolder = "import"

# Abfrage manuell einzugebender Credentials
# $UserCredential = Get-Credential

# Setzen der Credentials unter Verwendung eines SecureString-Files.
# Erzeuge den SecureString einmalig durch manuelle Eingabe von 
# PS C:\> read-host -assecurestring | convertfrom-securestring | out-file adminsecurestring.txt
$username = "administrator@firmenname.onmicrosoft.com"
$password = Get-Content 'adminsecurestring.txt' | ConvertTo-SecureString
$UserCredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password

# Connect zu Office 365 Exchange
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession -Session $Session

# Ermitteln aller existierenden Kontakte in Exchange
$existingContacts = Get-Contact

foreach($organization in $organizations)
{
  Write-Host "Copy contacts from" $organization
  # Kopiere die CSV-Dateien in lokalen Import-Folder
  $fileName = ($organization + "-users.csv")
  $remotePath = Join-Path $share $filename
  $localPath = Join-Path $importFolder $filename
  $importFolderExist = Test-Path -Path $importFolder
  if ($importFolderExist -eq $false) { New-Item -Type Directory -Path $importFolder }
  Copy-Item $remotePath $localPath;
	
  # Import der Benutzer aus der CSV-Datei
  Write-Host "Processing contacts from" $organization
  $organizationUsers = Import-Csv $localPath
  $filterOrganizationStar = "$organization*"

  foreach($u in $organizationUsers)
  {
    # Anlage des Mailkontakts falls noch nicht in Kontaktliste enthalten
    if ($u.ObjectGUID -notin $existingContacts.Identity)
    {
      Write-Host "  Try to create mail contact" $u.DisplayName
      New-MailContact -Name $u.ObjectGUID -DisplayName $u.DisplayName -ExternalEmailAddress $u.EmailAddress -FirstName $u.GivenName -LastName $u.Surname
    }
    # Aktualisierung des Mailkontakts
    Write-Host "  Try to update mail contact" $u.DisplayName
    Set-Contact -Identity $u.ObjectGUID -FirstName $u.GivenName -LastName $u.Surname -DisplayName $u.DisplayName -StreetAddress $u.StreetAddress `
    -City $u.City -StateorProvince $u.StateorProvince -Country $u.Country -PostalCode $_.PostalCode -WindowsEmailAddress $u.EmailAddress -Phone $u.OfficePhone -MobilePhone $u.MobilePhone -Company $u.Company -Title $u.Title -Department $u.Department -Fax $u.Fax -Initials $u.Initials -Office $u.Office
  }

  # Ermitteln der Differenz zwischen existierenden Kontakten und Inhalt der CSV-Datei mit dem Ziel, gelöschte User aus den Kontakten zu entfernen
  $organizationContactsToRemove = Get-Contact | Where { $_.Company -like $filterOrganizationStar -and $_.Identity -notin $organizationUsers.ObjectGuid }
  foreach($removeUser in $organizationContactsToRemove)
  {
    Write-Host "  Try to delete mail contact" $removeUser.DisplayName
    Remove-MailContact -Identity $removeUser.Identity -Confirm:$false
  }
}

# Schließen der Office 365 Exchange-Session
Remove-PSSession -Session $Session