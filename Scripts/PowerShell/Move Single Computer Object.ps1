Import-Module ActiveDirectory

Move-ADObject `
  -Identity "CN=WIN10-CLI03,CN=Computers,DC=corp,DC=lab" `
  -TargetPath "OU=CORP-Workstations,DC=corp,DC=lab"
