Import-Module ActiveDirectory
Get-ADUser corpadmin -Properties memberOf | Select-Object -ExpandProperty memberOf
