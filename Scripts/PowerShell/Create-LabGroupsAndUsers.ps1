Import-Module ActiveDirectory

$domainDn = (Get-ADDomain).DistinguishedName

# --- 1. OUs we expect (adjust names if needed) ---
$ouUsersDn = "OU=CORP-Users,$domainDn"
$ouWorkstationsDn = "OU=CORP-Workstations,$domainDn"

Write-Host "Using CORP-Users OU: $ouUsersDn" -ForegroundColor Cyan

# --- 2. Create some core groups ---
$groups = @(
    @{ Name = "GG-Workstation-Admins";  Scope = "Global";  Category = "Security" },
    @{ Name = "GG-Helpdesk";            Scope = "Global";  Category = "Security" }
)

foreach ($g in $groups) {
    if (-not (Get-ADGroup -Filter "SamAccountName -eq '$($g.Name)'" -ErrorAction SilentlyContinue)) {
        New-ADGroup -Name $g.Name `
                    -SamAccountName $g.Name `
                    -GroupScope $g.Scope `
                    -GroupCategory $g.Category `
                    -Path $ouUsersDn
        Write-Host "Created group: $($g.Name)" -ForegroundColor Green
    }
    else {
        Write-Host "Group already exists: $($g.Name)" -ForegroundColor Yellow
    }
}
# --- 3. Bulk create 10 lab users ---
$securePassword = ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force

for ($i = 1; $i -le 10; $i++) {
    $sam = "labuser$i"
    $name = "Lab User $i"
    $upn  = "$sam@corp.lab"

    if (-not (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue)) {
        New-ADUser `
            -SamAccountName $sam `
            -UserPrincipalName $upn `
            -Name $name `
            -GivenName "Lab" `
            -Surname "User$i" `
            -Path $ouUsersDn `
            -AccountPassword $securePassword `
            -Enabled $true `
            -PasswordNeverExpires $true

        Write-Host "Created user: $sam" -ForegroundColor Green

        # Add first 3 lab users to Helpdesk group as example
        if ($i -le 3) {
            Add-ADGroupMember -Identity "GG-Helpdesk" -Members $sam
            Write-Host "  -> added $sam to GG-Helpdesk" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "User already exists: $sam" -ForegroundColor Yellow
    }
}

Write-Host "Done. Check CORP-Users OU for groups and users." -ForegroundColor Cyan