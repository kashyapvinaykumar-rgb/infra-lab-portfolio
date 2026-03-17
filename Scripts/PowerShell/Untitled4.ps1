# Move-NewComputers.ps1
# Purpose:
#  1) Move Windows 10/11 computers from default Computers container to CORP-Workstations OU
#     - If object is ProtectedFromAccidentalDeletion, temporarily unprotect -> move -> re-protect
#  2) Ensure required users are in GG-RDP-Workstations (for RDP access)

Import-Module ActiveDirectory

$domainDn           = (Get-ADDomain).DistinguishedName
$computersContainer = "CN=Computers,$domainDn"
$workstationsOuDn   = "OU=CORP-Workstations,$domainDn"
$rdpGroupName       = "GG-RDP-Workstations"

Write-Host "=== Stage 1: Move Windows 10/11 computers to CORP-Workstations ===" -ForegroundColor Cyan

# 1. Move Windows 10/11 computers (with temporary unprotect/reprotect if needed)
$computers = Get-ADComputer -SearchBase $computersContainer -SearchScope OneLevel `
    -Filter * -Properties OperatingSystem,ProtectedFromAccidentalDeletion

foreach ($comp in $computers) {
    $os = $comp.OperatingSystem

    if ([string]::IsNullOrWhiteSpace($os)) {
        Write-Host ("Skipping {0}: OperatingSystem not set yet." -f $comp.Name) -ForegroundColor Yellow
        continue
    }

    if ($os -like "Windows 10*" -or $os -like "Windows 11*") {
        Write-Host ("Moving {0} (OS: {1}) to {2}" -f $comp.Name, $os, $workstationsOuDn) -ForegroundColor Green

        # Option B logic: remember if it was protected, temporarily unprotect, move, then re-protect
        $wasProtected = $false
        try {
            $adObj = Get-ADObject -Identity $comp.DistinguishedName -Properties ProtectedFromAccidentalDeletion
            $wasProtected = [bool]$adObj.ProtectedFromAccidentalDeletion

            if ($wasProtected) {
                Write-Host ("  {0} is protected; temporarily disabling protection..." -f $comp.Name) -ForegroundColor Yellow
                Set-ADObject -Identity $comp.DistinguishedName -ProtectedFromAccidentalDeletion:$false
            }

            # Move the object
            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $workstationsOuDn -ErrorAction Stop

            # Re-protect at new location if it was protected before
            if ($wasProtected) {
                $moved = Get-ADComputer -Identity $comp.SamAccountName
                Set-ADObject -Identity $moved.DistinguishedName -ProtectedFromAccidentalDeletion:$true
                Write-Host ("  Re-applied protection on {0} in new OU." -f $comp.Name) -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host ("  ERROR moving {0}: {1}" -f $comp.Name, $_.Exception.Message) -ForegroundColor Red
        }
    }
    else {
        Write-Host ("Leaving {0} in Computers (OS: {1})" -f $comp.Name, $os) -ForegroundColor Gray
    }
}

Write-Host "`n=== Stage 2: Ensure RDP group membership (GG-RDP-Workstations) ===" -ForegroundColor Cyan

# 2. Ensure core users are in GG-RDP-Workstations
#    Update these SamAccountNames to your real accounts:
$rdpUsersRequired = @(
    "labuser2",      # example admin account
    "labuser3",      # example user account
    "labuser1"
)

# Get group object
try {
    $rdpGroup = Get-ADGroup -Identity $rdpGroupName -ErrorAction Stop
}
catch {
    Write-Host ("ERROR: RDP group '{0}' not found. Create it first." -f $rdpGroupName) -ForegroundColor Red
    return
}

# Current user members of that group
$currentMembers = Get-ADGroupMember -Identity $rdpGroupName -Recursive `
    | Where-Object { $_.ObjectClass -eq 'user' } `
    | Select-Object -ExpandProperty SamAccountName

foreach ($user in $rdpUsersRequired) {
    if ($currentMembers -contains $user) {
        Write-Host ("User {0} already in {1}" -f $user, $rdpGroupName) -ForegroundColor Yellow
    }
    else {
        try {
            Add-ADGroupMember -Identity $rdpGroupName -Members $user -ErrorAction Stop
            Write-Host ("Added {0} to {1} (for RDP)" -f $user, $rdpGroupName) -ForegroundColor Green
        }
        catch {
            Write-Host ("  ERROR adding {0}: {1}" -f $user, $_.Exception.Message) -ForegroundColor Red
        }
    }
}

Write-Host "`nDone – computers moved (with protection preserved) and RDP group membership validated." -ForegroundColor Cyan
