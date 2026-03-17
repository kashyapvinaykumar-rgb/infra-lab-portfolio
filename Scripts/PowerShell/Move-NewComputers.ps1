
# ===== Logging =====
$LogFile = "C:\Scripts\JoinAutomation.log"

function Write-Log {
    param(
        [string]$Level,
        [string]$Component,
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "{0} {1,-5} {2,-10} {3}" -f $timestamp, $Level, $Component, $Message
    Add-Content -Path $LogFile -Value $line
}
# ====================


Import-Module ActiveDirectory

$domainDn           = (Get-ADDomain).DistinguishedName
$computersContainer = "CN=Computers,$domainDn"
$workstationsOuDn   = "OU=CORP-Workstations,$domainDn"
$serversOuDn        = "OU=CORP-Servers,$domainDn"      # target OU for servers
$rdpGroupName       = "GG-RDP-Workstations"

Write-Log -Level "INFO" -Component "SCRIPT" -Message "Run started."

Write-Host "=== Stage 1: Move computers to role-based OUs ===" -ForegroundColor Cyan

# Move Windows 10/11 workstations and Windows Server machines
$computers = Get-ADComputer -SearchBase $computersContainer -SearchScope OneLevel `
    -Filter * -Properties OperatingSystem

foreach ($comp in $computers) {
    $os = $comp.OperatingSystem

    if ([string]::IsNullOrWhiteSpace($os)) {
        Write-Host ("Skipping {0}: OperatingSystem not set yet." -f $comp.Name) -ForegroundColor Yellow
        continue
    }

    try {
        if ($os -like "Windows 10*" -or $os -like "Windows 11*") {
            # Workstations -> CORP-Workstations
            Write-Host ("Moving {0} (OS: {1}) to {2}" -f $comp.Name, $os, $workstationsOuDn) -ForegroundColor Green
            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $workstationsOuDn -ErrorAction Stop
        }
        elseif ($os -like "Windows Server*") {
            # Servers -> CORP-Servers
            Write-Host ("Moving {0} (OS: {1}) to {2}" -f $comp.Name, $os, $serversOuDn) -ForegroundColor Green
            Move-ADObject -Identity $comp.DistinguishedName -TargetPath $serversOuDn -ErrorAction Stop
        }
        else {
            Write-Host ("Leaving {0} in Computers (OS: {1})" -f $comp.Name, $os) -ForegroundColor Gray
        }
    }
    catch {
        Write-Host ("  ERROR moving {0}: {1}" -f $comp.Name, $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host "`n=== Stage 2: Ensure RDP group membership (GG-RDP-Workstations) ===" -ForegroundColor Cyan

# Users that must be in the RDP group
$rdpUsersRequired = @(
    "labuser1",
    "labuser2",
    "labuser3"
    "labuser4",
    "labuser5",
    "labuser6"
    "labuser7",
    "labuser8",
    "labuser9"
)

# Get the group
try {
    $rdpGroup = Get-ADGroup -Identity $rdpGroupName -ErrorAction Stop
}
catch {
    Write-Host ("ERROR: RDP group '{0}' not found. Create it first." -f $rdpGroupName) -ForegroundColor Red
    return
}

# Current user members
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

Write-Host "`nDone - computers moved and RDP group membership validated." -ForegroundColor Cyan

Write-Log -Level "INFO" -Component "SCRIPT" -Message "Run completed."

