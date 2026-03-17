Import-Module ActiveDirectory

# ================== Config ==================
$LogFile        = "C:\Scripts\JoinAutomation.log"
$StaleDays      = 60
$StaleCutoff    = (Get-Date).AddDays(-$StaleDays)
$HtmlReportPath = "C:\Scripts\ADHealthDashboard.html"
$Domain         = Get-ADDomain
$DomainDn       = $Domain.DistinguishedName
$DomainName     = $Domain.DNSRoot

$OuWorkstations = "OU=CORP-Workstations,$DomainDn"
$OuServers      = "OU=CORP-Servers,$DomainDn"
$OuStale        = "OU=CORP-Stale-Computers,$DomainDn"
$OuDomainCtrl   = "OU=Domain Controllers,$DomainDn"
$CnComputers    = "CN=Computers,$DomainDn"

# Thresholds
$PwdVeryOldDays = 365
$today          = Get-Date
$cut15Days      = $today.AddDays(-15)
$cut30Days      = $today.AddDays(-30)
$cut45Days      = $today.AddDays(-45)
# ============================================

Write-Host "=== AD Health & Security Dashboard ===" -ForegroundColor Cyan
Write-Host ("Domain: {0}" -f $DomainDn) -ForegroundColor Cyan
Write-Host ("Stale threshold (computers): lastLogon older than {0:yyyy-MM-dd} ({1} days)" -f $StaleCutoff, $StaleDays) -ForegroundColor Cyan
Write-Host ""

# ===== Helper: lastLogonTimestamp -> DateTime =====
function Get-LastLogonDate {
    param(
        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADComputer]$Computer
    )

    if ($Computer.lastLogonTimestamp) {
        try {
            return [DateTime]::FromFileTime($Computer.lastLogonTimestamp)
        }
        catch {
            return $null
        }
    }

    return $null
}

# ===== Helper: safe FromFileTime =====
function Convert-FileTimeToDate {
    param(
        [Parameter(Mandatory)]
        [long]$FileTime
    )
    try {
        return [DateTime]::FromFileTime($FileTime)
    }
    catch {
        return $null
    }
}

# ================== COMPUTER SECTION ==================

# ===== 1. AD Computer Overview =====
$allComputers = Get-ADComputer -Filter * -Properties Enabled, DistinguishedName, OperatingSystem, lastLogonTimestamp

if ($null -eq $allComputers) {
    $allComputers = @()
}
elseif ($allComputers -isnot [System.Collections.IEnumerable] -or $allComputers -is [string]) {
    $allComputers = @($allComputers)
}

$totalComputers    = [int]$allComputers.Count
$enabledComputers  = [int](@($allComputers | Where-Object { $_.Enabled -eq $true }).Count)
$disabledComputers = [int](@($allComputers | Where-Object { $_.Enabled -eq $false }).Count)

$inComputersCn    = [int](@($allComputers | Where-Object { $_.DistinguishedName -like "CN=*,CN=Computers,$DomainDn" }).Count)
$inWorkstationsOu = [int](@($allComputers | Where-Object { $_.DistinguishedName -like "CN=*,$OuWorkstations" }).Count)
$inServersOu      = [int](@($allComputers | Where-Object { $_.DistinguishedName -like "CN=*,$OuServers" }).Count)
$inStaleOu        = [int](@($allComputers | Where-Object { $_.DistinguishedName -like "CN=*,$OuStale" }).Count)
$inDcOu           = [int](@($allComputers | Where-Object { $_.DistinguishedName -like "CN=*,$OuDomainCtrl" }).Count)

Write-Host "=== 1) Computer Overview ===" -ForegroundColor Yellow
"{0,-30} {1,5}" -f "Total computers:",                $totalComputers
"{0,-30} {1,5}" -f "Enabled:",                        $enabledComputers
"{0,-30} {1,5}" -f "Disabled:",                       $disabledComputers
"{0,-30} {1,5}" -f "In CN=Computers:",                $inComputersCn
"{0,-30} {1,5}" -f "In CORP-Workstations:",           $inWorkstationsOu
"{0,-30} {1,5}" -f "In CORP-Servers:",                $inServersOu
"{0,-30} {1,5}" -f "In CORP-Stale-Computers:",        $inStaleOu
"{0,-30} {1,5}" -f "In Domain Controllers OU:",       $inDcOu
Write-Host ""

# ===== 2. OS Distribution =====
$win10Count = [int](@($allComputers | Where-Object { $_.OperatingSystem -like "Windows 10*" }).Count)
$win11Count = [int](@($allComputers | Where-Object { $_.OperatingSystem -like "Windows 11*" }).Count)
$serverCount = [int](@($allComputers | Where-Object {
        $_.OperatingSystem -like "Windows Server 2016*" -or
        $_.OperatingSystem -like "Windows Server 2019*" -or
        $_.OperatingSystem -like "Windows Server 2022*" -or
        $_.OperatingSystem -like "Windows Server 2025*"
    }).Count)

$otherOsCount = [int]($totalComputers - ($win10Count + $win11Count + $serverCount))

Write-Host "=== 2) OS Distribution ===" -ForegroundColor Yellow
"{0,-25} {1,5}" -f "Windows 10:",             $win10Count
"{0,-25} {1,5}" -f "Windows 11:",             $win11Count
"{0,-25} {1,5}" -f "Windows Server 2016+:",   $serverCount
"{0,-25} {1,5}" -f "Other / Unknown OS:",     $otherOsCount
Write-Host ""

# ===== 3. Stale Summary (computers) =====
$staleWorkstations = @()
$staleServers      = @()

foreach ($comp in $allComputers) {
    $os = $comp.OperatingSystem

    if ([string]::IsNullOrWhiteSpace($os)) {
        continue
    }

    $lastLogonDate = Get-LastLogonDate -Computer $comp
    if (-not $lastLogonDate) {
        continue
    }

    if ($lastLogonDate -gt $StaleCutoff) {
        continue
    }

    if ($os -like "Windows 10*" -or $os -like "Windows 11*") {
        $staleWorkstations += $comp
    }
    elseif (
        $os -like "Windows Server 2016*" -or
        $os -like "Windows Server 2019*" -or
        $os -like "Windows Server 2022*" -or
        $os -like "Windows Server 2025*"
    ) {
        $staleServers += $comp
    }
}

Write-Host ("=== 3) Stale Summary (cutoff: {0:yyyy-MM-dd}) ===" -f $StaleCutoff) -ForegroundColor Yellow
"{0,-35} {1,5}" -f "Stale workstations (Win10/11):",      $staleWorkstations.Count
"{0,-35} {1,5}" -f "Stale servers (2016+ candidates):",   $staleServers.Count
"{0,-35} {1,5}" -f "Currently in CORP-Stale-Computers:",  $inStaleOu
Write-Host ""

# ===== 4. Automation Health from JoinAutomation.log =====
Write-Host "=== 4) Automation Health (recent log entries) ===" -ForegroundColor Yellow

if (Test-Path -Path $LogFile) {
    $logContent = Get-Content -Path $LogFile -ErrorAction SilentlyContinue

    $components = @("OU-MOVE", "RDP-GROUP", "STALE-MANAGE", "STALE-RESTORE")

    foreach ($compName in $components) {
        Write-Host ("-- {0} (last 5 entries) --" -f $compName) -ForegroundColor Cyan

        $entries = $logContent |
            Where-Object { $_ -match $compName } |
            Select-Object -Last 5

        if ($entries) {
            $entries | ForEach-Object { Write-Host "  $_" }
        }
        else {
            Write-Host "  (no entries found)" -ForegroundColor DarkGray
        }

        Write-Host ""
    }
}
else {
    Write-Host "Log file not found: $LogFile" -ForegroundColor Red
}
Write-Host ""

# ================== USER SECTION ==================

Write-Host "=== 5) User Overview (created / disabled / deleted) ===" -ForegroundColor Yellow

$allUsers = Get-ADUser -Filter * -Properties Enabled, whenCreated, msDS-UserPasswordExpiryTimeComputed, PasswordNeverExpires, PasswordLastSet, LockedOut, memberOf

$totalUsers    = [int]$allUsers.Count
$enabledUsers  = [int](@($allUsers | Where-Object { $_.Enabled -eq $true }).Count)
$disabledUsers = [int](@($allUsers | Where-Object { $_.Enabled -eq $false }).Count)

# Creation buckets (whenCreated)
$created0to15  = $allUsers | Where-Object { $_.whenCreated -gt $cut15Days }
$created16to30 = $allUsers | Where-Object { $_.whenCreated -le $cut15Days -and $_.whenCreated -gt $cut30Days }
$created31to45 = $allUsers | Where-Object { $_.whenCreated -le $cut30Days -and $_.whenCreated -gt $cut45Days }

# Disabled list
$disabledUsersList = $allUsers | Where-Object { $_.Enabled -eq $false }

# Deleted users (tombstones) and age buckets [tombstones use whenChanged as deletion time] [web:102]
$deletedUsers = Get-ADObject -Filter 'isDeleted -eq $true -and objectClass -eq "user"' `
    -IncludeDeletedObjects -Properties lastKnownParent, whenChanged

$deleted0to15  = $deletedUsers | Where-Object { $_.whenChanged -gt $cut15Days }
$deleted16to30 = $deletedUsers | Where-Object { $_.whenChanged -le $cut15Days -and $_.whenChanged -gt $cut30Days }
$deleted31to45 = $deletedUsers | Where-Object { $_.whenChanged -le $cut30Days -and $_.whenChanged -gt $cut45Days }

Write-Host ("Total users:                         {0}" -f $totalUsers)
Write-Host ("Enabled users:                       {0}" -f $enabledUsers)
Write-Host ("Disabled users:                      {0}" -f $disabledUsers)
Write-Host ("Users created 0-15 days:             {0}" -f $created0to15.Count)
Write-Host ("Users created 16-30 days:            {0}" -f $created16to30.Count)
Write-Host ("Users created 31-45 days:            {0}" -f $created31to45.Count)
Write-Host ("Deleted users (0-15 days):           {0}" -f $deleted0to15.Count)
Write-Host ("Deleted users (16-30 days):          {0}" -f $deleted16to30.Count)
Write-Host ("Deleted users (31-45 days):          {0}" -f $deleted31to45.Count)
Write-Host ""

$recentCreatedSample = $created0to15 |
    Sort-Object whenCreated -Descending |
    Select-Object -First 10 -Property SamAccountName, whenCreated

$disabledUsersSample = $disabledUsersList |
    Sort-Object SamAccountName |
    Select-Object -First 10 -Property SamAccountName

$deletedUsersSample = $deletedUsers |
    Sort-Object whenChanged -Descending |
    Select-Object -First 10 Name, lastKnownParent, whenChanged

# Locked-out users
$lockedUsers = $allUsers | Where-Object { $_.LockedOut -eq $true }

Write-Host "=== 6) Locked-out Users ===" -ForegroundColor Yellow
Write-Host ("Locked-out users:                    {0}" -f $lockedUsers.Count)
Write-Host ""

# Password last set buckets and expiry
Write-Host "=== 7) Password Last Changed & Expiry (users) ===" -ForegroundColor Yellow

$usersWithPwdLastSet = $allUsers | Where-Object { $_.PasswordLastSet -ne $null }

$pwdChanged0to15  = $usersWithPwdLastSet | Where-Object { $_.PasswordLastSet -gt $cut15Days }
$pwdChanged16to30 = $usersWithPwdLastSet | Where-Object { $_.PasswordLastSet -le $cut15Days -and $_.PasswordLastSet -gt $cut30Days }
$pwdChanged31to45 = $usersWithPwdLastSet | Where-Object { $_.PasswordLastSet -le $cut30Days -and $_.PasswordLastSet -gt $cut45Days }

Write-Host ("Passwords changed 0-15 days:         {0}" -f $pwdChanged0to15.Count)
Write-Host ("Passwords changed 16-30 days:        {0}" -f $pwdChanged16to30.Count)
Write-Host ("Passwords changed 31-45 days:        {0}" -f $pwdChanged31to45.Count)

# Very old passwords
$pwdVeryOldCut = $today.AddDays(-$PwdVeryOldDays)
$pwdVeryOld = $usersWithPwdLastSet | Where-Object { $_.PasswordLastSet -lt $pwdVeryOldCut }

Write-Host ("Passwords older than {0} days:       {1}" -f $PwdVeryOldDays, $pwdVeryOld.Count)

# Password expiry (using msDS-UserPasswordExpiryTimeComputed) [web:58][web:111]
$usersWithExpiry = $allUsers | Where-Object {
    $_.Enabled -eq $true -and
    $_.PasswordNeverExpires -ne $true -and
    $_.'msDS-UserPasswordExpiryTimeComputed'
}

$usersExpiry = $usersWithExpiry | Select-Object `
    SamAccountName,
    @{Name='PasswordExpiryDate';Expression={ Convert-FileTimeToDate $_.'msDS-UserPasswordExpiryTimeComputed' }}

$expiringIn30Days = $usersExpiry | Where-Object {
    $_.PasswordExpiryDate -gt $today -and
    $_.PasswordExpiryDate -le $today.AddDays(30)
}
$alreadyExpired = $usersExpiry | Where-Object {
    $_.PasswordExpiryDate -le $today
}

Write-Host ("Total enabled users with expiring passwords: {0}" -f $usersExpiry.Count)
Write-Host ("Passwords expiring in next 30 days:        {0}" -f $expiringIn30Days.Count)
Write-Host ("Passwords already expired:                 {0}" -f $alreadyExpired.Count)
Write-Host ""

$topExpirySample = $usersExpiry |
    Where-Object { $_.PasswordExpiryDate -gt $today } |
    Sort-Object PasswordExpiryDate |
    Select-Object -First 10

# PasswordNeverExpires
$pwdNeverExpires = $allUsers | Where-Object { $_.PasswordNeverExpires -eq $true }
Write-Host ("Users with PasswordNeverExpires set:      {0}" -f $pwdNeverExpires.Count)
Write-Host ""

# ================== SECURITY / PRIVILEGED GROUPS ==================

Write-Host "=== 8) Privileged Groups ===" -ForegroundColor Yellow

$privGroups = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators",
    "Account Operators",
    "Server Operators",
    "Backup Operators"
)

$privGroupResults = @()

foreach ($gName in $privGroups) {
    $group = Get-ADGroup -Filter "Name -eq '$gName'" -ErrorAction SilentlyContinue
    if ($group) {
        $members = Get-ADGroupMember $group -Recursive -ErrorAction SilentlyContinue
        $count   = ($members | Where-Object { $_.objectClass -eq "user" }).Count
        Write-Host ("{0,-25} members: {1}" -f $gName, $count)
        $privGroupResults += [PSCustomObject]@{
            Name        = $gName
            MemberCount = $count
        }
    }
    else {
        Write-Host ("{0,-25} group not found" -f $gName)
        $privGroupResults += [PSCustomObject]@{
            Name        = $gName
            MemberCount = 0
        }
    }
}
Write-Host ""

# ================== DC HEALTH & DNS ==================

Write-Host "=== 9) Domain Controller Health (basic) ===" -ForegroundColor Yellow

$dcs = Get-ADDomainController -Filter * | Sort-Object Name
$dcHealth = @()

foreach ($dc in $dcs) {
    $dcName = $dc.HostName
    $os     = $dc.OperatingSystem
    $site   = $dc.Site
    $ip     = $dc.IPv4Address

    # Basic uptime via CIM (ignore errors)
    $uptimeDays = $null
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $dcName -ErrorAction SilentlyContinue
        if ($osInfo.LastBootUpTime) {
            $uptimeDays = [int]((Get-Date) - $osInfo.LastBootUpTime).TotalDays
        }
    } catch {}

    $dcHealth += [PSCustomObject]@{
        Name       = $dcName
        OS         = $os
        Site       = $site
        IPv4       = $ip
        UptimeDays = $uptimeDays
    }
}

foreach ($row in $dcHealth) {
    Write-Host ("{0,-20} {1,-25} Uptime: {2} days" -f $row.Name, $row.OS, ($row.UptimeDays -as [string]))
}
Write-Host ""

# Quick DNS SRV test for DC locator record [web:14][web:81]
Write-Host "=== 10) DNS SRV Check (_ldap._tcp.dc._msdcs) ===" -ForegroundColor Yellow
$dnsSrvOk = $false
try {
    $srvRecords = Resolve-DnsName -Type SRV "_ldap._tcp.dc._msdcs.$DomainName" -ErrorAction Stop
    if ($srvRecords) {
        $dnsSrvOk = $true
        Write-Host ("SRV records found: {0}" -f $srvRecords.Count)
    }
}
catch {
    Write-Host "Failed to resolve SRV record _ldap._tcp.dc._msdcs.$DomainName" -ForegroundColor Red
}
Write-Host ""

# Optional: dcdiag summary hook (just run /c /q for each DC and track failures) [web:131]
$dcDiagIssues = 0
foreach ($dc in $dcs) {
    try {
        $out = & dcdiag /s:$($dc.HostName) /c /q 2>$null
        if ($LASTEXITCODE -ne 0 -or $out) {
            $dcDiagIssues++
        }
    } catch {}
}

Write-Host "=== 11) dcdiag Summary ===" -ForegroundColor Yellow
Write-Host ("DCs with potential dcdiag issues: {0}" -f $dcDiagIssues)
Write-Host ""

# ================== STATUS INDICATOR ==================

$overallStatus = "OK"
$issues = @()

if ($staleWorkstations.Count -gt 0 -or $staleServers.Count -gt 0) {
    $issues += "Stale computers present"
}
if ($dcDiagIssues -gt 0 -or -not $dnsSrvOk) {
    $issues += "Domain controller / DNS issues"
}
if ($pwdVeryOld.Count -gt 0 -or $pwdNeverExpires.Count -gt 0) {
    $issues += "Password hygiene issues"
}
if ($lockedUsers.Count -gt 0) {
    $issues += "Locked-out users present"
}
if (($privGroupResults | Where-Object { $_.MemberCount -gt 5 }).Count -gt 0) {
    $issues += "Large privileged groups"
}

if ($issues.Count -gt 0) {
    if ($issues.Count -le 2) {
        $overallStatus = "Warning"
    }
    else {
        $overallStatus = "Critical"
    }
}

Write-Host "=== 12) Overall Status ===" -ForegroundColor Yellow
Write-Host ("Status: {0}" -f $overallStatus)
if ($issues.Count -gt 0) {
    Write-Host "Issues:"
    $issues | ForEach-Object { Write-Host (" - {0}" -f $_) }
}
else {
    Write-Host "No major issues detected."
}
Write-Host ""

# ================== HTML REPORT ==================

Write-Host ("Generating HTML report: {0}" -f $HtmlReportPath) -ForegroundColor Cyan

$html = @()
$html += "<html><head><title>AD Health Dashboard - $DomainDn</title>"
$html += "<style>body{font-family:Segoe UI,Arial;font-size:12px;} h2{color:#003366;} table{border-collapse:collapse;} th,td{border:1px solid #ccc;padding:4px 8px;} th{background:#e0e7f1;}</style>"
$html += "</head><body>"
$html += "<h2>AD Health & Security Dashboard - $DomainDn</h2>"
$html += "<p>Generated: $(Get-Date)</p>"

# Status
$html += "<h3>Status</h3>"
$html += "<p><b>Overall status:</b> $overallStatus</p>"
if ($issues.Count -gt 0) {
    $html += "<ul>"
    foreach ($i in $issues) {
        $html += "<li>$i</li>"
    }
    $html += "</ul>"
}

# Section 1 – Computer Overview
$html += "<h3>1) Computer Overview</h3>"
$html += "<table>"
$html += "<tr><th>Metric</th><th>Value</th></tr>"
$html += "<tr><td>Total computers</td><td>$totalComputers</td></tr>"
$html += "<tr><td>Enabled</td><td>$enabledComputers</td></tr>"
$html += "<tr><td>Disabled</td><td>$disabledComputers</td></tr>"
$html += "<tr><td>In CN=Computers</td><td>$inComputersCn</td></tr>"
$html += "<tr><td>In CORP-Workstations</td><td>$inWorkstationsOu</td></tr>"
$html += "<tr><td>In CORP-Servers</td><td>$inServersOu</td></tr>"
$html += "<tr><td>In CORP-Stale-Computers</td><td>$inStaleOu</td></tr>"
$html += "<tr><td>In Domain Controllers OU</td><td>$inDcOu</td></tr>"
$html += "</table>"

# Section 2 – OS Distribution
$html += "<h3>2) OS Distribution</h3>"
$html += "<table>"
$html += "<tr><th>OS Category</th><th>Count</th></tr>"
$html += "<tr><td>Windows 10</td><td>$win10Count</td></tr>"
$html += "<tr><td>Windows 11</td><td>$win11Count</td></tr>"
$html += "<tr><td>Windows Server 2016+</td><td>$serverCount</td></tr>"
$html += "<tr><td>Other / Unknown</td><td>$otherOsCount</td></tr>"
$html += "</table>"

# Section 3 – Stale summary (computers)
$html += "<h3>3) Stale Summary (computers, cutoff: $($StaleCutoff.ToString('yyyy-MM-dd')))</h3>"
$html += "<p>Stale workstations: $($staleWorkstations.Count) | Stale servers: $($staleServers.Count) | In Stale OU: $inStaleOu</p>"

# Section 4 – User Overview (created / disabled / deleted)
$html += "<h3>4) User Overview (created / disabled / deleted)</h3>"
$html += "<table>"
$html += "<tr><th>Metric</th><th>Value</th></tr>"
$html += "<tr><td>Total users</td><td>$totalUsers</td></tr>"
$html += "<tr><td>Enabled users</td><td>$enabledUsers</td></tr>"
$html += "<tr><td>Disabled users</td><td>$disabledUsers</td></tr>"
$html += "<tr><td>Users created 0-15 days</td><td>$($created0to15.Count)</td></tr>"
$html += "<tr><td>Users created 16-30 days</td><td>$($created16to30.Count)</td></tr>"
$html += "<tr><td>Users created 31-45 days</td><td>$($created31to45.Count)</td></tr>"
$html += "<tr><td>Deleted users 0-15 days</td><td>$($deleted0to15.Count)</td></tr>"
$html += "<tr><td>Deleted users 16-30 days</td><td>$($deleted16to30.Count)</td></tr>"
$html += "<tr><td>Deleted users 31-45 days</td><td>$($deleted31to45.Count)</td></tr>"
$html += "<tr><td>Locked-out users</td><td>$($lockedUsers.Count)</td></tr>"
$html += "</table>"

# Section 5 – Password Last Changed & Expiry
$html += "<h3>5) Password Last Changed & Expiry</h3>"
$html += "<table>"
$html += "<tr><th>Bucket</th><th>Count</th></tr>"
$html += "<tr><td>Password changed 0-15 days</td><td>$($pwdChanged0to15.Count)</td></tr>"
$html += "<tr><td>Password changed 16-30 days</td><td>$($pwdChanged16to30.Count)</td></tr>"
$html += "<tr><td>Password changed 31-45 days</td><td>$($pwdChanged31to45.Count)</td></tr>"
$html += "<tr><td>Password &gt; $PwdVeryOldDays days old</td><td>$($pwdVeryOld.Count)</td></tr>"
$html += "<tr><td>PasswordNeverExpires users</td><td>$($pwdNeverExpires.Count)</td></tr>"
$html += "</table>"

$html += "<p>Total enabled users with expiring passwords: $($usersExpiry.Count)</p>"
$html += "<p>Passwords expiring in next 30 days: $($expiringIn30Days.Count)</p>"
$html += "<p>Passwords already expired: $($alreadyExpired.Count)</p>"

# Section 6 – Privileged Groups
$html += "<h3>6) Privileged Groups</h3>"
$html += "<table>"
$html += "<tr><th>Group</th><th>User members</th></tr>"
foreach ($pg in $privGroupResults) {
    $html += "<tr><td>$($pg.Name)</td><td>$($pg.MemberCount)</td></tr>"
}
$html += "</table>"

# Section 7 – DC Health
$html += "<h3>7) Domain Controller Health</h3>"
$html += "<table>"
$html += "<tr><th>DC</th><th>OS</th><th>Site</th><th>IPv4</th><th>Uptime (days)</th></tr>"
foreach ($row in $dcHealth) {
    $html += "<tr><td>$($row.Name)</td><td>$($row.OS)</td><td>$($row.Site)</td><td>$($row.IPv4)</td><td>$($row.UptimeDays)</td></tr>"
}
$html += "</table>"

# Section 8 – DNS / dcdiag
$html += "<h3>8) DNS SRV / dcdiag Summary</h3>"
$html += "<p>_ldap._tcp.dc._msdcs SRV resolution OK: $dnsSrvOk</p>"
$html += "<p>DCs with dcdiag issues (non-silent): $dcDiagIssues</p>"

$html += "</body></html>"

$html -join "`r`n" | Set-Content -Path $HtmlReportPath -Encoding UTF8

Write-Host "Done – dashboard shown above and HTML report generated." -ForegroundColor Green
Write-Host ("Open in browser: {0}" -f $HtmlReportPath) -ForegroundColor Green
