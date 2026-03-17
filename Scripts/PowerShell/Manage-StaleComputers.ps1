# ===== Logging (reuse same log file as other automation) =====
$LogFile = "C:\Scripts\JoinAutomation.log"

function Write-Log {
    param(
        [string]$Level,
        [string]$Component,
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "{0} {1,-5} {2,-15} {3}" -f $timestamp, $Level, $Component, $Message
    Add-Content -Path $LogFile -Value $line
}
# =============================================================

Import-Module ActiveDirectory

$domainDn    = (Get-ADDomain).DistinguishedName
$staleOuDn   = "OU=CORP-Stale-Computers,$domainDn"
$staleDays   = 15
$staleCutoff = (Get-Date).AddDays(-$staleDays)

Write-Host "=== Stale Computer Management ===" -ForegroundColor Cyan
Write-Host ("Stale threshold: {0:yyyy-MM-dd} (older than {1} days)" -f $staleCutoff, $staleDays) -ForegroundColor Cyan
Write-Log -Level "INFO" -Component "STALE-MANAGE" -Message ("Run started. Cutoff = {0:yyyy-MM-dd}" -f $staleCutoff)

# Get all computers with lastLogonTimestamp, whenCreated and OS
$computers = Get-ADComputer -Filter * -Properties lastLogonTimestamp,whenCreated,DistinguishedName,OperatingSystem

foreach ($comp in $computers) {

    $dn = $comp.DistinguishedName
    $os = $comp.OperatingSystem

    # 1) Skip Domain Controllers OU entirely
    if ($dn -like "OU=Domain Controllers,*") {
        continue
    }

    # 2) Skip computers already in Stale OU
    if ($dn -like "*,OU=CORP-Stale-Computers,$domainDn") {
        continue
    }

    # 3) OS presence check
    if ([string]::IsNullOrWhiteSpace($os)) {
        $msg = ("Skipping {0}: OperatingSystem not set." -f $comp.Name)
        Write-Host $msg -ForegroundColor Yellow
        Write-Log -Level "WARN" -Component "STALE-MANAGE" -Message $msg
        continue
    }

    # 4) Decide behavior based on OS
    if ($os -like "Windows 10*" -or $os -like "Windows 11*") {
        # Workstations – candidates for stale move
        $osCategory = "Workstation"
    }
    elseif (
        $os -like "Windows Server 2016*" -or
        $os -like "Windows Server 2019*" -or
        $os -like "Windows Server 2022*" -or
        $os -like "Windows Server 2025*"
    ) {
        # Servers 2016+ – log only, do not move
        $msg = ("Server {0} (OS: {1}) is older-OS candidate. Logging only, no stale move." -f $comp.Name, $os)
        Write-Host $msg -ForegroundColor Yellow
        Write-Log -Level "INFO" -Component "STALE-MANAGE" -Message $msg
        continue
    }
    else {
        # Non-Windows or older/unknown Windows – ignore
        $msg = ("Skipping {0}: unsupported or non-target OS ({1})." -f $comp.Name, $os)
        Write-Host $msg -ForegroundColor Yellow
        Write-Log -Level "WARN" -Component "STALE-MANAGE" -Message $msg
        continue
    }

    # 5) Determine last activity date (only for Windows 10/11, as above)
    $lastLogonDate = $null

    if ($comp.lastLogonTimestamp) {
        try {
            $lastLogonDate = [DateTime]::FromFileTime($comp.lastLogonTimestamp)
        }
        catch {
            # If conversion fails for some reason, fall back to whenCreated
            $lastLogonDate = $comp.whenCreated
        }
    }
    else {
        # Never logged on? Use whenCreated as best-effort
        $lastLogonDate = $comp.whenCreated
    }

    # Safety: if we still don't have a valid date, skip
    if (-not $lastLogonDate) {
        $msg = ("Skipping {0}: no valid lastLogonTimestamp or whenCreated." -f $comp.Name)
        Write-Host $msg -ForegroundColor Yellow
        Write-Log -Level "WARN" -Component "STALE-MANAGE" -Message $msg
        continue
    }

    # 6) Compare with cutoff
    if ($lastLogonDate -gt $staleCutoff) {
        # Active enough, ignore
        continue
    }

    # 7) Mark as stale and move to Stale OU (Windows 10/11 only)
    $msg = ("Marking {0} as stale (OS: {1}, last activity: {2:yyyy-MM-dd}), moving to {3}" -f $comp.Name, $os, $lastLogonDate, $staleOuDn)
    Write-Host $msg -ForegroundColor Green
    Write-Log -Level "INFO" -Component "STALE-MANAGE" -Message $msg

    try {
        Move-ADObject -Identity $dn -TargetPath $staleOuDn -ErrorAction Stop
    }
    catch {
        $errMsg = ("ERROR moving {0} to Stale OU: {1}" -f $comp.Name, $_.Exception.Message)
        Write-Host $errMsg -ForegroundColor Red
        Write-Log -Level "ERROR" -Component "STALE-MANAGE" -Message $errMsg
    }
}

Write-Log -Level "INFO" -Component "STALE-MANAGE" -Message "Run completed."
Write-Host "`nDone - stale computer check completed." -ForegroundColor Cyan
