# ================================
# Hyper-V VM Provisioning Script
# ================================

#Change VMName, VMSwitch, ISOPath, VMPath, CPU, RAM, Disk


param(
    [string]$VMName = "Test-VM",
    [int]$RAMGB = 4,
    [int]$CPU = 2,
    [int]$DiskGB = 50,
    [string]$VMSwitch = "Default Switch",
    [string]$ISOPath = "D:\ISO\windows.iso",
    [string]$VMPath = "D:\HyperV\VMs"
)

# Convert values
$RAM = $RAMGB * 1GB
$DiskSize = $DiskGB * 1GB

Write-Host "Creating VM: $VMName"

# Create Dynamic Disk
$VHDPath = "$VMPath\$VMName\$VMName.vhdx"
New-Item -ItemType Directory -Path "$VMPath\$VMName" -Force | Out-Null

New-VHD -Path $VHDPath -SizeBytes $DiskSize -Dynamic

# Create VM
New-VM -Name $VMName `
       -MemoryStartupBytes $RAM `
       -Generation 2 `
       -VHDPath $VHDPath `
       -Path $VMPath `
       -SwitchName $VMSwitch

# Configure CPU
Set-VMProcessor -VMName $VMName -Count $CPU

# Enable Dynamic Memory (optional)
Set-VM -Name $VMName -DynamicMemory `
       -MemoryMinimumBytes 1GB `
       -MemoryMaximumBytes ($RAMGB * 2GB)

# Attach ISO
Add-VMDvdDrive -VMName $VMName -Path $ISOPath

# Set Boot Order
Set-VMFirmware -VMName $VMName -FirstBootDevice (Get-VMDvdDrive -VMName $VMName)

# Start VM
Start-VM -Name $VMName

Write-Host "VM $VMName created and started successfully!"