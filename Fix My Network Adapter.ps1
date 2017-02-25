
if ((Get-NetAdapter -Name "Ethernet" | Get-NetAdapterBinding -ComponentID "ms_tcpip" | Select-Object -ExpandProperty Enabled) -eq $false) {
    Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_tcpip" -Enabled $true
    }
if ((Get-NetAdapter -Name "Ethernet" | Get-NetAdapterBinding -ComponentID "ms_lltdio" | Select-Object -ExpandProperty Enabled) -eq $false) {
    Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_lltdio" -Enabled $true
    }