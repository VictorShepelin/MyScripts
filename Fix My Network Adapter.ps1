Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_tcpip" -Enabled $true
Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_lltdio" -Enabled $true
Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_tcpip" -Enabled $false
Get-NetAdapter -Name "Ethernet" | Set-NetAdapterBinding -ComponentID "ms_lltdio" -Enabled $false