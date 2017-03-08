#10000001
#268435457
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{23186eda-2be8-462b-910c-ea951ea9ac0a}' -Name "DeviceState").DeviceState
[Microsoft.Win32.Registry]::GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{23186eda-2be8-462b-910c-ea951ea9ac0a}","DeviceState",$null)

(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{37c2c081-4710-4c0f-a98a-bbe6f330ddd8}' -Name "DeviceState").DeviceState
[Microsoft.Win32.Registry]::GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{37c2c081-4710-4c0f-a98a-bbe6f330ddd8}","DeviceState",$null)

[Microsoft.Win32.Registry]::SetValue('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{23186eda-2be8-462b-910c-ea951ea9ac0a}',"DeviceState","10000001")
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{23186eda-2be8-462b-910c-ea951ea9ac0a}' -Name "DeviceState" -Value 1