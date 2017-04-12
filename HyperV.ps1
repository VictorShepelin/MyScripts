#region Переменные
 $hyperVHost = "pc10228"
 $vmName = "testtesla003"
 #$imageSource = "\\vs1026\deploymentshare$\Operating Systems\7 x64\sources\install.wim"
 $imageSource = "\\vs1026\deploymentshare$\Operating Systems\7 x86\sources\install.wim"
 $imagePath = "C:\Hyper-V"
 $image = "$imagePath\install.wim"
 $vhdx = "$imagePath\Virtual Hard Disks\$vmName.vhdx"
 #$unattend = "C:\Hyper-V\Unattend\7x64\Unattend.xml"
 $unattend = "C:\Hyper-V\Unattend\7x86\Unattend.xml"
 #$1stWave = "C:\Hyper-V\TeslaApps\packages64\1stWave"
 #$2ndWave = "C:\Hyper-V\TeslaApps\packages64\2ndWave"
 $1stWave = "C:\Hyper-V\TeslaApps\packages86\1stWave"
 $2ndWave = "C:\Hyper-V\TeslaApps\packages86\2ndWave"
 $appsPath = "C:\Hyper-V\TeslaApps\apps"
 $vmguestiso = "C:\Hyper-V\vmguest.iso"
 $psexecPath = "C:\Hyper-V\PsExec.exe" # обязательно проверьте запуск и согласитесь с EULA
 $password = "Answer42"
 . "\\vnxfs\powershell\Shepelin\C\Convert-WindowsImage.ps1"
 $pw = convertto-securestring -AsPlainText -Force -String $password
 $cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist "administrator", $pw
 #endregion
Function Copy-File {
    param( [string]$from, [string]$to)
    $ffile = [io.file]::OpenRead($from)
    $tofile = [io.file]::OpenWrite($to)
    Write-Progress -Activity "Copying file" -status "$from -> $to" -PercentComplete 0
    try {
        [byte[]]$buff = new-object byte[] 4096
        [long]$total = [int]$count = 0
        do {
            $count = $ffile.Read($buff, 0, $buff.Length)
            $tofile.Write($buff, 0, $count)
            $total += $count
            if ($total % 1mb -eq 0) {
                Write-Progress -Activity "Copying file" -status "$from -> $to" `
                   -PercentComplete ([int]($total/$ffile.Length* 100))
            }
        } while ($count -gt 0)
    }
    finally {
        $ffile.Dispose()
        $tofile.Dispose()
        Write-Progress -Activity "Copying file" -Status "Ready" -Completed
    }
 }
Function CheckVM ($hyperVHost, $vmName, $time) {
    Write-Host "Cheking $vmName..."
    if ((Get-VMIntegrationService -ComputerName $hyperVHost -VMName $vmName  | Where-Object {$_.name -eq "Heartbeat"}).PrimaryStatusDescription -ne "OK") {
        Start-VM -ComputerName $hyperVHost -VMName $vmName
        Write-Host "waiting..."
        Do {
            Start-Sleep 1
            }
        Until (((Get-VMIntegrationService -ComputerName $hyperVHost -VMName $vmName  | Where-Object {$_.name -eq "Heartbeat"}).PrimaryStatusDescription -eq "OK") -and ((get-service -Name winrm | Select-Object -ExpandProperty status) -eq "Running"))
        }
    $logon = $null
    if ($time -eq $null) {
        $time = (Get-Date).AddYears(-1)
        }
    Do {
        $logon = (Get-WinEvent -ComputerName $computer -FilterHashtable @{logname = ‘Microsoft-Windows-User Profile Service/Operational’; id = 2} -Credential $cred)[0] | Where-Object {
            ($_.TimeCreated -ge $time) -and ($_.Message -eq "Finished processing user logon notification on session 1.")
            } -ErrorAction SilentlyContinue
        }
    While ($logon -eq $null)
    Write-Host $logon.TimeCreated ":" $logon.Message
    $i = 0
    Do {
        $i++
        $error.Clear()
        Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
            $null
            } -ErrorAction SilentlyContinue
        if ($error.count -eq 0) {
            Write-Host "WinRM ready"
            $exit = 1
            } 
        else {
            Write-Host "WinRM not started, waiting. try = $i"
            if ($i -le 4) {
                Start-sleep 30
                }
            else {
                Start-sleep 10
                }
            }
        }
    While ($error.count -ne 0)
 }
Function InstallApp ($computer, $cred, $appPath, $arg) {
    Write-Host "Installing" $appPath $arg -ForegroundColor Cyan
    $errorCode = Invoke-Command -ComputerName $computer -Credential $cred -ArgumentList $appPath, $arg -ScriptBlock {
        param($appPath, $arg)
        $process = Start-Process $appPath -ArgumentList $arg -PassThru -Verb RunAs -Wait
        return $process
        }
    if ($errorCode.ExitCode -eq 0) {
        Write-Host "Success"
        }
    else {
        Write-Host "Error $errorCode" -ForegroundColor Red
        $errorCode | Select-Object *
        Read-Host "Admin!!! Stop the thing. Do the Thing" 
        }
 }
Function InstallAppWmi ($computer, $cred, $appPath, $arg, $log1, $log2, $matchStr) {
    Write-Host "Installing" $appPath $arg -ForegroundColor Cyan
    $errorCode = Invoke-Command -ComputerName $computer -Credential $cred -ArgumentList $appPath, $arg -ScriptBlock {
        param($appPath, $arg)
        $installString = "$appPath $arg"
        $process = ([WMICLASS]"\\localhost\root\cimv2:Win32_Process").Create($installString)
        While (Get-WmiObject Win32_Process -filter "ProcessID='$($process.ProcessID)'") {
            Start-Sleep 2
            }
        return $process
        }
    $content = get-content VM:\$log1
    if ($content -match $matchStr) {
        $str = $content -match $matchStr
        Write-Host $str -ForegroundColor Green
        }
    else {
        Write-Host "Error $errorCode" -ForegroundColor Red
        $errorCode | Select-Object *
        Read-Host "Admin!!! Stop the thing. Do the Thing"  
        }
    if ($log2 -ne $null) {
        $content = get-content VM:\$log2
        if ($content -match $matchStr) {
            $str = $content -match $matchStr
            Write-Host $str -ForegroundColor Green
            }
        else {
            Write-Host "Error $errorCode" -ForegroundColor Red
            $errorCode | Select-Object *
            Read-Host "Admin!!! Stop the thing. Do the Thing"  
            }
        }
 }
Function AddPackages ($path) {
    $packages = Get-ChildItem -Path $path -Recurse -Include "*.cab", "*.msu" | Sort-Object LastWriteTime
    $i = 1
    ForEach ($package in $packages) {
        $str = "($i of " + $packages.count + ") " + $package.Name
        Write-Host $str
        Add-WindowsPackage -Path "$imagePath\mount" -PackagePath $package.FullName | Out-Null
        if ($? -eq $TRUE) {
            $package.Name | Out-File -FilePath "$imagePath\Updates-Sucessful.log" -Append
            } 
        else {
            $package.Name | Out-File -FilePath "$imagePath\Updates-Failed.log" -Append
            }
        $i++
        }
 }
Function UpdateComputer ($computer, $cred) {
    Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
        $time = (Get-Date).AddSeconds(30).AddMinutes(1)
        if ($time.Hour -lt 10) {
            [string]$hour = "0" + $time.Hour
            }
        else {
            [string]$hour = $time.Hour
            }
        if ($time.Minute -lt 10) {
            [string]$min = "0" + $time.Minute
            }
        else {
            [string]$min = $time.Minute
            }
        $runtime = $hour + ":" + $min + ":00"
        $fullname = "c:\packages\psupdate.ps1"
        #$fullname = "`"" + $fullname + "`"" # Need to wrap in quotes as folder path may contain spaces путаница в выражении, слишком много кавычек, в итоге скрипт строго без пробелов
        $myCmd = "schtasks.exe /Delete /TN 'psupdate' /F"
        $result = invoke-expression "cmd.exe /c `"`"$myCmd 2>&1`"`""
        $result
        $myCmd = "schtasks.exe /Create /TN 'psupdate' /SC ONCE /ST $runtime /TR 'powershell.exe -WindowStyle hidden -executionpolicy bypass -file $fullname' /F"
        $result = invoke-expression "cmd.exe /c `"`"$myCmd 2>&1`"`""
        $result
        }
    }
#region slipstreaming image
 if (Test-Path $image) {
    Remove-Item $image -Force
    }
 if (Test-Path "$imagePath\mount") {
    Remove-Item "$imagePath\mount" -Force
    }
 New-Item -Path "$imagePath\mount"  -ItemType Directory
 Copy-File -from $imageSource -to $image
 Mount-WindowsImage -ImagePath $image -Index 1 -Path "$imagePath\mount"
 AddPackages $1stWave 
 AddPackages $2ndWave
 Dismount-WindowsImage –Path "$imagePath\mount" –Save
 #endregion
Convert-WindowsImage -SourcePath $image -VHDFormat VHDX -SizeBytes 128GB -VHDPartitionStyle MBR -VHDPath $vhdx -UnattendPath $unattend -passthru # -ShowUI
#region Создание VM
 $mem = 4096 #in mb
 New-VM -Name $vmName -MemoryStartupBytes ($mem*1024*1024) -Generation 1 -ComputerName $hyperVHost -VHDPath $vhdx -BootDevice "LegacyNetworkAdapter, IDE, CD, Floppy"
 Get-VM -ComputerName $hyperVHost -Name $vmName | Get-VMNetworkAdapter | Remove-VMNetworkAdapter # удаляем т.к. нужен legacy
 Get-VM -ComputerName $hyperVHost -Name $vmName | Add-VMNetworkAdapter -IsLegacy $true –Name "Legacy Network Adapter" -StaticMacAddress "00155df42700" -SwitchName "External"
 Set-VMProcessor $vmName -Count 2 -ComputerName $hyperVHost # -Reserve 10 -Maximum 75 -RelativeWeight 200
 Enable-VMIntegrationService -ComputerName $hyperVHost -VMName $vmName -Name *
 Set-VMDvdDrive -ComputerName $hyperVHost -VMName $vmName -Path $vmguestiso
 VMConnect.exe $hyperVHost $vmName
 $time = Get-Date
 Start-VM -ComputerName $hyperVHost -Name $vmName
 #endregion
#region Автоматически назначенная установка Integration Services из c:\Unattend.xml (68 строка) и получение IP адреса VM
 $i = 1
 $ip = $null
 do {
    Write-Host "Waiting for VM to get ready (attemp $i)"
    Start-Sleep 10
    $ip = (Get-VM -ComputerName $hyperVHost -Name $vmName | Get-VMNetworkAdapter).IPAddresses
    $ip[0]
    $i++
    }
 while ($ip[0] -eq $null)
 Get-VMDvdDrive -ComputerName $hyperVHost -VMName $vmName | Set-VMDvdDrive -Path $null
 #endregion
$computer = $ip[0]
 New-PSDrive -Name VM -Root \\$computer\c$\packages -Credential $cred -PSProvider FileSystem
 New-PSDrive -Name VMC -Root \\$computer\c$ -Credential $cred -PSProvider FileSystem
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "First Start" # создаём checkpoint
#region настройка WinRM на своей машине согласно IP VM
 try {
    enable-psremoting -force
    }
 catch {
    }
 Set-Item WSMan:\localhost\Client\TrustedHosts -Value $computer -Force
 #endregion

#region stage 1 - setup VM
 Write-Host "Stage 1 on $computer - setup VM" -ForegroundColor Cyan
 & $psexecPath \\$computer -u administrator -p $password -h -d powershell.exe -executionpolicy bypass -command "& {enable-psremoting -force; Set-Item WSMan:\localhost\Client\TrustedHosts -Value $computer -Force;Start-Service WinRM}"
 Do {
    $error.Clear()
    Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
        Set-Service -Name "RemoteRegistry" -StartupType Automatic #Включение службы RemoteRegistry
        Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -Name AutoConfigURL -Value "http://usrproxy.imb.ru/proxy-setup.pac" #Прокси для скачивания обновлений
        New-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name "AU" -Force | New-ItemProperty -Name "AUOptions" -PropertyType Dword -Value "2" -Force #Изменение Windows Update
        netsh advfirewall set allprofiles state off #Отключение firewall
        pkgmgr /iu:"TelnetClient" #Включение telnet
        New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS #Подключение ветки HKU
        Set-ItemProperty "HKU:\.Default\Control Panel\Keyboard" -Name InitialKeyboardIndicators -Value "2" #Установка NumLock
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value "0" #RDP
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "SecurityLayer" -Value "1" #RDP
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Value "0" #RDP
        New-ItemProperty -Path "HKU:\.Default\Keyboard Layout\Toggle" -Name "Hotkey" -PropertyType String -Value "2" 
        New-ItemProperty -Path "HKU:\.Default\Keyboard Layout\Toggle" -Name "Language Hotkey" -PropertyType String -Value "2" 
        New-ItemProperty -Path "HKU:\.Default\Keyboard Layout\Toggle" -Name "Layout Hotkey" -PropertyType String -Value "3" 
        Set-ItemProperty -Path "HKCU:\Keyboard Layout\Toggle" -Name "Hotkey" -Value "2" 
        Set-ItemProperty -Path "HKCU:\Keyboard Layout\Toggle" -Name "Language Hotkey" -Value "2" 
        Set-ItemProperty -Path "HKCU:\Keyboard Layout\Toggle" -Name "Layout Hotkey" -Value "3"
        $now = Get-Date -UFormat "%d.%m.%Y %H:%M:%S"
        New-Item -Path 'HKLM:\Software' -Name "UCB_PC_Image" -Force | New-ItemProperty -Name "Install Date" -PropertyType String -Value $now -Force
        New-Item -ItemType directory -path c:\UCBApps #Создание папки c:\UCBApps
        New-Item -ItemType directory -path c:\Temp #Создание папки c:\Temp
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "Model" -Value "Tesla 64 v.1.2 (march 2017 update)"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "SupportURL" -Value "http://Enter.ICT.unicredit.ru"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "Logo" -Value "\\ntd1\netlogon\logo.bmp"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "Manufacturer" -Value "AO UniCredit Bank"
        }
    if ($error.count -eq 0) {
        Write-Host "Success"
        $exit = 1
        } 
    else {
        Write-Host "Error, next try"
        Start-sleep 5
        & $psexecPath \\$computer -u administrator -p $password -h -d powershell.exe -executionpolicy bypass -command "& {enable-psremoting -force; Set-Item WSMan:\localhost\Client\TrustedHosts -Value $computer -Force;Start-Service WinRM}"
        }
    }
 While ($error.count -ne 0)
 #endregion
#region stage 2 - cleaning after deploy
 Write-Host "Stage 2 on $computer - cleaning after deploy" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Do {
    $error.Clear()
    Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
        Remove-Item -Path c:\Unattend.xml -Force # осталось после конвертации
        Remove-Item -Path c:\convert-windowsimageinfo.txt -Force # осталось после конвертации
        #([ADSI] "WinNT://localhost").Delete('User','admin') # осталось после конвертации
        }
    if ($error.count -eq 0) {
        Write-Host "Success"
        $exit = 1
        } 
    else {
        Write-Host "Error, next try"
        Start-sleep 5
        }
    }
 While ($error.count -ne 0)
 #endregion
#region stage 3 - setting Autologon
 Write-Host "Stage 3 on $computer - setting Autologon" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
    $password = "Answer42"
    New-ItemProperty -Path $regPath -PropertyType String -Name "DefaultPassword" -Value $password -Force
    Set-ItemProperty -Path $regPath -Name "AutoAdminLogon" -Value "1"
    Set-ItemProperty -Path $regPath -Name "DefaultUsername" -Value "Administrator"
    #Get-ItemProperty -Path $regPath -Name "AutoAdminLogon"
    #Get-ItemProperty -Path $regPath -Name "DefaultUsername"
    #Get-ItemProperty -Path $regPath -Name "DefaultPassword"
    }
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 $time = Get-Date
 #endregion
#region stage 4 - copying packages
 Write-Host "Stage 4 on $computer - copying packages" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 $packages = Get-ChildItem -Path $appsPath
 $packages | ForEach-Object {
    $destinationPath = "C:\packages\" + $_.Name
    Copy-VMFile -ComputerName $hyperVHost -Name $vmName -SourcePath $_.FullName -DestinationPath $destinationPath -FileSource Host -CreateFullPath -Force
    }
 #endregion
#region stage 5 - installing Microsoft Root Cert
 Write-Host "Stage 5 on $computer - installing Microsoft Root Cert" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Do {
    $error.Clear()
    Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
        $cert = new-object System.Security.Cryptography.X509Certificates.X509Certificate2
        $certPath = "C:\packages\MicrosoftRootCertificateAuthority2011.cer"
        $cert.import($certPath,"","Exportable,PersistKeySet") 
        $store = new-object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::Root, "localmachine")
        $store.open("MaxAllowed")
        $store.add($cert)
        $store.close()
        }
    if ($error.count -eq 0) {
        Write-Host "Success"
        $exit = 1
        } 
    else {
        Write-Host "Error, next try"
        Start-sleep 5
        }
    }
 While ($error.count -ne 0)
 #endregion
#region stage 6 - preparing to install .Net 4.6.2 via task scheduller
 Write-Host "Stage 6 on $computer - preparing to install .Net 4.6.2 via task scheduller" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    $time = (Get-Date).AddSeconds(30).AddMinutes(1)
    if ($time.Hour -lt 10) {
        [string]$hour = "0" + $time.Hour
        }
    else {
        [string]$hour = $time.Hour
        }
    if ($time.Minute -lt 10) {
        [string]$min = "0" + $time.Minute
        }
    else {
        [string]$min = $time.Minute
        }
    $runtime = $hour + ":" + $min + ":00"
    $fullname = "c:\packages\01_install_npd462.ps1"
    #$fullname = "`"" + $fullname + "`"" # Need to wrap in quotes as folder path may contain spaces путаница в выражении, слишком много кавычек, в итоге скрипт строго без пробелов
    $myCmd = "schtasks.exe /Create /TN '01 install npd462' /SC ONCE /ST $runtime /TR 'powershell.exe -WindowStyle hidden -executionpolicy bypass -file $fullname' /F"
    $result = invoke-expression "cmd.exe /c `"`"$myCmd 2>&1`"`""
    $result
    }
 #endregion
#region stage 7 - installing .Net 4.6.2
 Write-Host "Stage 7 on $computer - installing .Net 4.6.2" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Start-Sleep 60
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    Do {
        Start-Sleep 15
        }
    while ((Get-Process -Name NDP462-KB3151800-x86-x64-AllOS-ENU) -eq $null) # ждём когда планировщик запустится
    Write-Host "Идёт установка..."
    Do {
        Start-Sleep 5
        }
    while ((Get-Process -Name NDP462-KB3151800-x86-x64-AllOS-ENU) -ne $null) # ждём когда планировщик закончит (лучше на event завязаться)
    }
 $content = get-content VM:\ndp462.htm
 if ($content -match "The operation completed successfully") {
    $str = $content -match "The operation completed successfully"
    Write-Host $str -ForegroundColor Green
    }
 else {
    $str = $content -match "Error"
    Write-Host $str -ForegroundColor Red
    Read-Host "Admin!!! Stop the thing. Do the Thing"  
    }
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After Net462" # создаём checkpoint
#region stage 8 - installing PowerShell 5.1 via dism
 Write-Host "Stage 8 on $computer - installing PowerShell 5.1 via dism" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    #dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB2809215-x64.cab /NoRestart
    #dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB2872035-x64.cab /NoRestart
    #dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3033929-x64.cab /NoRestart
    #dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3191566-x64.cab /NoRestart
    dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB2872035-x86.cab /NoRestart
    dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3033929-x86.cab /NoRestart
    dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3191566-x86.cab /NoRestart
    }
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 $time = Get-Date
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 $events = $null
 Do {
    $events = (Get-WinEvent -ComputerName $computer -FilterHashtable @{logname = ‘setup’; id = 2}  -Credential $cred) | Where-Object {
            $_.Message -match "was successfully changed to the Installed state."
            }
    Start-Sleep 5
    }
 While (($events -contains "Package KB2809215 was successfully changed to the Installed state.") -and ($events -contains "Package KB2872035 was successfully changed to the Installed state.") -and ($events -contains "Package KB3033929 was successfully changed to the Installed state.") -and ($events -contains "Package KB3191566 was successfully changed to the Installed state."))
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After PS51" # создаём checkpoint
#region stage 9 - installing Windows Update Agent via dism
 Write-Host "Stage 9 on $computer - installing Windows Update Agent via dism" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    #dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3138612-x64.cab /NoRestart
    dism.exe /online /add-package /PackagePath:c:\packages\Windows6.1-KB3138612-x86.cab /NoRestart
    }
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 $time = Get-Date
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 $events = $null
 Do {
    $events = (Get-WinEvent -ComputerName $computer -FilterHashtable @{logname = ‘setup’; id = 2}  -Credential $cred) | Where-Object {
            $_.Message -match "was successfully changed to the Installed state."
            }
    Start-Sleep 5
    }
 While ($events -contains "Package KB3138612 was successfully changed to the Installed state.")
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After Windows Update Agent" # создаём checkpoint
#region stage 10 - installing Microsoft Visual C++ Redistributable Packages
 Write-Host "Stage 10 on $computer - installing Microsoft Visual C++ Redistributable Packages" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 #InstallApp -computer $computer -cred $cred -appPath "C:\packages\2005_vcredist_x64.exe" -arg "/Q"
 InstallApp -computer $computer -cred $cred -appPath "C:\packages\2005_vcredist_x86.exe" -arg "/Q"
 #InstallApp -computer $computer -cred $cred -appPath "C:\packages\2008_vcredist_x64.exe" -arg "/q"
 InstallApp -computer $computer -cred $cred -appPath "C:\packages\2008_vcredist_x86.exe" -arg "/q"
 #InstallApp -computer $computer -cred $cred -appPath "C:\packages\2010_vcredist_x64.exe" -arg "/q /norestart"
 InstallApp -computer $computer -cred $cred -appPath "C:\packages\2010_vcredist_x86.exe" -arg "/q /norestart"
 #InstallApp -computer $computer -cred $cred -appPath "C:\packages\2012_vcredist_x64.exe" -arg "/q /norestart"
 InstallApp -computer $computer -cred $cred -appPath "C:\packages\2012_vcredist_x86.exe" -arg "/q /norestart"
 #InstallApp -computer $computer -cred $cred -appPath "C:\packages\2013_vcredist_x64.exe" -arg "/q /norestart"
 InstallApp -computer $computer -cred $cred -appPath "C:\packages\2013_vcredist_x86.exe" -arg "/q /norestart"
 InstallAppWmi -computer $computer -cred $cred -appPath "C:\packages\2015_vc_redist.x86.exe" -arg "/q /norestart /log c:\packages\2015vs86.log" -log1 "2015vs86_000_vcRuntimeMinimum_x86.log" -log2 "2015vs86_001_vcRuntimeAdditional_x86.log" -matchStr "Installation completed successfully."
 #InstallAppWmi -computer $computer -cred $cred -appPath "C:\packages\2015_vc_redist.x64.exe" -arg "/q /norestart /log c:\packages\2015vs64.log" -log1 "2015vs64_000_vcRuntimeMinimum_x64.log" -log2 "2015vs64_001_vcRuntimeAdditional_x64.log" -matchStr "Installation completed successfully."
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After VC" # создаём checkpoint
#region stage 11 - installing Silverlight
 Write-Host "Stage 11 on $computer - installing Silverlight" -ForegroundColor Cyan
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 #InstallApp -computer $computer -cred $cred -appPath "c:\packages\Silverlight_x64.exe" -arg "/q"
 InstallApp -computer $computer -cred $cred -appPath "c:\packages\Silverlight.exe" -arg "/q"
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After Silverlight" # создаём checkpoint
#region stage 12 - Install all Windows Updates
 Write-Host "Stage 12 on $computer - Install all Windows Updates" -ForegroundColor Cyan
 Copy-VMFile -ComputerName $hyperVHost -Name $vmName -SourcePath "D:\Hyper-V\TeslaApps\apps\psupdate.ps1" -DestinationPath "C:\packages\psupdate.ps1" -FileSource Host -CreateFullPath -Force
 Copy-VMFile -ComputerName $hyperVHost -Name $vmName -SourcePath "D:\Hyper-V\TeslaApps\apps\PSWindowsUpdate.zip" -DestinationPath "C:\packages\PSWindowsUpdate.zip" -FileSource Host -CreateFullPath -Force
 # переделать скрипт 
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace(“C:\packages\PSWindowsUpdate.zip”)
    foreach($item in $zip.items()) {
        $shell.Namespace(“c:\windows\System32\WindowsPowerShell\v1.0\Modules”).copyhere($item)
    }
    New-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name "ElevateNonAdmins" -Force | New-ItemProperty -Name "ElevateNonAdmins" -PropertyType Dword -Value "1" -Force
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name "WUServer" -PropertyType String -Value "http://vs333.imb.ru:8530"
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate' -Name "WUStatusServer" -PropertyType String -Value "http://vs333.imb.ru"
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name "UseWUServer" -PropertyType Dword -Value "1"
    Restart-Service wuauserv
}
# получить значение "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" SusClientId "3eaa10ec-88fd-4947-bd83-389f7f0063cf"
 Get-Date
 UpdateComputer -computer $computer -cred $cred
 Write-Host "ждём ответа PSWindowsUpdate.log по установке обновлений"
 do {
    Start-Sleep 30
    }
 while (!(Test-Path VMC:\temp\PSWindowsUpdate.log))
 $i = 0
 do {
    $i++
    Write-Host "waiting for the PSWindowsUpdate.log (try = $i)"
    $lengthStart = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    # выводить инфу по размерам
    Start-Sleep 300
    $lengthEnd = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    # выводить инфу по размерам
    # если размер не изменился и нету пауршела в памяти, значит рестартануть службу винапдейт, чек susid и повторно запустить задачу из планировщика
    }
 while ($lengthStart -ne $lengthEnd -or $lengthEnd -eq 0)
 $i = 0
 Write-Host "ждём ответа WindowsUpdate.log по установке обновлений"
 do {
    $i++
    $WULog = Get-Content VMC:\Windows\WindowsUpdate.log
    $WULog | ForEach-Object {
        $currentLogTime = [datetime]$_.substring(0,19)
        if ($_ -match "Service: Service startup") {
            $markServiceStart = [datetime]$_.substring(0,19)
            }
        if ($_ -match "Reboot required = Yes") {
            $RebootMark = [datetime]$_.substring(0,19)
            }
        }
    if ($RebootMark -ge $markServiceStart) {
        Write-Host "требуется ребут в $RebootMark, последний ребут был $markServiceStart, ждём...(попытка $i)"
        }
    Start-Sleep 10
    }
 while ($RebootMark -ge $markServiceStart)
 Write-Host "ребут требовался в $RebootMark и уже ребутнулся $markServiceStart"
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Get-Date
 UpdateComputer -computer $computer -cred $cred
 #повторный запуск
 #делать пока нет апдейтов
 #endregion
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
Checkpoint-VM -Name $vmName -SnapshotName "After Update part 1" # создаём checkpoint
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
# вторая часть апдейтов
 Get-Date
 Write-Host "переименуем PSWindowsUpdate.log"
 if (Test-Path VMC:\temp\PSWindowsUpdate.log) {
     Rename-Item -Path VMC:\temp\PSWindowsUpdate.log -NewName PSWindowsUpdate_1.log
 }
 Write-Host "ждём ответа PSWindowsUpdate.log по установке обновлений"
 UpdateComputer -computer $computer -cred $cred
  do {
    Start-Sleep 30
    }
 while (!(Test-Path VMC:\temp\PSWindowsUpdate.log))
 $i = 0
 do {
    $i++
    Write-Host "waiting for the PSWindowsUpdate.log (try = $i)"
    $lengthStart = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthStart текущее значение"
    Start-Sleep 300
    $lengthEnd = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthEnd новое значение"
    # если размер не изменился и нету пауршела в памяти, значит рестартануть службу винапдейт, чек susid и повторно запустить задачу из планировщика
    }
 while ($lengthStart -ne $lengthEnd -or $lengthEnd -eq 0)
 $i = 0
 Write-Host "ждём ответа WindowsUpdate.log по установке обновлений"
 do {
    $i++
    $WULog = Get-Content VMC:\Windows\WindowsUpdate.log
    $WULog | ForEach-Object {
        $currentLogTime = [datetime]$_.substring(0,19)
        if ($_ -match "Service: Service startup") {
            $markServiceStart = [datetime]$_.substring(0,19)
            }
        if ($_ -match "Reboot required = Yes") {
            $RebootMark = [datetime]$_.substring(0,19)
            }
        }
    if ($RebootMark -ge $markServiceStart) {
        Write-Host "требуется ребут в $RebootMark, последний ребут был $markServiceStart, ждём...(попытка $i)"
        }
    Start-Sleep 10
    }
 while ($RebootMark -ge $markServiceStart)
 Write-Host "ребут требовался в $RebootMark и уже ребутнулся $markServiceStart"
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Checkpoint-VM -Name $vmName -SnapshotName "After Update part 2" # создаём checkpoint
# конец второй части апдейтов
# третья часть апдейтов
 Get-Date
 Write-Host "переименуем PSWindowsUpdate.log"
 if (Test-Path VMC:\temp\PSWindowsUpdate.log) {
     Rename-Item -Path VMC:\temp\PSWindowsUpdate.log -NewName PSWindowsUpdate_2.log
 }
 Write-Host "ждём ответа PSWindowsUpdate.log по установке обновлений"
 UpdateComputer -computer $computer -cred $cred
  do {
    Start-Sleep 30
    }
 while (!(Test-Path VMC:\temp\PSWindowsUpdate.log))
 $i = 0
 do {
    $i++
    Write-Host "waiting for the PSWindowsUpdate.log (try = $i)"
    $lengthStart = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthStart текущее значение"
    Start-Sleep 300
    $lengthEnd = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthEnd новое значение"
    # если размер не изменился и нету пауршела в памяти, значит рестартануть службу винапдейт, чек susid и повторно запустить задачу из планировщика
    }
 while ($lengthStart -ne $lengthEnd -or $lengthEnd -eq 0)
 $i = 0
 Write-Host "ждём ответа WindowsUpdate.log по установке обновлений"
 do {
    $i++
    $WULog = Get-Content VMC:\Windows\WindowsUpdate.log
    $WULog | ForEach-Object {
        $currentLogTime = [datetime]$_.substring(0,19)
        if ($_ -match "Service: Service startup") {
            $markServiceStart = [datetime]$_.substring(0,19)
            }
        if ($_ -match "Reboot required = Yes") {
            $RebootMark = [datetime]$_.substring(0,19)
            }
        }
    if ($RebootMark -ge $markServiceStart) {
        Write-Host "требуется ребут в $RebootMark, последний ребут был $markServiceStart, ждём...(попытка $i)"
        }
    Start-Sleep 10
    }
 while ($RebootMark -ge $markServiceStart)
 Write-Host "ребут требовался в $RebootMark и уже ребутнулся $markServiceStart"
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Checkpoint-VM -Name $vmName -SnapshotName "After Update part 3" # создаём checkpoint
# конец третьей части апдейтов
# четвертая часть апдейтов
 Get-Date
 Write-Host "переименуем PSWindowsUpdate.log"
 if (Test-Path VMC:\temp\PSWindowsUpdate.log) {
     Rename-Item -Path VMC:\temp\PSWindowsUpdate.log -NewName PSWindowsUpdate_3.log
 }
 Write-Host "ждём ответа PSWindowsUpdate.log по установке обновлений"
 UpdateComputer -computer $computer -cred $cred
  do {
    Start-Sleep 30
    }
 while (!(Test-Path VMC:\temp\PSWindowsUpdate.log))
 $i = 0
 do {
    $i++
    Write-Host "waiting for the PSWindowsUpdate.log (try = $i)"
    $lengthStart = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthStart текущее значение"
    Start-Sleep 300
    $lengthEnd = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthEnd новое значение"
    # если размер не изменился и нету пауршела в памяти, значит рестартануть службу винапдейт, чек susid и повторно запустить задачу из планировщика
    }
 while ($lengthStart -ne $lengthEnd -or $lengthEnd -eq 0)
 $i = 0
 Write-Host "ждём ответа WindowsUpdate.log по установке обновлений"
 do {
    $i++
    $WULog = Get-Content VMC:\Windows\WindowsUpdate.log
    $WULog | ForEach-Object {
        $currentLogTime = [datetime]$_.substring(0,19)
        if ($_ -match "Service: Service startup") {
            $markServiceStart = [datetime]$_.substring(0,19)
            }
        if ($_ -match "Reboot required = Yes") {
            $RebootMark = [datetime]$_.substring(0,19)
            }
        }
    if ($RebootMark -ge $markServiceStart) {
        Write-Host "требуется ребут в $RebootMark, последний ребут был $markServiceStart, ждём...(попытка $i)"
        }
    Start-Sleep 10
    }
 while ($RebootMark -ge $markServiceStart)
 Write-Host "ребут требовался в $RebootMark и уже ребутнулся $markServiceStart"
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Checkpoint-VM -Name $vmName -SnapshotName "After Update part 4" # создаём checkpoint
# конец четвертой части апдейтов
# 5ая часть апдейтов
 Get-Date
 Write-Host "переименуем PSWindowsUpdate.log"
 if (Test-Path VMC:\temp\PSWindowsUpdate.log) {
     Rename-Item -Path VMC:\temp\PSWindowsUpdate.log -NewName PSWindowsUpdate_4.log
 }
 Write-Host "ждём ответа PSWindowsUpdate.log по установке обновлений"
 UpdateComputer -computer $computer -cred $cred
  do {
    Start-Sleep 30
    }
 while (!(Test-Path VMC:\temp\PSWindowsUpdate.log))
 $i = 0
 do {
    $i++
    Write-Host "waiting for the PSWindowsUpdate.log (try = $i)"
    $lengthStart = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthStart текущее значение"
    Start-Sleep 300
    $lengthEnd = (Get-ItemProperty VMC:\temp\PSWindowsUpdate.log).Length
    Write-Host "$lengthEnd новое значение"
    # если размер не изменился и нету пауршела в памяти, значит рестартануть службу винапдейт, чек susid и повторно запустить задачу из планировщика
    }
 while ($lengthStart -ne $lengthEnd -or $lengthEnd -eq 0)
 $i = 0
 Write-Host "ждём ответа WindowsUpdate.log по установке обновлений"
 do {
    $i++
    $WULog = Get-Content VMC:\Windows\WindowsUpdate.log
    $WULog | ForEach-Object {
        $currentLogTime = [datetime]$_.substring(0,19)
        if ($_ -match "Service: Service startup") {
            $markServiceStart = [datetime]$_.substring(0,19)
            }
        if ($_ -match "Reboot required = Yes") {
            $RebootMark = [datetime]$_.substring(0,19)
            }
        }
    if ($RebootMark -ge $markServiceStart) {
        Write-Host "требуется ребут в $RebootMark, последний ребут был $markServiceStart, ждём...(попытка $i)"
        }
    Start-Sleep 10
    }
 while ($RebootMark -ge $markServiceStart)
 Write-Host "ребут требовался в $RebootMark и уже ребутнулся $markServiceStart"
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Stop-VM -ComputerName $hyperVHost -Name $vmName
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Checkpoint-VM -Name $vmName -SnapshotName "After Update part 5" # создаём checkpoint
# конец 5ой части апдейтов
# был ещё один апдейт - ручками сделал
Checkpoint-VM -Name $vmName -SnapshotName "After Update part 6" # создаём checkpoint
# добавил ручками через кашликова зибель
Checkpoint-VM -Name $vmName -SnapshotName "After Siebel" # создаём checkpoint
# удаление автологона и очистка
 CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
    Set-ItemProperty -Path $regPath -Name "DefaultPassword" -Value ""
    Set-ItemProperty -Path $regPath -Name "AutoAdminLogon" -Value "0"
    Set-ItemProperty -Path $regPath -Name "DefaultUsername" -Value ""
    Remove-Item -Path c:\packages -Recurse -Force
    Remove-Item -Path c:\temp -Recurse -Force
    New-Item -ItemType directory -path c:\Temp #Создание папки c:\Temp
    }
 Stop-VM -ComputerName $hyperVHost -Name $vmName
# конец удаления автологона
Checkpoint-VM -Name $vmName -SnapshotName "After cleaning" # создаём checkpoint
CheckVM -hyperVHost $hyperVHost -vmName $vmName -time $time
<#
 Restore-VMSnapshot -VMName $vmName -Name "First Start" -Confirm:$false
 Restore-VMSnapshot -VMName $vmName -Name "After Net462" -Confirm:$false
 Restore-VMSnapshot -VMName $vmName -Name "After PS51" -Confirm:$false
 Restore-VMSnapshot -VMName $vmName -Name "After Windows Update Agent" -Confirm:$false
 Restore-VMSnapshot -VMName $vmName -Name "After VC" -Confirm:$false
 Restore-VMSnapshot -VMName $vmName -Name "After Silverlight" -Confirm:$false


 Get-VMCheckpoint -VMName $vmName -Name "After PS51" | Remove-VMCheckpoint

 #region stage 3 - installing .Net 4.6.2 via System.Diagnostics
 Write-Host "Stage 3 on $computer - installing .Net 4.6.2 via System.Diagnostics"
 Invoke-Command -ComputerName $computer -Credential $cred -ScriptBlock {
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "C:\packages\NDP462-KB3151800-x86-x64-AllOS-ENU.exe"
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "/passive"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    Write-Host "stdout: $stdout"
    Write-Host "stderr: $stderr"
    Write-Host "exit code: " + $p.ExitCode
    }
 #endregion

 Stop-VM -ComputerName $hyperVHost -Name $vmName #-Save
 Remove-VM -ComputerName $hyperVHost –Name $vmName -Force
 Remove-Item -Path $vhdx -Force
 Copy-File -from 'C:\Hyper-V\Windows 7 x64 original + unattended.vhdx' -to $vhdx

 $start = get-date
 $end = Get-Date
 Write-Host ($end-$start).Minutes"min"($end-$start).Seconds"sec" -ForegroundColor Green
 Write-Host ($done-$end).Minutes"min"($done-$end).Seconds"sec" -ForegroundColor Green
 Write-Host ($done-$start).Minutes"min"($done-$start).Seconds"sec" -ForegroundColor Green

 Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse |
 Get-ItemProperty -name Version,Release -EA 0 |
 Where { $_.PSChildName -match '^(?!S)\p{L}'} |
 Select PSChildName, Version, Release, @{
  name="Product"
  expression={
      switch -regex ($_.Release) {
        "378389" { [Version]"4.5" }
        "378675|378758" { [Version]"4.5.1" }
        "379893" { [Version]"4.5.2" }
        "393295|393297" { [Version]"4.6" }
        "394254|394271" { [Version]"4.6.1" }
        "394802|394806" { [Version]"4.6.2" }
        {$_ -gt 394806} { [Version]"Undocumented 4.6.2 or higher, please update script" }
      }
    }
 }

 Examples Of ServiceID: 
 Windows Update 9482f4b4-e343-43b6-b170-9a65bc822c77 
 Microsoft Update 7971f918-a847-4430-9279-4a52d1efe18d 
 Windows Store 117cab2d-82b1-4b5a-a08c-4d62dbee7782 
 Windows Server Update Service 3da21691-e39d-4da6-8a4b-b43877bcb1b7 
 vs333 3eaa10ec-88fd-4947-bd83-389f7f0063cf
 ks133 9a1ad89a-584a-455b-8b8d-b936af2138d3
 #>