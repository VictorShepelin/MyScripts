Dim objShell
Set objShell=CreateObject("WScript.Shell")
strExpression="D:\OneDrive\My Scripts\Wallpapers\Wallpaper v1.5.ps1"
strCMD="powershell -NonInteractive -WindowStyle Hidden -NoLogo -file " & Chr(34) & strExpression & Chr(34)
objShell.Run strCMD,0
