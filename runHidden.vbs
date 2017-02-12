Dim objShell
Set objShell=CreateObject("WScript.Shell")
strExpression="D:\OneDrive\Git\Wallpaper Changer.ps1"
strCMD="powershell -NonInteractive -WindowStyle Hidden -NoLogo -file " & Chr(34) & strExpression & Chr(34)
objShell.Run strCMD,0
