#https://blogs.technet.microsoft.com/heyscriptingguy/2010/06/10/hey-scripting-guy-how-can-i-log-out-a-user-if-they-launch-a-particular-application/
<# 
.Synopsis 
Log out users from their workstations when a particular program launches. 
.Description 
Log out users from their workstations when a particular program launches. 
.Parameter Program 
The program that will trigger the user to log out. The default is calc.exe 
.Parameter Timeout 
Time in seconds the user will have to save their work before being logged off 
The default is 60 seconds. 
.Parameter Force 
By default we will not log off any user that has a Word document or Excel spreadsheet open. To 
Override this setting and log off regardless of any open application, apply 
the force parameter. 
.Example 
./Start-Logout -Program “calc.exe” -Timeout 90 
.Example 
./Start-Logout -Program “notepad.exe” -Timeout 15 -force 
#> 
[CmdletBinding(SupportsShouldProcess=$false,ConfirmImpact=”low”)] 
Param( 
[Parameter(Mandatory=$False, ValueFromPipelinebyPropertyName=$True)] 
[String] 
$Program = “calc.exe”, 
[Parameter(Mandatory=$False, ValueFromPipelinebyPropertyName=$True)] 
[Int] 
$Timeout = 60, 
[Parameter(Mandatory=$False)] 
[Switch] 
$Force 
) 
Begin 
{ 
# Create our event trigger 
Remove-Event -SourceIdentifier LogoffTrigger -ea SilentlyContinue 
Unregister-Event -SourceIdentifier LogoffTrigger -ea SilentlyContinue 
Try 
{ 
Register-WmiEvent -SourceIdentifier LogoffTrigger -Query @” 
SELECT * 
FROM __InstanceCreationEvent 
WITHIN 5 
WHERE TargetInstance ISA ‘Win32_Process’ 
AND TargetInstance.Name = ‘$Program’ 
“@ 
} 
Catch 
{ 
Write-Warning $_.Exception.Message 
exit; 
} 
#Helper function that will handle the system tray notification to the user. 
Function Wait-NotifyUser([int]$Timeout) 
{ 
# remove any previous events 
Unregister-Event notification_event_close -ea SilentlyContinue 
Remove-Event notification_event_close -ea SilentlyContinue 
#Initialize our balloon tip 
[void][System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | out-null 
[void][System.Reflection.Assembly]::LoadWithPartialName(“System.Timers”) 
$notification = new-object System.Windows.Forms.NotifyIcon 
$notification.Icon = [System.Drawing.SystemIcons]::Exclamation 
$notification.BalloonTipTitle = “Pending logout notification!” 
$notification.Visible = $True 
# register a new event on the balloon Tip Clicked Closed 
Register-ObjectEvent $notification BalloonTipClicked notification_event_close 
$global:loop = $true 
while ($Timeout -gt 0) 
{ 
$notification.BalloonTipText = “You will be logged out in $($Timeout) Seconds. Please save all work immediately” 
$notification.ShowBalloonTip(150000) 
# sleep until either the balloon tip is clicked or 1 second elapses. 
Wait-Event -Timeout 1 -SourceIdentifier notification_event_close | out-null 
Remove-Event notification_event_close -ErrorAction SilentlyContinue 
$Timeout– 
} 
Remove-Event notification_event_close -ErrorAction SilentlyContinue 
Unregister-Event notification_event_close 
$notification.Dispose() 
} 
# Converted some C# to Windows PowerShell to use PInvoke to log off the user. 
# There is no advantage to this over using WMI/Shutdown.exe… 
# Just wanted to highlight the trick, and show off add-type’s power. 
Function Exit-UserSession([switch]$Force) { 
$Win32ExitWindowsEx = Add-Type -Name ‘Win32ExitWindowsEx’ ` 
-namespace ‘Win32Functions’ ` 
-memberDefinition @” 
[DllImport(“user32.dll”)] 
public static extern int ExitWindowsEx(int uFlags, int dwReason); 
“@ -passThru 
IF ($Force) 
{ 
$Win32ExitWindowsEx::ExitWindowsEx(10,0) 
} 
Else 
{ 
$Win32ExitWindowsEx::ExitWindowsEx(0,0) 
} 
} 
} 
End 
{ 
# wait here until the specified app is launched 
Wait-Event -SourceIdentifier LogoffTrigger | out-null 
Write-Verbose “$($Program) Launch detected!” 
IF (!$Force) 
{ 
While (Get-WmiObject -Class Win32_Process -Filter “Name = ‘winword.exe’ OR Name = ‘excel.exe'”) 
{ 
Start-Sleep -Seconds 5 
} 
} 
Wait-NotifyUser -timeout $Timeout 
Exit-UserSession 
}