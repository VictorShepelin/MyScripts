#https://mcpmag.com/articles/2015/04/15/reporting-on-local-accounts.aspx
#http://stackoverflow.com/questions/203384/how-to-tell-when-windows-is-inactive
#https://www.codeproject.com/Articles/9104/How-to-check-for-user-inactivity-with-and-without
$discUsers = qwinsta | select-string "Disc" | select-string -notmatch "services"
$discUsers | ForEach-Object {
    logoff ($_.tostring() -split ' +')[2]
}