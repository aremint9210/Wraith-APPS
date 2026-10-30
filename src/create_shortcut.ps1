$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut("C:\Users\aremi\Desktop\WRAITH.lnk")
$sc.TargetPath = "wscript.exe"
$sc.Arguments = "`"C:\Users\aremi\Desktop\REPORT CBM\SCRIPT\Run GUI.vbs`""
$sc.WorkingDirectory = "C:\Users\aremi\Desktop\REPORT CBM"
$sc.IconLocation = "C:\Users\aremi\Desktop\REPORT CBM\SCRIPT\app_icon.ico"
$sc.Save()
Write-Host "WRAITH shortcut created."
