Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = scriptDir & "\dist\WRAITH_APP"
exePath = appDir & "\WRAITH_APP.exe"

WshShell.CurrentDirectory = appDir
WshShell.Run """" & exePath & """", 1, False
