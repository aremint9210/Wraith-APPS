Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\aremi\Desktop\REPORT CBM"
WshShell.Run "pythonw.exe ""C:\Users\aremi\Desktop\REPORT CBM\SCRIPT\app_gui.py""", 0, False
