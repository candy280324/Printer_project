Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "D:\django_printer_project\print.bat" & chr(34), 0
Set WshShell = Nothing
