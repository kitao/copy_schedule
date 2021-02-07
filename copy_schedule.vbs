Dim shell
Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = Split(WScript.ScriptFullName, Wscript.ScriptName)(0)
shell.run "cmd /c copy_schedule.bat", 0
