Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd C:\Program Files niszczycielzabawy.bat"
oShell.Run strArgs, 0, false