' Launches XLS Virtual Printer completely silent with no console window
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.Run """" & strDir & "\bin\XlsVirtualPrinter.exe""", 1, False

