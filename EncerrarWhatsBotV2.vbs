strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name = 'wscript.exe'")

For Each objProcess in colProcessList
 ObjProcess.Terminate()
Next