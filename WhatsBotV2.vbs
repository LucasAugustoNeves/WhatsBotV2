set inicio=WScript.CreateObject("WScript.Shell")
inicio.run "chrome.exe"
WScript.Sleep 100

Set objFSO=CreateObject("Scripting.FileSystemObject")

sPathFile = "***Coloque-o-caminho-para-o-seu-arquivo-.txt***"

Set objFile = objFSo.OpenTextFile(sPathFile)
	Do Until objFile.AtEndOfStream
		sLine = objFile.ReadLine
		'wscript.echo sLine
			' Go To WhatsApp
			Set WshShell = WScript.CreateObject("WScript.Shell")
			Return = WshShell.Run(sLine, 1)
			' Loading Time
			' Go To The WhatsApp Search Bar
			WScript.Sleep 2000
			WshShell.SendKeys "{TAB}"
			WScript.Sleep 50
			WshShell.SendKeys "{TAB}"
			WScript.Sleep 50
			WshShell.SendKeys "{ENTER}"
			WScript.Sleep 15000
			WshShell.SendKeys "{ENTER}"
			WScript.Sleep 150
			WshShell.SendKeys "^1"
			WScript.Sleep 100
			WshShell.SendKeys "^w"
	Loop
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery _	
	("Select * from Win32_Process Where Name = 'chrome.exe'")
	
	For Each objProcess in colProcessList
	objProcess.Terminate()
	Next
objFile.Close