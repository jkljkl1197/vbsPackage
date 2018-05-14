Dim inConsole
Set oShell = WScript.CreateObject ("WScript.Shell")
Const Quote = """"
inConsole = false

Set objSWbemServices = GetObject ("WinMgmts:Root\Cimv2") 
Set colProcess = objSWbemServices.ExecQuery ("Select * From Win32_Process where name = 'cscript.exe'") 
For Each objProcess In colProcess
	inConsole = true
Next

IF inConsole = true Then
	Set conteneur = GetObject("winmgmts:\\.\root\cimv2")
	Set maclasse = conteneur.ExecQuery("Select * from Win32_Processor")
	For Each processeur in maclasse
		For Each propriete in processeur.properties_
			WScript.Echo propriete.name
		Next
	Next
	WScript.Quit
ElSE
	oShell.Run "cmd /k cscript " & Quote & Wscript.ScriptName & Quote & " ", 1, true
	Set oShell = Nothing
	WScript.Quit
END If