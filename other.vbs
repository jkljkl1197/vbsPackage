Option Explicit
Dim objWMIService, objclass, objFSO, objFile, outputLine, outputLines, objNet
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNet = CreateObject("WScript.Network")
Set outputLines = CreateObject("System.Collections.ArrayList")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Set objFile = objFSO.CreateTextFile("P:\prog\exam\WMIclass.txt",True)
For Each objclass In objWMIService.SubclassesOf()
	If InStr(objClass.Path_.Class, "Win32_") > 0 Then
		outputLines.Add objClass.Path_.Class
	End If
Next
outputLines.Sort()

For Each outputLine In outputLines
	'WScript.Echo outputLine
	objFile.Write outputLine & vbCrLf
NEXT
 
objFile.Close
AddUser "Bonjour", "Administrateurs"
MapDefaultPrinter("\\Group\printer")

Function MapDefaultPrinter(chemin)

	On Error Resume Next
	Set WshNetwork = CreateObject("WScript.Network")
	Set rc = WshNetwork.AddWindowsPrinterConnection("" & chemin & "") '\\Group\printer
	If Not rc then
	      WScript.Echo("Printer Connection Failed!")
	End If
	WshNetwork.SetDefaultPrinter "" & chemin & ""
	
End Function

Function AddUser(strUser, sGroup)
	On Error Resume Next
	Dim strComputer, colAccounts, objUser, objLocalUser
	' Obtien le nom de l'ordinateur
    strComputer = objNet.Computername
	' crée un objet account
    Set colAccounts = GetObject("WinNT://" & strComputer & "") 
	' crée un objet du nouvelle utilisateur avec l'objet précédent
    Set objUser = colAccounts.Create("user", strUser)
	' defini le mot de passe utilisateur
    objUser.SetPassword "1234" 
    objUser.SetInfo
	' genere la fenetre
    Wscript.Echo "Added user " & strUser & " to " & strComputer 
	
    Set objLocalUser = GetObject("WinNT://" & strComputer & "/" & strUser) 
 
	intFlags = objLocalUser.GET("UserFlags") 
	intFlags = intFlags OR ADS_UF_DONT_EXPIRE_PASSWD 
	objLocalUser.Put "userFlags", intFlags  
	objLocalUser.SetInfo 
 
    Set objLocalAdmGroup = GetObject("WinNT://" & strComputer & "/" & sGroup & ",group")
    objLocalAdmGroup.Add(objLocalUser.AdsPath) 
    Wscript.Echo "Added user " & strUser & " to " & strComputer & "'s local admin group" 
End Function