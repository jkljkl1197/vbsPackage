Set DomainInfo = CreateObject("AdSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")
WScript.Echo "Bienvenue sur le domaine " & DomainInfo.DomainDNSName
Effacer()
' WScript.Quit

' config
Dim Server
Server = "10.99.64.91" ' ip or hostname

Dim BoolProfesseurs, BoolEtudiants, BoolTechniques
BoolProfesseurs = MembreDuGroupe("Professeurs")
BoolEtudiants = MembreDuGroupe("Etudiants")
BoolTechniques = MembreDuGroupe("Techniques")

'wscript.echo BoolProfesseurs
'wscript.echo BoolEtudiants
'wscript.echo BoolTechniques

If BoolProfesseurs = 1 Or BoolEtudiants = 1 Then
	If BoolProfesseurs = 1 Then
		Map "P:", "\\" & Server & "\Prof\" & objNetwork.UserName & ""
		Map "Q:", "\\" & Server & "\Departement"
		MapDefaultPrinter "\\" & Server & "\NewLaserJet PCL5"
	ElseIf BoolEtudiants = 1 Then
		Map "P:", "\\" & Server & "\Etudiant\" & objNetwork.UserName & ""
		MapDefaultPrinter "\\" & Server & "\OldLaserJetIV"
	End If
	
	Map "D:", "\\" & Server & "\Distribution"
	Map "R:", "\\" & Server & "\Reception"
End If

If BoolProfesseurs = 1 Or BoolTechniques = 1 Then
	Map "M:", "\\" & Server & "\Documentation"
	Map "N:", "\\" & Server & "\Applications"
End If


driveList = Split(Afficher(), ":")
Dim strMsg, strDriveName
For Each objDrive in driveList
	strMsg = strMsg & objDrive & ": "  & vbCrlF
Next
MsgBox strMsg

'-------------------------------------------------------------------------------------
Function MembreDuGroupe(groupe)

	Dim objRootLDAP, objGroup, objUser, objOU, objmemberOf
	Dim strOU, strUser, strDNSDomain, strLDAP, strList, retour
	
	' Commands to bind to AD and extract domain name
	Set objRootLDAP = GetObject("LDAP://RootDSE")
	strDNSDomain = objRootLDAP.Get("DefaultNamingContext")
	
	' Build the LDAP DN from strUser, strOU and strDNSDomain
	strUser ="cn=" & objNetwork.UserName & ","
	strOU ="CN=Users,"
	strLDAP ="LDAP://" & strUser & strOU & strDNSDomain
	
	Set objUser = GetObject(strLDAP)
	
	' Heart of the script, extract a list of Groups from memberOf
	objmemberOf  = objUser.GetEx("memberOf")
	retour = 0
	For Each objGroup in objmemberOf
	   objGroup = Split(objGroup, "=")(1)
	   objGroup = Split(objGroup, ",")(0)
	   If objGroup = groupe Then
	   		retour = 1
	   End If
	   strList = strList & objGroup & vbcr
	Next
	
	'WScript.Echo "Groups for " & strUser & vbCr & strList
	
	MembreDuGroupe = retour

End Function

'Map "P:", "\\sharename\folder"
Function Map(lecteur, chemin)

	Set objNetwork = CreateObject("WScript.Network")
	objNetwork.MapNetworkDrive lecteur, chemin
	
End Function


Function Afficher()

	On Error Resume Next
	
	Dim objNetwork, objDrives, objShell
	Dim strSubst, strSubstVal, strSubstName, strEnumDrive
	
	Set objNetwork = CreateObject("WScript.Network")
	Set objShell = CreateObject("Shell.Application")
	Set objDrives = objNetwork.EnumNetworkDrives
	
	For i = 0 to objDrives.Count - 1 Step 2
		strSubst = objShell.NameSpace(objDrives.Item(i) & Chr(92)).Self.Name 
		strSubstVal = inStr(1,strSubst, Chr(40)) - 2
		strSubstName = Mid(strSubst, 1, strSubstVal)
		strOnlyDrive = strOnlyDrive & objDrives.Item(i)
		strEnumDrive = strEnumDrive & "Drive Letter: " & objDrives.Item(i) & vbCrlF & _
			"Drive Name: " & strSubstName & vbCrlF & "Drive Path: " &  _
			objDrives.Item(i+1) & vbCrLf & vbCrLf
	Next
	'MsgBox strEnumDrive ,, "All mapped Drives"
	
	Afficher = strOnlyDrive
	
End Function


Function Effacer()

	driveList = Split(Afficher(), ":")
	Set objNetwork = CreateObject("WScript.Network")
	For Each objDrive in driveList
		objNetwork.RemoveNetworkDrive objDrive & ":"
	Next
	
End Function


Function MapDefaultPrinter(chemin)

	On Error Resume Next
	Set WshNetwork = CreateObject("WScript.Network")
	Set rc = WshNetwork.AddWindowsPrinterConnection("" & chemin & "") '\\Group\printer
	If Not rc then
	      WScript.Echo("Printer Connection Failed!")
	End If
	WshNetwork.SetDefaultPrinter "" & chemin & ""
	
End Function