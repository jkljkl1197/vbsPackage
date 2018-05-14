Set objNet = WScript.CreateObject("WScript.Network")
' CSV variables
Dim AccountName
Dim newUserFields
Dim objNewUser

' path du CSV
const varFileName = "P:\prog\Tp-03\users.cvs"
 
' ouvre le fichier CSV sans le modifier et li la premiere ligne qui est les columns.
Dim objFSO
Dim objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(varFileName, 1)
' skip la premiere ligne
objFile.ReadLine
 
 ' en cas d error ont continue au prochain.
 'on error resume next
 
' li le reste du fichier et crée les utilisateur a l'aide des informations lue.
Do Until objFile.AtEndOfStream
		' met dans un tableau les information lue.
		newUserFields = Split(objFile.ReadLine,",")
		AccountName = newUserFields(0)
	' créer les utilisateurs en utilisant la ligne lue precédament.
	AddUser AccountName, "Administrateurs"
Loop

Function AddUser(strUser, sGroup)
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
 
MsgBox("Les utilisateurs créer ont été créer!")
 
WScript.Quit  