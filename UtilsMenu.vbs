Set obj = CreateObject("WScript.Network")
Set wshShell = CreateObject( "WScript.Shell" )
Set objWMIService = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

Function Fib_Func(n)
  If (n < 2) Then
    Fib_Func = n
  Else
    Fib_Func = Fib_Func(n-1) + Fib_Func(n-2)
  End If
End Function

while(true)
	OptionSelected = inputbox("choisissez votre option: " + VbCrLf + " 0:Quitter; " + VbCrLf + " 1:Info; " + VbCrLf + " 2:ListLecteur; " + VbCrLf + " 3:VariableENV; " + VbCrLf + " 4:SuiteDuNombre; " + VbCrLf + " 5:InfoIp; ", "Menu principale")
	Select Case OptionSelected
		Case 0
			' MsgBox "Quitter"
			WScript.Quit
		Case 1
			user = obj.UserName
			pcname = obj.ComputerName
			domain = obj.UserDomain
			MsgBox "Nom du script: " & Wscript.ScriptName & "; Utilisateur: " & user & "; Nom PC: " & pcname & "; Groupe ou Domaine: " & domain
		Case 2
			Set colDrives = obj.EnumNetworkDrives
			Set colPrinter = obj.EnumPrinterConnections
			i = 0
			For Each objDItem in colDrives
				TempDStr = TempDStr & objDItem & "; "
				i = i + 1
			NEXT
			i = 0
			For Each objPItem in colPrinter
				TempPStr = TempPStr & objPItem & "; "
				i = i + 1
			NEXT
			MsgBox TempDStr & " + VbCrLf + " & TempPStr
		Case 3
			MsgBox wshShell.ExpandEnvironmentStrings( "PATH=%PATH%" )
		Case 4
			Fnumber = inputbox("choisissez votre Nombre!", "Fibonacci Suite")
			MsgBox "Le nombre correspondant selon Fibonacci est : " & Fib_Func(Fnumber)
		Case 5
			For Each Card in objWMIService
				if Card.IPEnabled then
					StrIP = Card.IPAddress(i)
					CompName= obj.Computername
					TempStr = "IP Address:  " & StrIP & vbNewLine
					TempCStr = TempCStr & TempStr & "; "
				end if
			next
			MsgBox TempCStr
		Case Else
			MsgBox "Entrer un nombre valide."
	End Select
Wend

wscript.echo "END"