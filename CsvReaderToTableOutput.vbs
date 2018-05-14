
	' Definition des variables
	Dim fso, outPutFile, inPutFile, arrayLenght, tmpNumber, tmpStr, lenght
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set outPutFile = fso.CreateTextFile("output.txt", True)
	Set inPutFile = fso.OpenTextFile("list.csv")

	tmpNumber = 0
	lenght = 16
	tmpStr = ""

	' Boucle sur le nombre de ligne.
	Do while NOT inPutFile.AtEndOfStream
		' Crée un tableau de chaque donnée separer par des ;
		arrStr = split(inPutFile.ReadLine,";")
		' Obtien la dimention du tableau d'une seul ligne
		arrayLenght = uBound(arrStr) + 1
		' Boucle sur le nombre de colum; par decroissance
		For i=arrayLenght to 0 Step -1
			' Execute si i est la longeur maximal de l'array
			If i = arrayLenght Then
				' String utiliser dans les deux cas par l'affichage; prend la var longeur et soustrait la longeur de la string d'affichage et transforme ce nombre en espace. 
				tmpStr = arrStr(0) & Space(lenght - Len(arrStr(0))) & arrStr(1) & Space(lenght - Len(arrStr(1))) & arrStr(2) & Space(lenght - Len(arrStr(2))) & arrStr(3) & Space(lenght - Len(arrStr(3))) & arrStr(4) & Space(lenght - Len(arrStr(4))) & arrStr(5)
					' Execute si le while est a la premiere ligne du csv
					If tmpNumber = 0 Then
						' Affiche la ligne des titres dans le terminale et ecrie dans le fichier avec la meme syntaxe
						outPutFile.WriteLine(tmpStr)
						WScript.Echo(tmpStr)
						tmpStr = "_____________________________________________________________________________________________"
						outPutFile.WriteLine(tmpStr)
						Wscript.Echo(tmpStr)
					Else
						' Affiche la ligne d'info dans le terminale et ecrie dans le fichier avec la meme syntaxe
						outPutFile.WriteLine(tmpStr)
						WScript.Echo(tmpStr)
					End If
			End If
		Next
		
		' Compteur representant notre position de ligne du csv
		tmpNumber = tmpNumber + 1
	Loop

	' Ferme tout les fichier et vide une portion de la memoir utiliser
	inPutFile.Close
	set inPutFile = Nothing
	outPutFile.Close
	set outPutFile = Nothing
	set fso = Nothing
	
	WScript.Quit