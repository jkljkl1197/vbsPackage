Set MonShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Dim str, output, i2, newAsc, tmp
const CharWhitelist = "BCDFGHJKMPQRTVWXY2346789"

' on recupere la clé de registre Brute
CleBrute = MonShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
' on utilise la fonction de traduction de la clée Brute
WinCle = ExtractKey(CleBrute)

' on change le 2 derniere lettre
For i2=28 To Len(WinCle)
	' genrenation d'un random
	Dim max,min
	max=24
	min=1
	Randomize
	tmp = Int((max-min+1)*Rnd+min)
	
	' gerere des caractere aléatoire depuis un index
    newAsc = AscW(Mid(CharWhitelist,tmp,1))-1
    If newAsc < 0 Then 
        newAsc = 65535
    End If
    output = output & ChrW(newAsc)
Next

' on lance l'interface avec comme message la clée réelle avec en plus 2 caractere de moin et 2 caractere generer par le random
MsgBox Left(WinCle, 27) & output, 0, "Clee Windows"

' écrie le fichier avec la clée
outFile="P:\prog\Tp-03\key.txt"
Set fso = fso.CreateTextFile(outFile,True)
fso.Write WinCle & vbCrLf
fso.Close

Function ExtractKey(KeyInput)
	' base de la clée pour la dechiffrer
    Const KeyOffset = 52
	' i est composer du nombre de caractere possible "CharWhitelist" de la cle et de 4 trait "-" sois 28
    i = 28
	' liste des caractere qu'une clée est composer
    Do
		' variable temporaire de chaque itération
        Cur = 0
        x = 14
        Do
            Cur = Cur * 256
            Cur = KeyInput(x + KeyOffset) + Cur
            KeyInput(x + KeyOffset) = (Cur \ 24) And 255
            Cur = Cur Mod 24
            x = x -1
        Loop While x >= 0
        i = i -1
		' MID() Recupere un caractere d'un index dans la string CharWhitelist
		' KeyOutput s'additionne lui meme en tring pour former la cle complete
        KeyOutput = Mid(CharWhitelist, Cur + 1, 1) & KeyOutput
        If (((29 - i) Mod 6) = 0) And (i <> -1) Then
            i = i -1
            KeyOutput = "-" & KeyOutput
        End If
    Loop While i >= 0
	' "Return" permet de returner une string de la clée generer par la fucntion
    ExtractKey = KeyOutput
End Function