' ce secript envoie via la rest api de pushbullet une notification munie d'un message
' se connecter avec son compte gmail a https://www.pushbullet.com/
' acceder a https://www.pushbullet.com/#settings/account et creer un acces token a mettre comme ApiKey
Dim objXmlHttpMain , URL, ApiKey, message
ApiKey = "o.b9JTi4olNxv2VtH3bbfXJ15JjXBlW0No"

' message est Lorem Ipsum Lorem Ipsum Lorem Ipsum.
strJSONToSend = "{""type"": ""note"", ""title"": ""Alert"", ""body"": ""Lorem Ipsum Lorem Ipsum Lorem Ipsum.""}"

URL="https://api.pushbullet.com/v2/pushes" 
Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP") 
on error resume next 
objXmlHttpMain.open "POST",URL, False 
objXmlHttpMain.setRequestHeader "Authorization", "Bearer " & ApiKey
objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
objXmlHttpMain.send strJSONToSend

set objJSONDoc = nothing 
set objResult = nothing