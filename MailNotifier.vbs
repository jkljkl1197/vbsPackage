' vous devez vous connecter sur votre navigateur et autoriser les application moin securiser ou le script ne marchera pas et error sera "Not Availble"
' https://myaccount.google.com/lesssecureapps?rfn=27&rfnc=1&eid=-8756118372491187715&et=0&asae=2&pli=1

' Mettre votre mail GMAIL
Const fromEmail	= "" ' monemail@gmail.com
Const password	= ""

Function SendMail(fromEmail, password)
	' Object creation   
	Set objMsg = CreateObject("CDO.Message")   
	Set msgConf = CreateObject("CDO.Configuration")   
	  
	' Server Configuration   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = fromEmail   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password   
	msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1   
	msgConf.Fields.Update   
	  
	' Email   
	objMsg.To = fromEmail
	objMsg.From = fromEmail   
	objMsg.Subject = "Test send with Gmail account"   
	objMsg.HTMLBody = "HTML/Plain text message."
	objMsg.Sender = "Mr. Name ggaajj"
	  
	Set objMsg.Configuration = msgConf
	  
	' Send   
	objMsg.Send   
	  
	' Clear   
	Set objMsg = nothing  
	Set msgConf = nothing
End Function

SendMail fromEmail, password