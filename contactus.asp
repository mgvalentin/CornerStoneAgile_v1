<%
'-----EDIT THE MAILING DETAILS IN THIS SECTION-----
dim fromName, fromAddress, recipientName, recipientAddress, subject, body, sentTo, varEmail, varName, varMsg

varEmail = Request.Form("youremail")
varName = Request.Form("yourname")
varMsg = Request.Form("yourmsg")

fromName = varName
fromAddress = varEmail
recipientName = "CornerStone Agile"
recipientAddress= "info@cornerstoneagile.com"
subject = "CornerStone Agile Website Inquiry"
body = varMsg

'-----YOU DO NOT NEED TO EDIT BELOW THIS LINE-----

sentTo = "NOBODY"
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName = fromName
Mailer.FromAddress = fromAddress
Mailer.RemoteHost = "mrelay.perfora.net"
if Mailer.AddRecipient (recipientName, recipientAddress) then
sentTo=recipientName & " (" & recipientAddress & ")"
end if
Mailer.Subject = subject
Mailer.BodyText = body

if Mailer.SendMail then
response.redirect("contactus_thankyou.html")
else
response.redirect("index.html")
end if


'response.redirect("contactus_thankyou.html")

'response.redirect("/workshop_reg_payment.asp")
'server.transfer("/workshop_reg_payment.asp")

'javascript:history.back()

%>

