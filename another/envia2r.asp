<%
' change to address of your own SMTP server
strHost = "smtp-mail.outlook.com"
If Request("Send") <> "" Then
   Set Mail = Server.CreateObject("Persits.MailSender")
   ' enter valid SMTP host
   Mail.Host = strHost

Mail.From = "ki.ke@live.com" ' manter igual ao mail.username
Mail.FromName = Request("FromName") ' opcional
Mail.AddAddress "ki.keh@live.com", "kikeh" 'O e-mail que receberá o resultado do form.
Mail.Username = "ki.keh@live.com" 'Conta válida do servidor de mail para fazer autenticação.
Mail.Password = "Iloveu4ever" ' Informe a senha da conta de e-mail acima especificada.
Mail.Port = 587

   ' message subject
   Mail.Subject = Request("Subject")
   ' message body
   Mail.Body = Request("Body")
   strErr = ""
   bSuccess = False
   On Error Resume Next ' catch errors
   Mail.Send ' send message
   If Err <> 0 Then ' error occurred
      strErr = Err.Description
   else
      bSuccess = True
   End If
End If
%>
<HTML>
<BODY BGCOLOR="#FFFFFF">
<% If strErr <> "" Then %>
<h3>Error occurred: <% = strErr %>
<% End If %>
<% If bSuccess Then %>
Success! Message sent.
<% End If %>