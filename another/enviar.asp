<%
sch="http://schemas.microsoft.com/cdo/configuration/"
Set cdoConfig=Server.CreateObject("CDO.Configuration")


'Variaveis
Dim meuservidorsmtp
Dim minhacontaautenticada
Dim minhasenhaparaenvio
Dim emailorigem
Dim emaildestino

'Abaixo seguem algumas definicoes de variaveis para o envio de seu formulario. Por favor preencha os campos abaixo.

meuservidorsmtp = "localhost" ' Informacoes so seu servidor SMTP
minhacontaautenticada = "kikeh@kikeh.com" ' conta de e-mail utilizada para enviar
minhasenhaparaenvio = "PittBull19" ' senha da conta de e-mail
emailorigem = "heykikeh@gmail.com" ' e-mail que indica de onde partiu a mensagem
emaildestino = "ki.keh@live.com" ' e-mail que vai receber as mensagens do formulario

'Fim da definição manual de parâmetros.

cdoConfig.Fields.Item(sch & "sendusing") = 2
cdoConfig.Fields.Item(sch & "smtpauthenticate") = 1
cdoConfig.Fields.Item(sch & "smtpserver") = meuservidorsmtp
cdoConfig.Fields.Item(sch & "smtpserverport") = 25
cdoConfig.Fields.Item(sch & "smtpconnectiontimeout") = 30
cdoConfig.Fields.Item(sch & "sendusername") = minhacontaautenticada
cdoConfig.Fields.Item(sch & "sendpassword") = minhasenhaparaenvio
cdoConfig.fields.update
Set cdoMessage = Server.CreateObject("CDO.Message")
Set cdoMessage.Configuration = cdoConfig

cdoMessage.BodyPart.Charset = "utf-8"
cdoMessage.From = emailorigem
cdoMessage.To = emaildestino
cdoMessage.Subject = "Formulario de Contato"
cdoMessage.ReplyTo = Request("emailrem")

strBody = "Dados <br> <br>" & _
"Nome:"& Request("nomerem")& "<br>" & _
"E-Mail:"& Request("emailrem")& "<br>" & _
"Assunto:"& Request("assunto")& "<br>" & _
"Mensagem:"& Request("recado")

strBody = strBody & "."
cdoMessage.HTMLBody = strBody
cdoMessage.Send

Set cdoMessage = Nothing
Set cdoConfig = Nothing

response.write "O e-mail foi processado e enviado com sucesso"
%>