<!--#include file = "admin/config.asp"-->
<%
'Abre a conexão com o banco de dados
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("database/dados.mdb")

'Pega os dados digitados no forluário
nome = request("nome")
email = request("email")
icq = request("icq")
pais = request("pais")
site = request("site")
mensagem = request("mensagem")

'verificamos os campos do formulário
if nome = "" then
alerta = alerta + "<br>- Nome"
erro = true
end if

if email = "" then
alerta = alerta + "<br>- E-mail"
erro = true
end if

if icq = "" then
alerta = alerta + "<br>- ICQ"
erro = true
end if

if pais = "" then
alerta = alerta + "<br>- País"
erro = true
end if

if site = "" then
alerta = alerta + "<br>- Página pessoal"
erro = true
end if

if mensagem = "" then
alerta = alerta + "<br>- Mensagem"
erro = true
end if

'Se encontrar erros exibe uma mensagem com o nome dos campos
if erro = true then
response.Write("<b>Os seguintes erros foram encontrados:<br></b>")
response.Write(alerta)
response.Write("<br><br><a href='' onclick='javaScrip:window.history.go(-1)'>Voltar</a>")

'caso contrário
else

'Salva os dados
Set salva = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens"
salva.open sql, conex, 3,3
salva.addnew
salva("nome")= nome
salva("email") = email
salva("icq") = icq
salva("pais") = pais
salva("site") = site
salva("data") = date
salva("hora") = time
salva("mensagem") = mensagem
salva.update

'Exibe a mensagem de que a mensagem foi enviada corretamente
response.Write("<html><head><title>"&titulo_site&"</title><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'></head><body bgcolor='#999999'><style type='text/css'><!--a:link {	font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;	color: #CCCCCC;	text-decoration: none;}a:visited {	font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;	color: #CCCCCC;	text-decoration: none;}a:hover {	font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;	color: #CCCCCC;	text-decoration: underline;}a:active {	font-family: Verdana, Arial, Helvetica, sans-serif;	font-size: 11px;	color: #CCCCCC;	text-decoration: none;}--></style><table width='100%' height='100%'><tr><td><table align='center'><tr><td><div align='center'><font color='#FFFFFF' size='1' face='Verdana, Arial, Helvetica, sans-serif'>A mensagem foi enviada corretamente.<br><br><a href='#' onClick='javaScript:window.close()'>Fechar </a></font><br></div></td></tr></table></td></tr></table></body></html>")
end if
%>