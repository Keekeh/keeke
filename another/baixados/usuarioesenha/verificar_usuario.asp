<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
%>
<!--#include file="conexao.asp"-->
<%
Dim usuario, senha

usuario = Trim(LCase(Request.Form("usuario")))
senha = Trim(LCase(Request.Form("senha")))

if len(usuario) = 0 then
response.redirect "default.asp?erro=1"
end if

if len(senha) = 0 then
response.redirect "default.asp?erro=2"
end if

Call AbrirDB

sql = "SELECT * FROM login WHERE nome='" & usuario & "'"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open sql,ConnDB,3,3

if not RS.EOF then

if RS("senha") <> senha then
response.redirect "default.asp?erro=4"
else
Session("login") = True
response.redirect "logado.asp"
end if

else

response.redirect "default.asp?erro=3"

end if

RS.close
Set RS = Nothing
Call FecharDB

'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
%>