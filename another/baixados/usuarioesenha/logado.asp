<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
if not Session("login") then response.redirect "default.asp?erro=5" end if
%>
<!--#include file="conexao.asp"-->
<%
Dim status
Dim strMsg
strMsg = cstr(Request.QueryString("action"))

Select Case strMsg

Case "add"

Dim usuario, senha

usuario = Trim(LCase(Request.Form("usuario")))
senha = Trim(LCase(Request.Form("senha")))

if len(usuario) = 0 then
response.redirect "logado.asp?action=add1"
end if

if len(senha) = 0 then
response.redirect "logado.asp?action=add1"
end if

sql = "INSERT INTO login (nome,senha) VALUES ('" & usuario & "','" & senha & "')"
Call AbrirDB
ConnDB.Execute(sql)
Call FecharDB
%>
<center><font color="red">Adicionado com sucesso.</font></center>

<%
Case "add1"
status = 1
Sub Dados() 
%>
<form action="?action=add" method="POST">
<center>
<table border><tr><td>Nome do usuário:</td><td align="right"><input type="text" name="usuario"></td></tr>
<tr><td>Senha:</td><td align="right"><input type="password" name="senha"></td></tr>
<tr><td align="center" colspan="2"><input type="submit" value="Adicionar"></td></tr></table>
</center>
</form>
<% 
end sub

Case "del"

Dim user, pass, deletar
user = cstr(Request.QueryString("usuario"))
pass = cstr(Request.QueryString("senha"))
if len(user) = 0 then response.redirect "logado.asp" end if
if len(pass) = 0 then response.redirect "logado.asp" end if
deletar = "DELETE * FROM login WHERE nome='" & user & "' and senha='" & pass & "'"
Call AbrirDB
ConnDB.Execute(deletar)
Call FecharDB
%>
<center><font color="red">Deletado com sucesso.</font></center>
<%
Case "del1"
Dim i, RS
sql = "SELECT * FROM login"
Set RS = Server.CreateObject("ADODB.RecordSet")
Call AbrirDB
RS.open sql,ConnDB,3,3
status = 2
Sub Dados2()
%>
<center><table border><tr><td align="center"><b>Nome</b></td><td align="center"><b>Senha</b></td></tr>
<% Dim name, password %>
<% For i = 1 to RS.RecordCount %>
<%
name = RS("nome")
password = RS("senha")
%>
<tr><td><a href=?action=del&usuario=<%=name %>&senha=<%=password%>><%=name%></a></td>
<td><a href=?action=del&usuario=<%=name %>&senha=<%=password%>><%=String(len(password),"*")%></td></tr>
<% 
RS.MoveNext
if RS.eof then Exit For
Next

%>
</table>
<%
RS.close
set RS = nothing
Call FecharDB
end sub

Case "logout"
Session("login") = False
response.redirect "default.asp"

end select
%>


<html>
<body>
Você agora esta logado.<br>
Se quiser adicionar um usuario <a href="?action=add1">clique aqui</a>.<br>
Se quiser vizualizar os usuarios <a href="?action=del1">clique aqui</a>. Clique em cima do nome ou senha para deletar<br>
Se quiser fazer o LogOut <a href="?action=logout">clique aqui</a>.<br>
<% if status = 1 then Call Dados end if%>
<% if status = 2 then Call Dados2 end if%>

</body>
</html>
<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
%>