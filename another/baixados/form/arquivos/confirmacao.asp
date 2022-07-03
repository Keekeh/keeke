<!-- #include file="conexaobd.asp" -->
<!-- #include file="ADOVBS.inc" -->
<%

dim conconexao

'Aqui você pode inserir mais dados, siga o exemplo abaixo
Session("nome") = Request.Form("nome")
Session("endereco") = Request.Form("endereco")
Session("tel") = Request.Form("tel")
Session("email") = Request.Form("email")

Sub ProcessaPagina

Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "form", conconexao, adOpenDynamic, adlockoptimistic, adcmdtable

rs.Addnew

'Aqui você pode inserir mais dados, siga o exemplo abaixo
rs.fields("nome") = Session("nome")
rs.fields("endereco") = Session("endereco")
rs.fields("tel") = Session("tel")
rs.fields("email") = Session("email")

rs.update
end sub

processapagina
%>

<html>
<title> Form </title>
<body>
<h4>Dados enviados com sucesso!!!</h4>
</body>
</html>