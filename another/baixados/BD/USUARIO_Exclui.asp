<%Option Explicit%>
<!--#Include File="conexao3h.asp"-->

<HTML>
<HEAD>
	<TITLE>Excluir usuario</TITLE>
</HEAD>

<BODY>
<% Dim rstExclui, vCodigo
   vCodigo=Request.QueryString("Codigo")
   Set rstExclui=conTurma3h.Execute("Delete FROM Usuario where Usr_Codigo=" & vCodigo &"")
   'Set rstExclui=conTurma3h.Execute("Delete FROM Contato where Usr_Codigo=" & vCodigo &"")
   Response.Redirect "USUARIO_LISTA.asp"
%>

</BODY>
</HTML>