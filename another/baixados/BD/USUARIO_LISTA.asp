<!--#Include File= "conexao3h.asp"-->

<HTML>
<HEAD>
	<TITLE>Teste</TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
</HEAD>

<BODY>
<table width="80%" align="center">
  <tr> 
    <td bgcolor="#006666" align="center"><strong><i> <font face="Monotype Corsiva" size="4">Nome</font></i></strong></td>
    <td bgcolor="#006666" align="center"><strong><i> <font face="Monotype Corsiva" size="4">E-mail</font></i></strong></td>
    <td width="20%" bgcolor="#006666" align="center"><strong><i> <font face="Monotype Corsiva" size="4">Telefone</font></i></strong></td>
    <td bgcolor="#006666" align="center">&nbsp;</td>
    <td bgcolor="#006666" align="center">&nbsp;</td>
    <td bgcolor="#006666" align="center">&nbsp;</td>
  </tr>
  <tr align="center" bgcolor="#dddddd"> 
    <%
	Set rstConsulta = Server.CreateObject("ADODB.RecordSet")
	rstConsulta.Open "SELECT * FROM Usuario", conTurma3h	
	cor = "#eeeeee"
	while not rstConsulta.EOF 
		if cor = "#eeeeee" then
			cor = "#ffffff"
		else
			cor = "#eeeeee"
		end if
	
		%>
  <tr bgcolor="<%= cor %>"> 
    <td bgcolor="#FFFFFF"> 
      <% if isNull(rstConsulta("Usr_Nome")) or rstConsulta("Usr_Nome") = "" then
						Response.write "------"
					else
						Response.write rstConsulta("Usr_Nome")
					end if %>
    <td bgcolor="#FFFFFF"> 
      <% if isNull(rstConsulta("Usr_Email")) or rstConsulta("Usr_Email") = "" then
						Response.write "------"
					else
						Response.write  rstConsulta("Usr_Email")
					end if %>
    <td bgcolor="#FFFFFF"><%= rstConsulta("Usr_Telefone") %> 
    <td align=center bgcolor="#FFFFFF"><a href="Usuario_edt.asp?codigo=<%=rstConsulta("Usr_Codigo")%>">Editar</a></td>
    <td align=center bgcolor="#FFFFFF"><a href="Usuario_Exclui.asp?codigo=<%=rstConsulta("Usr_Codigo")%>">Apagar</a></td>
    <td align=center bgcolor="#FFFFFF"><a href="frmCadUsuario.html">Novo</a></td>
    <% rstConsulta.MoveNext
	wend

Set rstConsulta = nothing
Set conTurma3h = nothing
%>
</table>
</BODY>
</HTML>