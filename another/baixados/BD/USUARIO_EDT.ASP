
<!--#Include File="conexao3h.asp"-->

<HTML>
<HEAD>
	<TITLE>Usuario</TITLE>
	<style>
		body {font-family: arial}
		h1 {color: #a0a0a0; font-size: 18px}
		td {font-size: 10px}
	</style>
</HEAD>
<BODY>
<%
  Dim rstConsulta,SQL
if Request.QueryString("codigo") <> "" then
	Sql = "SELECT TOP 1 * From Usuario Where Usr_Codigo="& Request.QueryString("codigo")
else
	Sql = "SELECT TOP 1 * From Usuario Order by Usr_Codigo Desc"
end if

  Set rstConsulta = Server.CreateObject("ADODB.RecordSet")
  rstConsulta.open SQL, conTurma3h

If not rstConsulta.EOF then
%>
<h1 align=center>Confirme os Dados Cadastrados</h1>
<form method="post" action="Usuario_Altera.asp">
	<input type="hidden" name="Usr_Codigo" value="<%=rstConsulta("Usr_Codigo")%>">
	
  <table>
    <tr> 
      <td width="83"> Nome: 
      <td width="425"> <input name="Usr_Nome" type="text" value="<%=rstConsulta("Usr_Nome")%>" size="70"> 
        <br> 
      <td width="138">Apelido: 
      <td width="289">
      <input type="text" name="Usr_Apelido" value="<%=rstConsulta("Usr_Apelido")%>" size="20" size="20"> 
    <tr> 
      <td> Sexo: 
      <td> <input type="radio" name="Usr_Sexo" value="M" <% if rstConsulta("Usr_Sexo") = "M" then Response.write "checked" end if %>>
        Masculino&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="Usr_Sexo" value="F" <% if rstConsulta("Usr_Sexo") = "F" then Response.write "checked" end if %>>
        Feminino<br> 
      <td>&nbsp; 
      <td>&nbsp; 
    <tr> 
      <td> E-mail: 
      <td> <input name="Usr_Email" type="text" value="<%=rstConsulta("Usr_Email")%>" size="70"> <br> 
      <td>&nbsp; 
      <td>&nbsp; 
    <tr> 
      <td> Senha: 
      <td> 
      <input name="Usr_Senha" type="password" value="<%=rstConsulta("Usr_Senha")%>" size="20"> <br> 
      <td>Confirma Senha: 
      <td>
      <input name="confirma" type="password" value="<%=rstConsulta("Usr_Senha")%>" size="20"> 
    <tr> 
      <td> Rua/Av.: 
      <td> <input name="Usr_Rua" type="text" value="<%=rstConsulta("Usr_Rua")%>" size="70"> <br> 
      <td>Numero: 
      <td>
      <input name="Usr_Numero" type="text" value="<%=rstConsulta("Usr_Numero")%>" size="20"> 
    <tr> 
      <td> Complemento: 
      <td> 
      <input name="Usr_Complemento" type="text" value="<%=rstConsulta("Usr_Complemento")%>" size="20"> <br> 
      <td>Bairro: 
      <td>
      <input name="Usr_Bairro" type="text" value="<%=rstConsulta("Usr_Bairro")%>" size="20"> 
    <tr> 
      <td> Cidade: 
      <td> <input name="Usr_Cidade" type="text" value="<%=rstConsulta("Usr_Cidade")%>" size="50"> <br> 
      <td>Estado: 
      <td><input name="Usr_UF" type="text" value="<%=rstConsulta("Usr_UF")%>" size="5"> 
    <tr> 
      <td> Cep: 
      <td> 
      <input name="Usr_Cep" type="text" value="<%=rstConsulta("Usr_Cep")%>" size="20"> <br> 
      <td>&nbsp; 
      <td>&nbsp; 
    <tr> 
      <td> Telefone: 
      <td> 
      <input name="Usr_Telefone" type="text" value="<%=rstConsulta("Usr_Telefone")%>" size="20"> <br> 
      <td>Fax: 
      <td>
      <input name="Usr_Fax" type="text" value="<%=rstConsulta("Usr_Fax")%>" size="20"> 
    <tr> 
      <td> Celular: 
      <td> 
      <input name="Usr_Celular" type="text" value="<%=rstConsulta("Usr_Celular")%>" size="20"> <br> 
      <td>&nbsp; 
      <td>&nbsp; 
  </table>	
	<center><input type="submit" value="Alterar">
			<input type="button" value="Novo" onclick="document.location='frmCadUsuario.html'">
			<input type="button" value="Listagem" onclick="document.location='Usuario_Lista.asp'">
	</center>
</form>
<%else%>
	<h1> Erro <h1>
<%end if %> </h1>
</BODY>
</HTML>