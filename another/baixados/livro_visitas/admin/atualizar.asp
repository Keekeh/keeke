<!--#include file = "config.asp"-->
<%
'Verifica se existe uma seção
if session("login") = "" then
'se não existir redireciona o usuário para a página principal
response.Redirect("index.asp")
end if

'Pegamos o id da mensagem
id = request("id")

'Abre a conexão com o banco de dados
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("../database/dados.mdb")

'Procuramos a mensagem com o respectivo id para listarmos seu conteúdo nos campos
Set ver = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens where id="&id
ver.open sql, conex, 1,1

%>
<html>
<head>
<title><%=titulo_site%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
a:link {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #CCCCCC;
	text-decoration: none;
}
a:visited {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #CCCCCC;
	text-decoration: none;
}
a:hover {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #CCCCCC;
	text-decoration: underline;
}
a:active {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #CCCCCC;
	text-decoration: none;
}

input {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	background-color: #cccccc;
	border: 1px solid #333333;
}
textarea {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	background-color: #CCCCCC;
	border: 1px solid #333333;
}
-->
</style>
</head>

<body bgcolor="#999999">
<form name="form1" method="post" action="salvar.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" style="Border-top:#000000 1px solid;Border-left:#000000 1px solid;Border-left:#000000 1px solid;Border-right:#000000 1ps solid;Border-bottom:#000000 1px solid"">
    <tr bgcolor="#666666"> 
      <td height="26" colspan="2"> <div align="center">
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="none">
            <tr> 
              <td width="88%"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nova 
                  mensagem</strong></font></div></td>
              <td width="12%"><div align="center"><a href="javaScript:window.close();">Fechar</a></div></td>
            </tr>
          </table>
          <font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
    </tr>
    <tr bgcolor="efefef">
      <td>&nbsp;</td>
      <td><div align="right"><a href="javaScript:window.close();">&nbsp;</a></div></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td width="28%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td width="72%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Nome:</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="nome" type="text" id="nome" value="<%=ver("nome")%>" size="40">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;E-mail:</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="email" type="text" id="email" value="<%=ver("email")%>" size="40">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;ICQ:</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="icq" type="text" id="icq" value="<%=ver("icq")%>" size="40">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Pa&iacute;s:</font></td>
      <td> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="pais" type="text" id="pais" value="<%=ver("pais")%>">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;P&aacute;gina 
        pessoal:</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="site" type="text" id="site" value="<%=ver("site")%>" size="40">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Mensagem:</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <textarea name="mensagem" cols="40" rows="10" id="mensagem"><%=ver("mensagem")%></textarea>
        <input name="id" type="hidden" id="id" value="<%=id%>">
        </font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td colspan="2"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="submit" name="Submit" value="Atualizar mensagem">
          </font></div></td>
    </tr>
    <tr bgcolor="efefef"> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>