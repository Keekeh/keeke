<%
'Verifica se existe uma seção
if session("login") = "" then
'se não existir redireciona o usuário para a página principal
response.Redirect("index.asp")
end if

'Abre a conexão com o banco de dados
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("../database/dados.mdb")

'Lista o conteúdo do banco de dados
Set listar = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens"
listar.open sql, conex, 1,1

'Contamos quantos registros existem no banco
torpedos = listar.Recordcount
%>
<html>
<head>
<title>Administra&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<style type="text/css">
<!--
a:link {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-decoration: none;
}
a:visited {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-decoration: none;
}
a:hover {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-decoration: underline;
}
a:active {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
	text-decoration: none;
}
-->
</style>
<body bgcolor="#999999">
<p align="center"> 
  <script language="JavaScript">
//Função para abrir janelas pop-up
function janela(popupfile,winheight,winwidth)
{
open(popupfile,"PopupWindow","resizable=no,height=" + winheight + ",width=" + winwidth + ",scrollbars=no");
}

//Função para confirmar exclusão
function excluir() 
{
if(window.confirm('Confirma a exclusão?')){
return true;}
else {
return false;
}}
</script>
  <a href="http://www.salamito.fdp.com.br/forum">Ajuda</a> | <a href="sair.asp">Sair 
  da administra&ccedil;&atilde;o</a> <br>
</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="Border-top:#666666 1px solid;Border-left:#666666 1px solid;Border-left:#666666 1px solid;Border-right:#666666 1ps solid;Border-bottom:#666666 1px solid">
  <%while not listar.eof%>
  <tr bgcolor="#CCCCCC"> 
    <td width="50%" height="23" valign="bottom"><strong><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=listar("nome")%></font></strong></td>
    <td width="50%" valign="bottom"> <div align="right"><strong><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=listar("data")%> 
        - <%=listar("hora")%>&nbsp;&nbsp;</font></strong></div></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td height="22" colspan="2"> <table width="99%" height="99%" border="0" align="center" cellpadding="0" cellspacing="5" bgcolor="efefef" style="Border-top:#000000 1px solid;Border-bottom:#000000 1ps solid;">
        <tr> 
          <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;<%=listar("mensagem")%></font></td>
        </tr>
        <tr> 
          <td><div align="right"><a href="mailto:<%=listar("email")%>"><img src="../imagens/email.gif" width="15" height="9" border="0"></a> 
              <a href="#"><img src="../imagens/icq.gif" width="15" height="16" border="0" onClick="javaScript:alert('ICQ: ' + <%=listar("icq")%>)"></a> 
              <a href="<%=listar("site")%>" target="_blank"><img src="../imagens/site.gif" width="19" height="17" border="0"></a>&nbsp;</div></td>
        </tr>
        <tr>
          <td><a href="remover.asp?id=<%=listar("id")%>" onClick="return excluir()">Remover</a> | <a href="javascript:janela('atualizar.asp?id=<%=listar("id")%>',400,500)">Atualizar</a></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="2"><div align="right"> &nbsp;</div></td>
  </tr>
  <%listar.movenext
wend
%>
</table>
<div align="center"><br>
  <font color="#CCCCCC" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estamos 
  com 
  <%response.Write(torpedos)
  if torpedos = 1 then response.Write(" torpedo.") else response.Write(" torpedos.") end if%>
  </font></div>
<br>

</body>
</html>
