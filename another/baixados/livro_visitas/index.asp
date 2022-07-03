<!--#include file = "admin/config.asp"-->
<%
'Abre a conex�o com o banco de dados 
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("database/dados.mdb")

'Lista o conte�do do banco de dados
Set listar = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens"
listar.open sql, conex, 1,1

'Contamos quantos registros existem no banco
torpedos = listar.Recordcount
%>

<html>
<head>
<title><%=titulo_site%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body id="index">
<script language="JavaScript">
function janela(popupfile,winheight,winwidth)
{
open(popupfile,"PopupWindow","resizable=no,height=" + winheight + ",width=" + winwidth + ",scrollbars=no");
}
</script>
<div><a href="javascript:janela('nova_mensagem.asp',400,500)" onClick="">Nova 
  mensagem</a><br>
</div>
<br>
<%if listar.eof then%>
<div>
  <br>
  N&atilde;o 
  h&aacute; mensagens!</div>
<br>
<%else%>
<table width="100%"  cellspacing="0" cellpadding="0" style="Border-top:#666666 1px solid;Border-left:#666666 1px solid;Border-left:#666666 1px solid;Border-right:#666666 1ps solid;Border-bottom:#666666 1px solid">
  <%while not listar.eof%>
  <tr bgcolor="#CCCCCC"> 
    <td width="50%" height="23" valign="bottom"><strong><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=listar("nome")%></font></strong></td>
    <td width="50%" valign="bottom"> 
      <div align="right"><strong><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=listar("data")%> 
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
          <td><div align="right"><a href="mailto:<%=listar("email")%>"><img src="imagens/email.gif" width="15" height="9" border="0"></a> 
              <a href="#"><img src="imagens/icq.gif" width="15" height="16" border="0" onClick="javaScript:alert('ICQ: ' + <%=listar("icq")%>)"></a> 
              <a href="<%=listar("site")%>" target="_blank"><img src="imagens/site.gif" width="19" height="17" border="0"></a>&nbsp;</div></td>
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

<div><br>
 Estamos com <%response.Write(torpedos)
  if torpedos = 1 then response.Write(" torpedo.") else response.Write(" torpedos.") end if%> </div> <br>
  <%end if%>
  <br>
  <br>
  <br>
  <div>
  
  --------------------------------------------------------- <br>
  Developed by D&iacute;ogenes G&ouml;tz<br>
  diogotz@bol.com.br</div> 
</body>
</html>
