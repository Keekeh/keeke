<html>

<head>
<title>Logando - se</title>
<STYLE>
BODY{FONT-FAMILY:VERDANA;FONT-SIZE:11;COLOR:#000000}
</STYLE>
</head>

<body>

<%
tipo = request.querystring("tipo")
if IsEmpty(tipo) or tipo = "" then
%>
<div align="center">
  <center>
  <form action="logar.asp?tipo=logar" method="post">
  <table border="0" width="300" style="font-family: verdana; color: #000000; font-size: 12">
    <TR>
      <TD colspan="2" style="background-color:#c0c0c0;padding:3" valign="middle" align="center">
      <b>&Aacute;rea administrativa</b>
      </TD>
    </TR>
    <tr>
      <td width="52" valign="middle" align="right" style="background-color: #f1f1f1; padding: 3">login:</td>
      <td width="232" style="background-color: #f1f1f1; padding: 3" valign="middle" align="left"><input type="text" name="login" value="livro" style="font-family: verdana; color: #ff0000; font-size: 12;border:1px solid #000000;background-color:#f1f1f1" size="20"></td>
    </tr>
    <tr>
      <td width="52" valign="middle" align="right" style="background-color: #f1f1f1; padding: 3">Senha:</td>
      <td width="232" style="background-color: #f1f1f1; padding: 3" valign="middle" align="left"><input type="password" name="senha" value="123456" style="font-family: verdana; color: #ff0000; font-size: 12;border:1px solid #000000;background-color:#f1f1f1" size="20"></td>
    </tr>
    <TR>
      <TD colspan="2" style="background-color:#f1f1f1;padding:3" valign="middle" align="center">
      <input type="submit" value="Logar-se" style="border:1px solid #000000;font-family:verdana;font-size:11;font-weight:bold;background-color:#ffff00">
      </TD>
    </TR>
  </table>
  </form>
  </center>
</div>
<%
end if
if tipo = "logar" then
login = request.form("login")
senha = request.form("senha")

If login = "livro" and senha = "123456" then
session("livro") = "logado"
response.redirect "administracao.asp"
else
response.write "<CENTER><br><br><br><br><SPAN STYLE=""BACKGROUND-COLOR:#F1F1F1;BORDER:1PX SOLID #000000;PADDING:3;WIDTH:250;TEXT-ALIGN:CENTER"">ALGO DE ERRADO OCORREU, OU VOC&Ecirc; N&Atilde;O TEM ACESSO A ESTA &Aacute;REA!</SPAN></CENTER>"
end if
end if
%>

</body>

</html>
