<%if Session("livro") <> "logado" then
Response.Redirect "logar.asp"
else%>
<html>

<head>
<title>Deletar ou Atualizar MENSAGENS</title>
</head>

<body topmargin="5">
<%
tipo = request.querystring("tipo")
If IsEmpty(tipo) or tipo = "" then
%>
<center>
<table width="650">
<tr>
  <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" width="148">
    Novela
  </td>
  <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" width="70">
    Data
  </td>
  <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" width="231">
    MENSAGEM
  </td>
  <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" width="91">
    ATUALIZAR
  </td>
  <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" width="108">
    DELETAR
  </td>
</tr>
<%
DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
          DSNtemp=dsntemp & "DBQ=" & server.mappath("livro.mdb")
          sqlstmt = "SELECT * FROM assinantes ORDER BY data DESC"
          Set rs = Server.CreateObject("ADODB.Recordset")
          rs.Open sqlstmt, DSNtemp, 3, 3
TotalRecs = rs.recordcount
x = 0
For x = 1 to 9999
	If rs.eof then
		Exit For
	Else
   %>
    <tr>
      <td style="background-color: #f1f1f1; padding: 3" valign="top" align="left" width="148">
        <span style="color:#336699;font-weight:bold"><%=rs("nome")%></span><br>
      </td>
      <td style="background-color: #f1f1f1; padding: 3" valign="top" align="center" width="70">
        <%=rs("data")%>
      </td>
      <td style="background-color: #f1f1f1; border: 1px solid #000000; padding: 3" valign="middle" align="left" width="266">
        <%=rs("mensagem")%>
      </td>
      <td style="background-color: #f1f1f1; padding: 3" valign="middle" align="center" width="91">
        <a OnMouseOver="window.status='ATUALIZAR mensagem postada por <%=rs("nome")%>:: Data <%=rs("data")%>'; return true" href="administracao.asp?tipo=atualiza1&amp;id=<%=rs("id")%>nome=<%=rs("nome")%>&amp;mensagem=<%=rs("mensagem")%>&amp;site<%=rs("site")%>&amp;email=<%=rs("email")%>">ATUALIZAR</a>
      </td>
      <td style="background-color: #f1f1f1; padding: 3" valign="middle" align="center" width="108">
        <a OnMouseOver="window.status='DELETAR mensagem postada por <%=rs("nome")%>:: Data <%=rs("data")%>'; return true" href="administracao.asp?tipo=deleta&id=<%=rs("id")%>">DELETAR</a>
      </td>
    </tr>
    <%
rs.MoveNext
End If
Next%>
  <tr>
    <td colspan="3" valign="top" align="right" style="background-color: #f1f1f1; padding: 3" width="228">
      Total de Grava&ccedil;&otilde;es do banco:
    </td>
    <td valign="top" align="center" style="background-color: #f1f1f1; padding: 3" colspan="2" width="330">
     <span style="color:#ff0000;font-weight:bold"><%=TotalRecs%></span>
    </td>
  </tr>
</table>
<%
rs.close
set rs = nothing
%>
</center>
<p>
<%
end if
if tipo = "atualiza1" then
id = request.querystring("id")
if IsEmpty(id) or id = "" then
response.write "<br><br><br><br><center style=""width:310;background-color:#e3e3e3;color:#ff0000;border:1px solid #000000;padding:3;font-weight:bold"">"
response.write "Algum t&iacute;tulo deve ser selecionado para ser APAGADO do branco de dados!</center>"
else
nome = request.querystring("nome")
mensagem = request.querystring("mensagem")
site = request.querystring("site")
email = request.querystring("email")
%>
<center>
<form action="administracao.asp?tipo=atualiza" method="post">
<table width="450" colspan="2">
  <tr>
    <td style="background-color: #e3e3e3; font-weight: bold; color: #ff0000; padding: 3" valign="top" align="center" colspan="2" width="438">
      Atualizar dados do livro
    </td>
  </tr>
  <tr>
    <td style="background-color: #f1f1f1; padding: 3" valign="middle" align="right" width="85">
      Nome:      
    </td>
    <td  style="background-color: #f1f1f1; padding: 3" valign="middle" align="left" width="343">
      <input type="text" name="novela" size="53" id="caixa" value="<%=nome%>">
    </td>
  </tr>
  <tr>
    <td style="background-color: #f1f1f1; padding: 3" valign="middle" align="right" width="85">
      E-mail:      
    </td>
    <td  style="background-color: #f1f1f1; padding: 3" valign="middle" align="left" width="343">
      <input type="text" name="novela" size="53" id="caixa" value="<%=email%>">
    </td>
  </tr>
  <tr>
    <td style="background-color: #f1f1f1; padding: 3" valign="middle" align="right" width="85">
      Site:      
    </td>
    <td  style="background-color: #f1f1f1; padding: 3" valign="middle" align="left" width="343">
      <input type="text" name="novela" size="53" id="caixa" value="<%=site%>">
    </td>
  </tr>
  <tr>
    <td style="background-color: #f1f1f1; padding: 3" valign="top" align="right" width="85">
      Mensagem:      
    </td>
    <td  style="background-color: #f1f1f1; padding: 3" valign="middle" align="left" width="343">
      <textarea name="descricao" size="53" id="caixa" rows="4" cols="52"><%=mensagem%></textarea>
    </td>
  </tr>
    <td  style="background-color: #f1f1f1; padding: 3" valign="middle" align="center" colspan="2" >
      <input type="submit" value="Atualizar dados" id="bot">
    </td>
  <input type="hidden" name="id" value="<%=id%>">
</table>
</center>
</form>
<%
end if
end if
if tipo = "atualiza" then

mensagem = request.form("mensagem")
nome = request.form("nome")
id = request.form("id")

strCon = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("livro")
Set conn = Server.CreateObject("ADODB.Connection")
conn.open strCon
sql = "UPDATE assinantes SET nome = '" & nome & "', mensagem = '" & mensagem & "' WHERE id = " & id & ""
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn

if err.Number <> "0" then
Response.Write "<center style=""width:250"">Algo de errado aconteceu e as informa&ccedil;&atilde;es n&atilde; foram ATUALIZADAS</center>"
else
Response.Write "<br><br><br><center style=""width:250;background-color:#f1f1f1;padding:3;border:1px solid #ff0000"">Informa&ccedil;&otilde;es ATUALIZADAS com sucesso!"
Response.Write "<br><br><a href=""administracao.asp?tipo=redireciona""><b>Clique aqui</b> para atualizar ou deletar outra mensagem!</a></center><br><br><br><br><br><br><br>"
end if
end if%>
<%
if tipo = "deleta" then
id = request.querystring("id")
DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
          DSNtemp=dsntemp & "DBQ=" & server.mappath("livro.mdb")
			sql = "DELETE * FROM assinantes WHERE id = " & id & ""
			Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, DSNtemp
rs.close
set rs = nothing
%>
<br><br><br><center><span style="font-size:12;font-weight:bold;background-color:#f1f1f1;width:250;text-align:center;border:1px solid #000000;padding:5">
O registro<br><b><%=id%></b><br>foi deletado com sucesso!<br><br>
<a href="administracao.asp?tipo=redireciona">Clique aqui para deletar ou atualizar outro</a></span></center><br><br><br><br><br><br><br>
<%end if
if tipo = "redireciona" then
response.redirect "administracao.asp"
end if
end if%>
</body>
</html>