<html>

<head>
<title>Mensagens do Livro de Visitas</title>
<style>
body {
scrollbar-face-color: 336699;
scrollbar-highlight-color: #ffff00; 
scrollbar-shadow-color: #ffff00;
scrollbar-3dlight-color: #0718ff;
scrollbar-arrow-color: #ffff00; 
scrollbar-darkshadow-color: #0718ff; 
scrollbar-base-color: #ffff00;}

.barra {background-color:#336699;padding:2;height:18;font-weight:bold;text-align:center;font-family:verdana;font-size:12;color:#ffffff;margin-bottom:3}
#barra {margin-top:3;background-color:#336699;padding:2;height:18;font-weight:bold;text-align:center;font-family:verdana;font-size:12;color:#ffffff;margin-bottom:3}

a{text-decoration:none;color:#000000}
a:hover{text-decoration:underline}
#botao {background-color:#1E80FF;color:#ffffff;font-weight:bold;text-align:center;font-family:verdana;font-size:11;padding:1;width:150;border:1px solid #ffff00}
#botao a{background-color:#336699;color:#ffffff;font-weight:bold;text-align:left;font-family:verdana;font-size:11;padding-left:3;width:100%;height:18;border-left:1px solid #ffff00;border-right:1px solid #ffff00;border-bottom:1px solid #ffff00;cursor:help}
#botao a:hover{background-color:#ffff00;color:#ff0000;font-weight:bold;text-align:left;font-family:verdana;font-size:11;padding-left:3;width:100%;height:18;border:1px solid #ff0000;text-decoration:none;cursor:help}

body{font-family:verdana;font-size:12;color:#000000}
table{font-family:verdana;font-size:12;color:#000000}

.1{background-color:#c0c0c0;text-align:center;font-weight:bold;padding:3}
.2{background-color:#f1f1f1;font-weight:normal;padding:3}
.caixa{background-color:#f1f1f1;font-family:verdana;fotn-size:12;color:#ff0000;border:1px solid #000000}

#link {font-family:Verdana;font-size:11;color:#000000;font-weight:bold}
#link a{font-family:Verdana;font-size:11;color:#000000;font-weight:bold;width:100%;padding:3;border:1px solid #000000;cursor:default}
#link a:hover{font-family:Verdana;font-size:11;color:#000000;font-weight:bold;text-decoration:none;cursor:default;background-color:#ffff00}

#pagina {font-family:verdana;font-size:11;color:#000000;font-weight:bold;width:100%}
#pagina a{height:15;border:1px solid #ffffff;width:18;color:#336699;background-color:#CEE6FF;padding:3;background-color:#ffffff}
#pagina a:hover{height:15;border:1px solid #000000;width:18;color:#ff0000;background-color:#CEE6FF;font-family:verdana;text-decoration:none}
</style>

</head>

<body topmargin="0">

<center style="margin-top:0">

<%
'AQUI VOCE VAI CONFIGURAR O NUMERO DE GRAVACOES DO BANCO QUE SERA EXIBIDO POR CADA PAGINA
Const GravacoesPorPagina = 3 'MUDE AQUI SE QUISER 

If Request.QueryString("PaginaAtual") = "" Then
	PosicaoDaPagina = 1

Else
	PosicaoDaPagina = CInt(Request.QueryString("PaginaAtual"))
End If	

%>  
<%
Set con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
'Abre conexao com o banco de dados
livro = "DRIVER={Microsoft Access Driver (*.mdb)};"
livro = livro & "DBQ=" & Server.MapPath("livro.mdb")
SQL = "SELECT * FROM assinantes ORDER By id DESC"
rs.Open SQL, livro, 3

TotalDEGravacoes = rs.RecordCount
TotalDePaginas = int(TotalDEGravacoes/GravacoesPorPagina)
If TotalDEPaginas MOD GravacoesPorPagina <> 0 Then
TotalDEPaginas = TotalDePaginas + 1
end if

%>
<table width="500" colspan="2">
  <tr>
    <td id="barra" valign="middle" align="center" colspan="2" width="490">
      Mensagens do Livro de Visitas
    </td>
  </tr>
  <tr>
    <td valign="middle" align="left" colspan="2" width="492">
      <span style="height:15;border:1px solid #000000;width:150;color:#000000;background-color:#f1f1f1;padding:3;font-family:arial;text-decoration:none;font-size:11;font-weight:bold">
        P&aacute;gina &nbsp;<span style="color:#ff0000;font-style:italic"><%=PosicaoDaPagina%></span>&nbsp; de &nbsp;<span style="color:#ff0000;font-style:italic"><%If TotalDePaginas = "0" then response.write "1" else response.write TotalDePaginas end if%></span>
        </span><%response.write "&nbsp;&nbsp;&nbsp;&copy; Marcos Oliveira"%>
    </td>
  </tr>
<%
   rs.PageSize = GravacoesPorPagina
   If NOT rs.EOF Then rs.AbsolutePage = PosicaoDaPagina
   For RepeteGravacoes = 1 to GravacoesPorPagina
	If rs.EOF Then Exit For
%>
  <tr>
    <td class="2" width="100" valign="middle" align="right">
      Assinante:<!--Livro de Visitas feito por Marcos Oliveira
                email: marcos_804@yahoo.com.br  -->
    </td>
    <td width="382" class="2">
      <span style="color:#ff0000;font-weight:bold"><%=rs("nome")%></span>
    </td>
  </tr>
  <tr>
    <td class="2" width="100" valign="middle" align="right">
      E-mail:
    </td>
    <td width="382" class="2">
      <a href="mailto:email"><span style="font-weight:bold;font-style:italic"><%=rs("email")%></span></a>
    </td>
  </tr>
  <%if rs("site") <> "0" then%>
  <tr>
    <td class="2" width="100" valign="middle" align="right">
      Site:
    </td>
    <td width="382" class="2">
      <a target="janela" href="<%=rs("site")%>"><span style="font-style:italic;font-weight:bold"><%=rs("site")%></span>
    </td>
  </tr>
  <%else
   response.write "&nbsp"
   end if%>
  <tr>
    <td class="2" width="100" valign="top" align="right">
      Mensagem:
    </td>
    <td width="382" class="2" style="background-color:#f1f1f1;color:#ff0000;font-weight:bold;border:1px solid #000000">
      <%=rs("mensagem")%> 
    </td>
  </tr>
  <tr>
    <td colspan="2">
      <hr>
    </td>
  </tr>
<%
rs.MoveNext
Next%>
  <tr>
    <td colspan="2">
      <table width="100%">
        <tr>
          <td width="50%" id="link" valign="middle" align="center" bgcolor="#FFFFFF">
            <a href="assina.asp">Assinar no livro</a>
          </td>
          <td width="50%" id="link" valign="middle" align="center" bgcolor="#FFFFFF">
            <a href="mensagem.asp">Ver mensagens postadas</a>
          </td>
        </tr>        
      </table>
      <table width="100%">
        <tr>
          <td width="15%" style="background-color:#f1f1f1">
            <div id="link">
              <%
              'CRIA UM LINK PARA A PAGINA ANTERIOR SE A PAGINA FOR MAIOR QUE 1    	
              If PosicaoDaPagina > 1 Then 
              	Response.Write "	<a href=""mensagem.asp?PaginaAtual=" &  PosicaoDaPagina - 1  & """ target=""_self"">P&aacute;gina Anterior<br>&lt;&lt;&lt;</a>"   	     	
              End If
              %>
            </div>
          </td>
          <td valign="top" align="center" style="font-weight:bold;color:#ff0000;background-color:#f1f1f1;padding:3">
            <div id="pagina">
            <%
            If TotalDePaginas > 1 Then
              Response.Write "	P&aacute;ginas do Livro "
            End If
            %><br>
            <%
            atual = request.querystring("PaginaAtual")
            For paginass = 1 to TotalDEPaginas
            If paginass = PosicaoDaPagina then
            Response.Write "&nbsp;<span style=""height:15;border:1px solid #000000;width:18;color:#ffff00;background-color:#336699;padding:3;font-family:arial;text-decoration:none;font-size:11;font-weight:bold"">"
            else
            Response.Write "&nbsp;<a onmouseover=""window.status='Clique para ir &agrave; p&aacute;gina: " & paginass & "';return true"" href=""mensagem.asp?PaginaAtual=" & Paginass & """>" 
            end if
            Response.Write Paginass
            If paginass = PosicaoDaPagina then
            Response.Write "</span>&nbsp;"
            else
            Response.write "</a>&nbsp;"
            end if
            Next
           %></div>
          </td>
          <td width="15%" style="background-color:#f1f1f1">
            <div id="link">
              <%
              'CRIA UM LINK PARA A PROXIMA PAGINA SE EXISTIR A TAL      	
               If NOT rs.EOF then   	
               	Response.Write "	<a href=""mensagem.asp?PaginaAtual=" &  PosicaoDaPagina + 1  & """ target=""_self"">Pr&oacute;xima p&aacute;gina<br>&gt;&gt;&gt;</a>"	   	
               End If 
               %>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
'Reseta os objetos do servidor
Set con = Nothing
rs.Close
Set rs = Nothing       
%>
</body>
</HTML>
