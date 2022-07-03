<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
    <!--Ol galera, espero que usem este cdigo com queira, mas peo
apenas que no tirem esta nota e que enviem sugestes para mim, para
que possa estar melhorando cada vez mais!

Marcos Oliveira, meu e-mail : marcos_804@yahoo.com.br, 
desde j agradeo! -->
<%
tipo = request.querystring("tipo")
If IsEmpty(tipo) or tipo = "" then
%>
<form action="assina.asp?tipo=insere" method="post">
<table colspan="2" width="500" style="border:1px solid #000000">
  <tr>
    <td colspan="2" class="1" width="488">
      Livro de Visitas
    </td>
  </tr>
  <tr>
    <td width="102" class="2" valign="middle" align="right">
      <b>Nome</b>:
    </td>
    <td width="384" class="2" valign="middle" align="left">
      <input type="text" name="nome" value="" class="caixa" size="44">
    </td>
  </tr>
  <tr>
    <td width="102" class="2" valign="middle" align="right">
      <b>E-mail</b> *:
    </td>
    <td width="384" class="2" valign="middle" align="left">
      <input type="text" name="email" value="" class="caixa" size="44">
    </td>
  </tr>
  <tr>
    <td width="102" class="2" valign="middle" align="right">
      <b>Site</b>:
    </td>
    <td width="384" class="2" valign="middle" align="left">
      <input type="text" name="site" value="http://" class="caixa" size="44">
    </td>
  </tr>
  <tr>
    <td width="102" class="2" valign="top" align="right">
      <b>Mensagem</b> *:
    </td>
    <td width="384" class="2" valign="middle" align="left">
      <textarea cols="44" rows="4" class="caixa" name="mensagem"></textarea>
    </td>
  </tr>
  <tr>
    <td colspan="2" class="2" style="text-align:center;font-size:12" valign="top">
      <input type="submit" value="Assinar" id="bot"> &nbsp;&nbsp;<input type="reset" value="Limpar" id="bot">
    </td>
  </tr>
  <tr>
    <td colspan="2" class="2" style="text-align:left;font-size:12" valign="top">
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
    </td>
  </tr>
  <tr>
    <td colspan="2" class="2" style="text-align:left;font-size:12" valign="top">
      ATEN&Ccedil;&Atilde;O: Os campos marcados com &quot; <b>*</b> &quot; s&atilde;o
      de preenchimento obrigat&oacute;rio!
    </td>
  </tr>
</table>
</form>
<%
'Aqui finaliza a parte onde mostra ao internauta a pagina de assinatura
end if
'Aqui comea a pagina onde vai inserir os dados
if tipo = "insere" then

'AQUI BUSCA AS INFORMACOES DO FORMULARIO
Session.LCID = 1046
nome = request.form("nome")
email = request.form("email")
site = request.form("site")
mensagem = request.form("mensagem")
data = date()
hora = Hour(time) + 3 & ":" & Minute(time) & ":" & Second(time)
ip = request.servervariables("REMOTE_ADDR")
mensagem = Replace(mensagem, chr(39), "&quot;")
mensagem = Replace(mensagem, "<", "&lt;")
mensagem = Replace(mensagem, ">", "&gt;")
mensagem = Replace(mensagem, VbCrLf, "<br>")

'MODIFICA DADOS NAO OBRIGATORIOS
If IsEmpty(nome) or nome = "" then
nome = "Anônimo"
else
nome = nome
end if
If site = "http://" then
site = "0" 
else
site = site
end if

'AQUI NOTIFICA A ALGUM CAMPO OBRIGATORIO VAZIO
If email = "" or mensagem = "" then
response.write "<br><br><br><table><tr><td>Erro!</td></tr>"
response.write "<tr><td>ALGUM CAMPO OBRIGATÓRIO NÃO FOI PREENCHIDO!<BR>"
response.write "<BR><a href=javascript:history.go(-1);>CLIQUE AQUI PARA VOLTAR E PREENCHE-LO</a></td></table><br><br><br><br><br><br>"
else

'AQUI ESTA ABRINDO O BANCO E INSERINDO AS INFORMACOES NO MESMO!
SET kikeh = CreateObject("ADODB.Connection")
meuLivro = "DRIVER={Microsoft Access Driver (*.mdb)};"
meuLivro = meulivro &"DBQ="& Server.MapPath("livro.mdb")
kikeh.open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=livro.mdb"
SQL = "INSERT INTO assinantes(nome, email, site, mensagem, data, hora, ip)"
SQL = SQL & "VALUES('"& nome &"','"& email &"','"& site &"','"& mensagem &"','"& data &"','"& hora &"','"& ip &"')"
set GRAVAR = marcos.EXECUTE(SQL)

'AQUI REDIRECIONA A PAGINA DA MENSAGENS JA INSERIDA NO BANCO DE DADOS!
response.redirect "mensagem.asp"

'AQUI TERMINA A NOTIFICACAO DE ALGUM CAMPO VAZIO
end if
'AQUI TERMINA DE RODAR A PAGINA DE INSERSAO DE DADOS
end if
%>
</body>
</html>