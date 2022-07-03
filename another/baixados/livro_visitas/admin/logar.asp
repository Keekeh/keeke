<!--#include file = "config.asp"-->
<%
'Pega os dados digitados no formulário
usuario = request("usuario")
senha = request("senha")

'Verifica se os dados estão corretos, se estiverem
if usuario = user and senha = password then
'Cria uma seção com o nome do usuário
session("login") = true
'Redireciona ele para a administração
response.Redirect("admin.asp")
'Cao contrário
else
'Redireciona ele para a página principal
response.Redirect("index.asp")
end if

%>