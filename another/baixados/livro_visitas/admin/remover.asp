<%
'Verifica se existe uma seзгo
if session("login") = "" then
'se nгo existir redireciona o usuбrio para a pбgina principal
response.Redirect("index.asp")
end if

'Pegamos o id da mensagem
id = request("id")

'Abre a conexгo com o banco de dados
Set Conex = Server.CreateObject ("ADODB.Connection") 
Conex.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath("../database/dados.mdb")

'Procuramos a mensagem com o respectivo id e removermos o registro
Set remover = Server.CreateObject("ADODB.Recordset")
sql = "Select * from mensagens where id="&id
remover.open sql, conex, 3,3
remover.delete

'Redirecionamos o usuario para a pбgina principal
response.Redirect("admin.asp")
%>