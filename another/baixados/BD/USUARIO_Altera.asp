<%Option Explicit%>
<!--#Include File="conexao3h.asp"-->

 <%Dim campo_valor,valor_campo,Sql,rstInclui,x
 campo_valor=""
 valor_campo=""


for each x in Request.Form
	if x <> "confirma" and x <> "Usr_Codigo" then
		if Request.Form(x) = "" then
			valor_campo = "null"
		else
			valor_campo = "'" & Request.Form(x) & "'"
		end if
		if campo_valor = "" then
			campo_valor = x & " = "&  valor_campo 
		else
			campo_valor = campo_valor & ", "& x & " = "&  valor_campo 
		end if 
	end if
next

Sql = "Update Usuario set " & campo_valor & " where Usr_Codigo= "& Request.Form("Usr_Codigo") 

'Response.write sql
'Response.end

Set rstInclui=Server.CreateObject("ADODB.RecordSet")
rstInclui.Open Sql, conTurma3h

Set rstInclui = nothing
Set conTurma3h = nothing

Response.Redirect "Usuario_Lista.asp"
%>