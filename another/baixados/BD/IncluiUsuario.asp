<!--#Include File= "conexao3h.asp"-->

<%
Dim campo, valor

campo = ""
valor = ""

if Request.Form("Usr_Email") <> "" then
	Set rstConsulta = conTurma3h.execute("Select * From Usuario Where Usr_Email='"& Request.Form("Usr_Email")& "'")
	if not rstConsulta.EOF then %>
		<script language="JavaScript">
			alert("Este E-mail já está cadastrado, por favor tente novamente !!!")
			history.back()
		</script>
	<%end if
else %>
	<script language="JavaScript">
		alert("Preencha o campo de e-mail !!!")
		history.back()
	</script>
<% end if

for each x in Request.Form
	if Request.Form(x) <> "" and x <> "confirma" then
		if campo = "" then
			campo = x
			valor = "'"&  Request.Form(x) & "'"
		else
			campo = campo & ", "& x
			valor = valor & ", '"&  Request.Form(x) & "'"
		end if 
	end if
next

Sql = "Insert Into Usuario (" & campo & ") Values ("& valor & ")"


Set rstInclui = Server.CreateObject("ADODB.Recordset")
rstInclui.Open Sql, conTurma3h 
Set rstConsulta = nothing
Set conTurma3h = nothing
Response.Write("SEU CADASTRO FOI BEM SUCEDIDO")
Response.Redirect "Usuario_lista.asp"
%>