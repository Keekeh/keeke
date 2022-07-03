<!--#Include file="cabecalho.asp"-->
<!--#Include file="conexao3h.asp"-->
<HTML>
<HEAD>
<TITLE>Contatos</TITLE>
</HEAD>
<BODY>
	<% Dim rstContatos
		Set rstContatos=Server.CreateObject("ADODB.Recordset")
		rstContatos.open "Select con_Nome, con_Email from Contato", conTurma3h
		If rstContatos.EOF then
			response.write "Não há contato cadastrado!"
		Else
		While not  rstContatos.EOF
			Response.write rstContatos ("con_Nome") & "_"
			Response.write rstContatos ("con_Email") & "<br>"
			rstContatos.movenext
		wend
		rstContatos.close
		end if
		conTurma3H.close
		%>
		
	<!--#Include file="rodape.asp"-->	
	
</BODY>    
</HTML>