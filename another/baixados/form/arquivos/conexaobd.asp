<%@ Language=VBScript %>
<%
option explicit         

	dim connstring
	
	response.expires=-1     

	Set conconexao = Server.CreateObject("ADODB.Connection")
	conconexao.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&Server.MapPath("dados.mdb")
%>
