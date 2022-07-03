<%
'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------

Dim ConnDB

Sub AbrirDB()

	Set ConnDB = Server.CreateObject("ADODB.Connection")
	conexao = "DBQ=" & Server.MapPath("usuarios.mdb")
	ConnDB.Open "DRIVER={Microsoft Access Driver (*.mdb)};" & conexao

end sub

Sub FecharDB()

	if ConnDB.state = 1 then
		ConnDB.Close
		Set ConnDB = Nothing
	end if

end sub

'---------------------------------------------------------------------------------------
'		Script by Fabio Franco
'	Email: fabio_franco@ofm.com.br ICQ: 164613668
'---------------------------------------------------------------------------------------
%>