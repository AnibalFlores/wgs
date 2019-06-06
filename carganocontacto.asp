<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
sql = "SELECT Clientes.ID_Clientes, Clientes.Cli_Status, Clientes.Cli_Fechaultimo, Clientes.Cli_Fechaprox FROM Clientes WHERE Clientes.Cli_Numero="& clienteid
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2
sql= "Select * FROM Estados WHERE Estados.Id_Estados="& Request.Form("lst_Estados")
	Set rsEstados = Server.CreateObject("ADODB.Recordset")
	rsEstados.Open sql, conn, adOpenStatic, adLockOptimistic
sql= "Select * FROM Eventos"	
	Set rsE = Server.CreateObject("ADODB.Recordset")
	rsE.Open sql, conn, 1, 2

rs.Fields("Cli_Status") = rsEstados.Fields("Est_Logica")
rs.Fields("Cli_Fechaultimo") = Now()

rsE.AddNew
rsE.Fields("Eve_Fecha") = Now()
rsE.Fields("Eve_Usuario") = session("wgsANM_usr")
rsE.Fields("Eve_Cliente") = rs.Fields("ID_Clientes")

if IsDate(Request.Form("f_proximo")) then
	rsE.Fields("Eve_Fechaprox") = Request.Form("f_proximo")
	rs.Fields("Cli_Fechaprox") = Request.Form("f_proximo")
else
	rsE.Fields("Eve_Fechaprox") = Null
	rs.Fields("Cli_Fechaprox") = Null
End if

rsE.Fields("Eve_Tipo") = Request.Form("lst_Estados")
rsE.Fields("Eve_Producto") = 1
rsE.Fields("Eve_Aparato") = 96
rsE.Fields("Eve_Nota") = Request.Form("txt_nota")
rsE.Update 
rs.Update
rsEstados.close
set rsEstados=nothing
rsE.Close
set rsE=nothing
rs.Close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect("exitosa.asp")  
%>
