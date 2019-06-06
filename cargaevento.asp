<!--#include file="Connections/database.asp" -->
<%
if session("wgsANM_usr")="" then
    Response.Redirect "Logout.asp"
End if
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
sql = "SELECT Clientes.ID_Clientes, Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_TipoDNI, Clientes.Cli_DNInro, Clientes.Cli_Nacimiento, Clientes.Cli_Genero, Clientes.Cli_Email, Status.Sta_Nombre, Empresas.Emp_Corto, Clientes.Cli_Contacto, Clientes.Cli_Fechaultimo, Clientes.Cli_Fechaprox, Clientes.Cli_Comentario, Clientes.Cli_Telefono2, Clientes.Cli_CPNuevo, Clientes.Cli_Eventos, Clientes.Cli_Valor, Clientes.Cli_Usuario, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Localidades.Loc_Localidad, Localidades.Loc_Provincia, Localidades.Loc_CP, Localidades.Loc_CPAprov, Clientes.Cli_Inicio, Clientes.Cli_Calle, Clientes.Cli_CalleNro, Clientes.Cli_Piso, Clientes.Cli_Depto, Clientes.Cli_Ampli, Clientes.Cli_Inicio, Clientes.Cli_Status FROM Localidades INNER JOIN ((Empresas INNER JOIN (Status INNER JOIN Clientes ON Status.ID_Status = Clientes.Cli_Status) ON Empresas.ID_Empresas = Clientes.Cli_Empresa) INNER JOIN Usuarios ON Clientes.Cli_Usuario = Usuarios.ID_Usuarios) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE (((Clientes.Cli_Numero)="& clienteid & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, adOpenKeyset, adLockPessimistic
sql= "Select * FROM Eventos"	
	Set rsE = Server.CreateObject("ADODB.Recordset")
	rsE.Open sql, conn, adOpenKeyset, adLockPessimistic
sql= "Select * FROM Estados WHERE Estados.Id_Estados="& Request.Form("lst_Estados")
	Set rsEstados = Server.CreateObject("ADODB.Recordset")
	rsEstados.Open sql, conn, adOpenStatic, adLockOptimistic	
rs.Fields("Cli_Usuario") = session("wgsANM_usr")
rs.Fields("Cli_Apellido") = Request.Form("txt_Apellidos")
rs.Fields("Cli_Nombre") = Request.Form("txt_Nombres")
rs.Fields("Cli_TipoDNI") = Request.Form("lst_TipoDNI")
rs.Fields("Cli_DNInro") = Request.Form("txt_NroDNI")
rs.Fields("Cli_Genero") = Request.Form("lst_Genero")

If IsDate(Request.Form("f_nacim")) then
rs.Fields("Cli_Nacimiento") = Request.Form("f_nacim")
else
rs.Fields("Cli_Nacimiento") = Null
End if

rs.Fields("Cli_Calle") = Request.Form("txt_calle")
rs.Fields("Cli_CalleNro") = Request.Form("txt_calleNro")

If Request.Form("txt_piso")<>"" then
rs.Fields("Cli_Piso") = CByte(Request.Form("txt_piso"))
else
rs.Fields("Cli_Piso") = Null
End if

rs.Fields("Cli_Depto") = Request.Form("txt_dpto")
rs.Fields("Cli_Ampli") = Request.Form("txt_ampli")
rs.Fields("Cli_CPNuevo") = Request.Form("txt_cpa")
rs.Fields("Cli_Contacto") = Request.Form("txt_contacto")
rs.Fields("Cli_Email") = Request.Form("txt_email")
rs.Fields("Cli_Telefono2") = Request.Form("txt_fijo")
rs.Fields("Cli_Comentario") = Request.Form("txt_Observa")
rs.Fields("Cli_Comentario") = Request.Form("txt_Observa")

If IsDate(Request.Form("f_Proximo")) then
rs.Fields("Cli_Fechaprox") = Request.Form("f_Proximo")
else
rs.Fields("Cli_Fechaprox") = Null
End if
'//Fija el al cliente con la logica de Status correspondiente//'
rs.Fields("Cli_Status") = rsEstados.Fields("Est_Logica")
rs.Fields("Cli_Fechaultimo") = Now()
if Not IsDate(rs.Fields("Cli_Inicio").Value) then
rs.Fields("Cli_Inicio") = Now()
End if 
rs.Update 
rsE.AddNew
rsE.Fields("Eve_Fecha") = Now()
rsE.Fields("Eve_Usuario") = session("wgsANM_usr")
rsE.Fields("Eve_Cliente") = rs.Fields("ID_Clientes")

If IsDate(Request.Form("f_Proximo")) then
rsE.Fields("Eve_Fechaprox") = Request.Form("f_Proximo")
else
rsE.Fields("Eve_Fechaprox") = Null
End If

rsE.Fields("Eve_Tipo") = Request.Form("lst_Estados")
rsE.Fields("Eve_Producto") = Request.Form("lst_Producto")
rsE.Fields("Eve_Aparato") = Request.Form("lst_Modelo")
rsE.Fields("Eve_Nota") = Request.Form("txt_nota")
rsE.Update 

'//Si es "Positivo a Confirmar" envio mail//'
If Request.Form("lst_Estados")=1 then
Set Mail = Server.CreateObject ("CDO.Message")
Mail.From = Remitente
Mail.To = Receptor
Mail.Subject = "Solicitud [" & clienteid & "]"
Mail.TextBody = "Cuerpo central del mensaje a enviar " & session("wgsANM_usr") & " " & Request.Form("txt_Apellidos") & "Evento: " & rsE.Fields("Id_Eventos") & "Cliente: " & rs.Fields("ID_Clientes")
Mail.Send
Set Mail = Nothing
End If


rsEstados.close
set rsEstados=nothing
rsE.close
set rsE=nothing
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect("exitosa.asp")   
%>
