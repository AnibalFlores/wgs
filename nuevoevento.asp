<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<TITLE>Lista Usuarios</TITLE>
<link href="wgs.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgcolor="#999999">
<%
clienteid = request("clienid")
Discar = request("Prefix")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
	Set rs = Server.CreateObject("ADODB.Recordset")
    
	sql = "SELECT * FROM Clientes WHERE Clientes.Cli_Numero="& clienteid 
	rs.Open sql, conn, 1, 2
	If rs.Fields("Cli_Status").Value=3 AND rs.Fields("Cli_Usuario").Value <> session("wgsANM_usr") 	  	    then
		rs.Close
		set rs=nothing
		Response.Redirect("listar.asp")
	Else
		rs("Cli_Status")=3
		rs("Cli_Usuario")= session("wgsANM_usr")
		rs.Update
	End if
	rs.Close
	sql = "SELECT Clientes.Cli_Numero, Clientes.Cli_Status, Clientes.Cli_Usuario, Eventos.ID_Eventos, Estados.Est_Nombre, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Eventos.Eve_Fecha, Eventos.Eve_Nota, Productos.Prod_Nombre, Eventos.Eve_Fechaprox, Usuarios.Usr_Nivel FROM (Usuarios RIGHT JOIN ((Clientes RIGHT JOIN Eventos ON Clientes.ID_Clientes = Eventos.Eve_Cliente) INNER JOIN Productos ON Eventos.Eve_Producto = Productos.ID_Producto) ON Usuarios.ID_Usuarios = Eventos.Eve_Usuario) INNER JOIN Estados ON Eventos.Eve_Tipo = Estados.ID_Estados WHERE (((Clientes.Cli_Numero)="& clienteid & "))"
    rs.Open sql, conn, 3, 3
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" colspan="2" align="left"><div align="center" class="Titulo">Historial de Contactos </div></td>
  </tr>
<tr class="Msjerror">

<td colspan="2" align="center"><%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
<TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=1 BGCOLOR=#336699 class="Texbuche">
        <THEAD>
          <TR bordercolor="#CCCCCC" class="SloganTrebu">
            <TH height="21" >Evento</TH>
            <TH height="21" >Apellidos</TH>
            <TH height="21" >Nombres</TH>
            <TH height="21" >Resultado</TH>
            <TH height="21" >Hecho</TH>
            <TH height="21" >Proximo</TH>
            <TH height="21" >Nivel</TH>
            <TH height="21" >Producto</TH>
          </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
contador = 5
do while Not rs.eof and contador > 0
 %>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("ID_Eventos").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Est_Nombre").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Eve_Fecha").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Eve_Fechaprox").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Nivel").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Prod_Nombre").Value)%></TD>
          </TR>
          <%
rs.MoveNext
contador = contador - 1
loop%>
        </TBODY>
      </TABLE>   
<%
rs.close
set rs=nothing
%></td>
  </tr>
<%End if%>
  <tr class="Titulo">
    <td height="33"><div align="center">LINEA A LLAMAR: <%=Discar%></div></td>
    <td height="33">Estado Cliente: Tomado</td>
  </tr>
  <tr bgcolor="#6699CC" class="Texbuche">
    <td height="33" bgcolor="#FF6666"><div align="center">No contesta, llamar en otro momento o dejar de llamar. </div></td>
    <td height="33" bgcolor="#66FF99"><div align="center">Para  cargar evento pendiente, negativo, rectifica o positivo. </div></td>
  </tr>
  <tr class="Texbuche">
    <td height="33" bgcolor="#FF6666">
      <div align="center"><span class="texto">
        <input name="button322" type="reset" class="btn" id="button322" onClick="javascript:location.href='nocontacto.asp?clienid=<%=Clienteid%>'" value="No Contacto">
      </span></div></td>
    <td bgcolor="#66FF99"><div align="center"><span class="texto">
      <input name="button3222" type="reset" class="btn" id="button3222" onClick="javascript:location.href='contacto.asp?clienid=<%=Clienteid%>'" value="Contacto">
    </span></div></td>
  </tr>
  <tr class="Texbuche">
    <td height="33" colspan="2"><div align="center">
      <%
conn.close
set conn=nothing
Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
</BODY>
</HTML>
