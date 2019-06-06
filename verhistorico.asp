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
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
    sql = "SELECT Clientes.Cli_Numero, Eventos.ID_Eventos, Estados.Est_Nombre, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Eventos.Eve_Fecha, Eventos.Eve_Nota, Productos.Prod_Nombre, Eventos.Eve_Fechaprox, Usuarios.Usr_Nivel FROM (Usuarios RIGHT JOIN ((Clientes RIGHT JOIN Eventos ON Clientes.ID_Clientes = Eventos.Eve_Cliente) INNER JOIN Productos ON Eventos.Eve_Producto = Productos.ID_Producto) ON Usuarios.ID_Usuarios = Eventos.Eve_Usuario) INNER JOIN Estados ON Eventos.Eve_Tipo = Estados.ID_Estados WHERE (((Clientes.Cli_Numero)="& clienteid & ")) ORDER BY Eventos.Eve_Fecha DESC"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">Historial de Contactos </div></td>
  </tr>
<tr class="Msjerror">

<td align="center"><%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
<TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=1 BGCOLOR=#336699 class="Texbuche">
        <THEAD>
          <TR bordercolor="#CCCCCC" class="SloganTrebu">
            <TH height="21" >Evento</TH>
            <TH height="21" >Asesor</TH>
            <TH height="21" >Resultado</TH>
            <TH height="21" >Hecho</TH>
            <TH height="21" >Proximo</TH>
            <TH height="21" >Producto</TH>
          </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
ID_Eventos = HTMLEncode(rs.Fields("ID_Eventos").Value)%>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24"  ALIGN=center><a href="verevento.asp?evenid=<%=ID_Eventos%>"><%=ID_Eventos%></a></TD>
            <TD height="24"  ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Nivel").Value)%>:&nbsp;<%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%>,&nbsp;<%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%></TD>
            <TD height="24"  ALIGN=center><%=HTMLEncode(rs.Fields("Est_Nombre").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Eve_Fecha").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Eve_Fechaprox").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Prod_Nombre").Value)%></TD>
          </TR>
          <%
rs.MoveNext
loop%>
        </TBODY>
      </TABLE>
   </td>
  </tr>
<%end if%>
  <tr class="Texbuche">
    <td height="33"><div align="center">
      <input type=button class="btn" onClick="history.go(-1)" value="Volver">
      </p>
    </div>
    <div align="center">
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
</BODY>
</HTML>