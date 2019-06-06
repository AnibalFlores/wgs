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
'// WG Sistemas (c) 2006
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
    sql = "SELECT Usuarios.ID_Usuarios, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Usuarios.Usr_Usuario, Usuarios.Usr_Pass, Filtros.fil_Nombre, Usuarios.Usr_Log, Usuarios.Usr_Nivel, Usuarios.Usr_Estado, Usuarios.Usr_Passtime FROM Filtros INNER JOIN Usuarios ON Filtros.ID_Filtros = Usuarios.Usr_Base"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="5">
  
  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">Usuarios Registrados </div></td>
  </tr>
<tr class="Msjerror">

<td align="center"><%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
<TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=1 BGCOLOR=#336699 class="Texbuche">
        <THEAD>
          <TR bordercolor="#CCCCCC" class="SloganTrebu">
            <TH height="21" >ID</TH>
            <TH height="21" >Apellidos</TH>
            <TH height="21" >Nombres</TH>
            <TH height="21" >Usuario</strong></TH>
            <TH height="21" >Clave</TH>
            <TH height="21" >Base</TH>
            <TH height="21" >Nivel</TH>
            <TH > Ingreso</TH>
            <TH >Egreso</TH>
            <TH >Estado</TH>
          </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
 %>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("ID_Usuarios").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Usuario").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Pass").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("fil_Nombre").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Nivel").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Usr_Log").Value)%></TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Usr_Passtime").Value)%></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Usr_Estado").Value)%></TD>
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