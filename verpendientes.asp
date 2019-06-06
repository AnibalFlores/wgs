<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<TITLE>ListaPrestaciones</TITLE>
<link href="wgs.css" rel="stylesheet" type="text/css"></HEAD>
<BODY bgcolor="#999999">
<%
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
    Set Session("baseWG.mdb_conn") = conn
   sql = "SELECT Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_Fechaprox, Empresas.Emp_Corto, Status.Sta_Nombre, Localidades.Loc_Localidad FROM Localidades INNER JOIN (Status INNER JOIN (Empresas INNER JOIN Clientes ON Empresas.ID_Empresas = Clientes.Cli_Empresa) ON Status.ID_Status = Clientes.Cli_Status) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE Sta_Camino AND Status.Sta_Nombre = 'Pendiente' AND Cli_Usuario=" & session("wgsANM_usr") & " ORDER BY Clientes.Cli_Fechaprox ASC"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo"><span class="verdana9"><strong>Mis Pendientes Disponibles</strong></span></div></td>
  </tr>
<tr class="Msjerror">

<td align="center" class="verdana9"><%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
  <TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=1 BGCOLOR=#336699 class="Texbuche">
        <CAPTION>&nbsp;
        </CAPTION>
        <THEAD>
          <TR align="center" valign="middle" bordercolor="#CCCCCC" class="SloganTrebu">
            <TH class="texto9normal" ><strong>L&iacute;nea</strong></TH>
            <TH class="texto9normal" ><strong>Apellidos</strong></TH>
            <TH class="texto9normal" ><strong>Nombres</strong></TH>
            <TH class="texto9normal" ><strong>Localidad</strong></TH>
            <TH class="texto9normal" ><strong>Estado</strong></TH>
            <TH class="texto9normal" ><strong>Empresa</strong></TH>
            <TH class="texto9normal" ><strong>Fecha</strong></TH>
          </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
count = 25
do while Not rs.eof and count>0
 Cli_Numero=rs("Cli_Numero")%>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF">
            <TD height="24"  ALIGN=center><a href="vercliente.asp?clienid=<%=Cli_Numero%>"><%=Server.HTMLEncode(rs.Fields("Cli_Numero").Value)%></a></TD>
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Cli_Apellido").Value)%>&nbsp;</TD>
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Cli_Nombre").Value)%>&nbsp;</TD>
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Loc_Localidad").Value)%>&nbsp;</TD>
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Sta_Nombre").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center ><%=Server.HTMLEncode(rs.Fields("Emp_Corto").Value)%>&nbsp;</TD>
            <TD height="24"  ALIGN=center><%=FormatMediumDate(rs.Fields("Cli_Fechaprox").Value)%>&nbsp;</TD>
          </TR>
          <%
rs.MoveNext
count= count-1
loop%>
        </TBODY>
      </TABLE>    </td>
  </tr>
<%end if%>
  <tr class="Texbuche">
    <td height="33"><div align="center">
      <input name="button" type=button class="btn" onClick="history.go(-1)" value="Volver">
    </div></td>
  </tr>
  <tr class="Texbuche">
    <td height="33"><div align="center">
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
Set Session("baseWG.mdb_conn") = nothing
Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
<div align="center"></div>
</BODY>
</HTML>