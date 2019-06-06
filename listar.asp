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
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql="SELECT * From Clientes WHERE Clientes.Cli_Status=3 AND Clientes.Cli_Usuario=" & session("wgsANM_usr")
	rs.Open sql, conn, adOpenStatic, adLockOptimistic
	if Not rs.eof then
	redir = "vercliente.asp?clienid="&rs.Fields("Cli_Numero")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect(redir)
	End if 
    sql = "SELECT Usuarios.ID_Usuarios, Filtros.fil_sql FROM Filtros INNER JOIN Usuarios ON Filtros.ID_Filtros = Usuarios.Usr_Base WHERE (((Usuarios.ID_Usuarios)=" & session("wgsANM_usr") & "))"
	rs.close
    rs.open sql, conn, adOpenStatic, adLockOptimistic

	UsrFiltro = rs("fil_sql").value
	
	if UsrFiltro = "*" then 
			UsrFiltro = " "
	else
			UsrFiltro = "AND " & UsrFiltro
	End if
		
    sql = "SELECT Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_Fechaprox, Empresas.Emp_Corto, Status.Sta_Nombre, Localidades.Loc_Localidad FROM Localidades INNER JOIN (Status INNER JOIN (Empresas INNER JOIN Clientes ON Empresas.ID_Empresas = Clientes.Cli_Empresa) ON Status.ID_Status = Clientes.Cli_Status) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE Sta_Camino " & UsrFiltro & " ORDER BY Clientes.Cli_Fechaprox ASC"
    rs.close
    rs.open sql, conn, adOpenStatic, adLockOptimistic
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo"><span class="verdana9"><strong>Clientes Disponibles</strong></span></div></td>
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
count = 10
do while Not rs.eof and count>0
 Cli_Numero=HTMLEncode(rs.Fields("Cli_Numero").Value)%>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF">
            <TD height="24" ALIGN=center><a href="vercliente.asp?clienid=<%=Cli_Numero%>"><%=Cli_Numero%></a></TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Cli_Apellido").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Cli_Nombre").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Loc_Localidad").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center><%=HTMLEncode(rs.Fields("Sta_Nombre").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center ><%=HTMLEncode(rs.Fields("Emp_Corto").Value)%>&nbsp;</TD>
            <TD height="24" ALIGN=center><%=FormatMediumDate(rs.Fields("Cli_Fechaprox").Value)%>&nbsp;</TD>
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
      <input name="Mis Pendientes" type="button" class="btn" onClick="javascript:parent.main.location.href='verpendientes.asp'" value="Mis Pendientes"main"'">
    </div></td>
  </tr>
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
<div align="center"></div>
</BODY>
</HTML>