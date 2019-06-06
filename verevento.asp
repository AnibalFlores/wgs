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
eventoid=request("evenid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
    sql = "SELECT Eventos.ID_Eventos, Eventos.Eve_Tipo, Eventos.Eve_Fecha, Eventos.Eve_Fechaprox, Eventos.Eve_Nota, Clientes.Cli_Numero, Eventos.Eve_Cliente, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Usuarios.Usr_Nivel, Estados.Est_Nombre, Productos.Prod_Nombre, Productos.Prod_Codigo, Marcas.Mar_Nombre, Aparatos.Ap_Modelo FROM Clientes RIGHT JOIN (Marcas INNER JOIN (Aparatos INNER JOIN (Usuarios RIGHT JOIN (Productos INNER JOIN (Estados INNER JOIN Eventos ON Estados.ID_Estados = Eventos.Eve_Tipo) ON Productos.ID_Producto = Eventos.Eve_Producto) ON Usuarios.ID_Usuarios = Eventos.Eve_Usuario) ON Aparatos.Id_Aparato = Eventos.Eve_Aparato) ON Marcas.ID_Marcas = Aparatos.Ap_Marca) ON Clientes.ID_Clientes = Eventos.Eve_Cliente WHERE (((Eventos.ID_Eventos)=" & eventoid &"))"

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">Detalle del Evento <%=HTMLEncode(rs.Fields("ID_Eventos").Value)%></div></td>
  </tr>
<tr class="Msjerror">

<td align="center"><%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
<TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=0 BGCOLOR=#336699 class="Texbuche">
        <THEAD>
          <TR bordercolor="#CCCCCC" class="SloganTrebu">
            <TH height="21" colspan="2" >Resultado: <%=HTMLEncode(rs.Fields("Est_Nombre").Value)%></TH>
          </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
 %>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="right">Linea:</div></TD>
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="left"><b><%=HTMLEncode(rs.Fields("Cli_Numero").Value)%></b></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD width="50%" height="24"  ALIGN=center bgcolor="#F0F0F0"><div align="right">Realizado por el <%=HTMLEncode(rs.Fields("Usr_Nivel").Value)%>:</div></TD>
            <TD width="50%" height="24"  ALIGN=center bgcolor="#F0F0F0"><div align="left"><b><%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%>,&nbsp;<%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%></b></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="right">El d&iacute;a:</div></TD>
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="left"><b><%=FormatMediumDate(rs.Fields("Eve_Fecha").Value)%></b></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD width="50%" height="24"  ALIGN=center bgcolor="#F0F0F0"><div align="right">Proximo evento: </div></TD>
            <TD width="50%" height="24"  ALIGN=center bgcolor="#F0F0F0"><div align="left"><b><%=FormatMediumDate(rs.Fields("Eve_Fechaprox").Value)%></b></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="right">Producto Ofrecido: </div></TD>
            <TD width="50%" height="24"  ALIGN=center bgcolor="#CCCCCC"><div align="left"><b><%=HTMLEncode(rs.Fields("Prod_Nombre").Value)%></b></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24" colspan="2"  ALIGN=center bgcolor="#F0F0F0"><div align="center">Aparato Ofrecido</div>
            <div align="left"></div></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#CCCCCC" class="Texbuche">
            <TD height="24" colspan="2"  ALIGN=center>Marca <b><%=HTMLEncode(rs.Fields("Mar_Nombre").Value)%></b> Modelo <b><%=HTMLEncode(rs.Fields("Ap_Modelo").Value)%></b></TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#CCCCCC" class="Texbuche">
            <TD height="24" colspan="2"  ALIGN=center bgcolor="#F0F0F0">Nota del Evento</TD>
          </TR>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24" colspan="2"  ALIGN=center bordercolor="#000000" bgcolor="#CCCCCC"><textarea name="textarea" cols="80" rows="3" readonly="readonly" class="textbox"><%=HTMLEncode(rs.Fields("Eve_Nota").Value)%></textarea>            </TD>
          </TR>
          <%
rs.MoveNext
loop%>
        </TBODY>
    </TABLE>   </td>
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