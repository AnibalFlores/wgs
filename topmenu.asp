<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
    sql = "SELECT Usuarios.ID_Usuarios, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Filtros.fil_Nombre FROM Filtros INNER JOIN Usuarios ON Filtros.ID_Filtros = Usuarios.Usr_Base WHERE (((Usuarios.ID_Usuarios)=" & session("wgsANM_usr") & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
UsrNombre =rs("Usr_Apellidos").value & ", " & rs("Usr_Nombres").value
UsrBase = rs("fil_Nombre").value
lvl=session("wgsANM_lvl")
%>
<html>
<head>
<title>WG Sistemas Administrador de Contactos</title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=charset%>">
<link href="wgs.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF">
<form name="form1" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#e6e6e6">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="75">
          <tr> 
            <td colspan="2" class="SloganTrebu"><div align="center">
              <p><a href="http://www.wgsistemas.com.ar/" target="_blank" class="Titulo">WG Sistemas</a></p>
              </div></td>
            <td width="59%" rowspan="3" align="center" bgcolor="#999999">
              <input name="Buscar" type="button" class="btn" onClick="javascript:parent.main.location.href='listar.asp'" value="Clientes"main"'">
              <input name="Buscar" type="button" class="btn" id="Buscar" onClick="javascript:parent.main.location.href='buscarclientes.asp'" value="Buscar"main"'">
              <input name="Productos" type="button" class="btn" onClick="javascript:parent.main.location.href='productos.asp'" value="Productos"main"'">
              <%if lvl<>"Operador" then%>
              <input name="Eventos" type="button" class="btn" onClick="javascript:parent.main.location.href='referido.asp'" value="Referidos"main"'">
              <input name="Usuarios" type="button" class="btn" onClick="javascript:parent.main.location.href='verusuarios.asp'" value="Usuarios"main"'">
			  <%end if%><%if lvl="Administrador" then%>
              <input name="Configurar" type="button" class="btn" onClick="javascript:parent.main.location.href='configurar.asp'" value="Configurar"main"'">
              <%end if%>
			  <input name="Mi Perfil" type="button" class="btn" onClick="javascript:parent.main.location.href='miperfil.asp'" value="Mi Perfil"main"'">
			  <input name="Info" type="button" class="btn" onClick="window.open('info.html')" value="Info"main"'">
            <input name="Salir" type="button" class="btn" id="Salir" onClick="javascript:parent.main.location.href='logout.asp'" value="Salir">            </td>
          </tr>
          <tr>
            <td width="21%" class="SloganTrebu"><div align="left"><%=lvl%>: <%=UsrNombre%> </div></td>
            <td width="20%" class="SloganTrebu"><div align="left"></div></td>
          </tr>
          <tr>
            <td class="SloganTrebu"><div align="left">Logueados: <%= Application("Logueados")%></div></td>
            <td class="SloganTrebu"><div align="left">Base: <%=UsrBase%></div></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="2" bgcolor="#999999"></td>
    </tr>
    <tr> 
      
    <td bgcolor="#6a6a6a" height="9"><table width="61%" border="0" cellspacing="0" cellpadding="0" align="right">
        <tr>
          <td align="center"></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td bgcolor="#353535" height="2" align="left"></td>
    </tr>
  </table>
  </form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>

