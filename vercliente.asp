<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
%>
<html>
<head>
	<title>Administrar Clientes</title>
  <script language="JavaScript" type="text/javascript">
		function manejoVisualizacion(nombre)
		{
			objRow=eval('capa' + nombre)
			objImg=eval('flecha' + nombre)
			if (objRow(0).style.display=='none')
						{
				objRow(0).style.display='list-item'
				objRow(1).style.display='list-item'
				objImg.src="gif/flecha2.gif"

			}
			else
			{
				objRow(0).style.display='none'
				objRow(1).style.display='none'
				objImg.src="gif/flecha.gif"
			}
		}


	</script>


<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="wgs.css" rel="stylesheet" type="text/css">

</head>

<body bgcolor="#999999" text="#000000" leftmargin="0" topmargin="20">
<%
'// WG Sistemas (c) 2006//'
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
sql = "SELECT Clientes.ID_Clientes, Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_TipoDNI, Clientes.Cli_DNInro, Clientes.Cli_Nacimiento, Clientes.Cli_Genero, Clientes.Cli_Email, Status.Sta_Nombre, Empresas.Emp_Corto, Clientes.Cli_Contacto, Clientes.Cli_Fechaultimo, Clientes.Cli_Fechaprox, Clientes.Cli_Comentario, Clientes.Cli_Telefono2, Clientes.Cli_CPNuevo, Clientes.Cli_Eventos, Clientes.Cli_Valor, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Localidades.Loc_Localidad, Localidades.Loc_Provincia,Localidades.Loc_Prefijo, Clientes.Cli_Inicio, Clientes.Cli_Calle, Clientes.Cli_CalleNro, Clientes.Cli_Piso, Clientes.Cli_Depto, Clientes.Cli_Ampli FROM Localidades INNER JOIN ((Empresas INNER JOIN (Status INNER JOIN Clientes ON Status.ID_Status = Clientes.Cli_Status) ON Empresas.ID_Empresas = Clientes.Cli_Empresa) INNER JOIN Usuarios ON Clientes.Cli_Usuario = Usuarios.ID_Usuarios) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE (((Clientes.Cli_Numero)="& clienteid & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, adOpenStatic, adLockOptimistic
	activo = Isdate(rs.Fields("Cli_Inicio").Value)
	
	if activo then
	 msje = "Cliente inicializado el " + FormatMediumDate(rs.Fields("Cli_Inicio").Value)
	Else
	 msje = "Datos del cliente aun no relevados"
    End If
	A_Discar = NroDiscar(HTMLEncode(rs.Fields("Cli_Numero").Value),HTMLEncode(rs.Fields("Loc_Prefijo")))
%>
<table width='90%' border='0' cellspacing='2' cellpadding='4' height='30' align='center'>
		  	<tr bgcolor='#FFFFFF' class="Titulo">
				<td width="50%">Linea : <b><%=A_Discar%></b></td>
				<td width="50%" align='left' class='Titulo'>Estado : <b><%=HTMLEncode(rs.Fields("Sta_Nombre").Value)%></b></td>
			</tr>
			<tr bgcolor='#FFFFFF' class="Texbuche">
				<td class='texto' align='left'><%=msje%></td>
				<td class='texto' align='left'>Empresa: <%=HTMLEncode(rs.Fields("Emp_Corto").Value)%></td>
		  </tr>
			<tr bgcolor='#FFFFFF' class="SloganTrebu">
			  <td class='texto' align='left'>Ultimo contacto  el <%=FormatMediumDate(rs.Fields("Cli_Fechaultimo").Value)%> por <%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%>&nbsp;<%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%></td>
			  <td class='texto' align='left'>Proximo contacto agendado para el <%=HTMLEncode(rs.Fields("Cli_Fechaprox").Value)%></td>
		  </tr>
</table>

	<!--*********************************************** -->
<% If activo then %>
	<!-- Se traen el historial de la solicitud y se imprime -->
	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
  <tr bgcolor='#FFFFFF'><td width='10' bgcolor="#FFFFFF"><a href='#' onclick=manejoVisualizacion('1')><img src='gif/flecha2.gif' alt="Ver Detalles" name="flecha1" width="18" height="18" border='0' id=flecha1></a></td>
	<td class='Titulo'><div align="center">Datos Personales</div></td>
	</tr><tr><td id='capa1' style='display:list-item'>&nbsp;</td>
	<td colspan='4' style='display:list-item' id='capa1' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
<tr bgcolor="#CCCCCC" class="trebuche">
<td width="50%" align="left" class="Texbuche"><div align="right">Apellidos: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Apellido").Value)%></b></span></div></td>
<td width="50%" align="left" bgcolor="#CCCCCC" class="Texbuche"><div align="left">Nombres: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Nombre").Value)%></b></span></div></td>
</tr>
<tr bgcolor="#F0F0F0" class="trebuche">
<td width="50%" align="left" class="Texbuche"><div align="right">Tipo Documento: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_TipoDNI").Value)%></b></span></div></td>
<td width="50%" align="left" bgcolor="#F0F0F0" class="Texbuche"><div align="left">N&uacute;mero:<span class="texto">
  <b><%=HTMLEncode(rs.Fields("Cli_DniNro").Value)%></b></span></div></td>
</tr>
<tr align="left" bgcolor="#CCCCCC" class="trebuche">
  <td width="50%" class="Texbuche"><div align="right">
      G&eacute;nero: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Genero").Value)%></b></span></div></td>
  <td width="50%" bgcolor="#CCCCCC" class="texto">
    <div align="left">
        <span class="Texbuche">Nacimiento: <b><%=HTMLEncode(rs.Fields("Cli_Nacimiento").Value)%></b> </span>
  	  </div>  </td>
</tr>
</table></td></tr></table>

	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
      <tr bgcolor='#FFFFFF'>
        <td width='10' bgcolor="#FFFFFF"><a href='#' onclick=manejoVisualizacion('2')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha2" width="18" height="18" border='0' id=flecha2></a></td>
        <td class='Titulo'><div align="center">Domicilio</div></td>
      </tr>
	  <tr>
	    <td id='capa2' style='display:none'>&nbsp;</td>
	    <td colspan='4' style='display:none' id='capa2' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td width="50%" align="left" class="Texbuche"><div align="right">Calle: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Calle").Value)%></b></span></div></td>
              <td width="50%" align="left" class="texto"><div align="left" class="Texbuche">N&uacute;mero: <b><%=HTMLEncode(rs.Fields("Cli_CalleNro").Value)%></b> Piso:              
                      <b><%=HTMLEncode(rs.Fields("Cli_Piso").Value)%></b> Depto.:              
                      <b><%=HTMLEncode(rs.Fields("Cli_Depto").Value)%></b></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td colspan="2" align="left" class="Texbuche"><div align="center">Ampliaci&oacute;n: <span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Ampli").Value)%></b></span></div>              </td>
            </tr>
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td width="50%" align="left" class="Texbuche"><div align="right">Localidad: <b><%=HTMLEncode(rs.Fields("Loc_Localidad").Value)%></b></div></td>
              <td width="50%" align="left" class="texto"><div align="left"><span class="Texbuche">Provincia: <b><%=HTMLEncode(rs.Fields("Loc_Provincia").Value)%></b> </span></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td colspan="2" align="left" class="Texbuche"><div align="center">C&oacute;digo Postal Nuevo: <b><%=HTMLEncode(rs.Fields("Cli_CPNuevo").Value)%></b></div></td>
            </tr>
        </table></td>
      </tr>
</table>
	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
      <tr bgcolor='#FFFFFF'>
        <td width='10' bgcolor="#FFFFFF"><a href='#' onclick=manejoVisualizacion('3')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha3" width="18" height="18" border='0' id=flecha3></a></td>
        <td class='Titulo'><div align="center">Datos de Contacto</div></td>
      </tr>
	  <tr>
	    <td id='capa3' style='display:none'>&nbsp;</td>
	    <td colspan='4' style='display:none' id='capa3' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Persona de Contacto: <b><%=HTMLEncode(rs.Fields("Cli_Contacto").Value)%></b> </div>
              <div align="left" class="Texbuche"></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Correo electronico:<b> <%=HTMLEncode(rs.Fields("Cli_Email").Value)%></b></div>
              <div align="left"></div></td>
            </tr>

            <tr bgcolor="#CCCCCC" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Telefono Fijo: <b><%=HTMLEncode(rs.Fields("Cli_Telefono2").Value)%></b> </div>
              <div align="left"></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Observaciones</div>
                <div align="center">
                  <textarea name="textarea" cols="80" rows="4" class="textbox"><%=HTMLEncode(rs.Fields("Cli_Comentario").Value)%></textarea>
              </div></td>
            </tr>
        </table></td>
      </tr>
</table>
	<% End If %>
	  <table width="90%" border="0" align="center" cellpadding="4" cellspacing="0" class="Texbuche">
        <tr>
          <td><div align="center"><span class="subtitulorep">
            <input name="button" type=button class="btn" onClick="history.go(-1)" value="Volver">
          <span class="texto">
          <input name="button4" type="button" class="btn" id=button4 onClick="javascript:location.href='verhistorico.asp?clienid=<%=Clienteid%>'" value="Ver Historial">
          <input name="button3" type="button" class="btn2" id="button3" onClick="javascript:location.href='nuevoevento.asp?clienid=<%=Clienteid%>&Prefix=<%=A_Discar%>'" value="Llamar">
          </span></span></div></td>
        </tr>
        <tr>
          <td bgcolor="#CCCCCC"><div align="center">Recuerde: Al llamar al cliente, este quedará seleccionado de forma exclusiva por Ud. hasta cerrar el contacto. </div></td>
        </tr>
</table>
      <div align="center" class="Texbuche">
        <%
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write(piepagina)
%>
      </div>
</body>
</html>
