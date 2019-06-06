<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
%>
<html>
<head>
	<title>Administrar Clientes</title>
  <script language="javascript">
		function manejoVisualizacion(nombre)
		{
			objRow=eval('capa' + nombre)
			objImg=eval('flecha' + nombre)
			if (objRow(0).style.display=='none') and (activo)
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
'// WG Sistemas (c) 2006
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
'Verifico si la consulta actual es la correcta sino la creo
If IsObject(Session("VerCliente_rs")) Then
    Set rs = Session("VerCliente_rs")
Else
sql = "SELECT Clientes.ID_Clientes, Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_TipoDNI, Clientes.Cli_DNInro, Clientes.Cli_Nacimiento, Clientes.Cli_Genero, Clientes.Cli_Email, Estados.Est_Nombre, Empresas.Emp_Corto, Clientes.Cli_Contacto, Clientes.Cli_Fechaultimo, Clientes.Cli_Fechaprox, Clientes.Cli_Comentario, Clientes.Cli_Telefono2, Clientes.Cli_CPNuevo, Clientes.Cli_Eventos, Clientes.Cli_Valor, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Localidades.Loc_Localidad, Localidades.Loc_Provincia, Clientes.Cli_Inicio, Clientes.Cli_Calle, Clientes.Cli_CalleNro, Clientes.Cli_Piso, Clientes.Cli_Depto, Clientes.Cli_Ampli FROM Localidades INNER JOIN ((Empresas INNER JOIN (Estados INNER JOIN Clientes ON Estados.ID_Estados = Clientes.Cli_Estado) ON Empresas.ID_Empresas = Clientes.Cli_Empresa) INNER JOIN Usuarios ON Clientes.Cli_Usuario = Usuarios.ID_Usuarios) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE (((Clientes.Cli_Numero)="& clienteid & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
	activo = Isdate(rs.Fields("Cli_Inicio").Value)
	
	if activo then
	 msje = "Cliente inicializado el " + FormatMediumDate(rs.Fields("Cli_Inicio").Value)
	Else
	 msje = "Datos del cliente aun no relevados"
    End If
End If
%>
<table width='90%' border='0' cellspacing='2' cellpadding='4' height='30' align='center'>
		  	<tr bgcolor='#FFFFFF' class="Titulo">
				<td width="50%">Linea : <b><%=HTMLEncode(rs.Fields("Cli_Numero").Value)%></b></td>
				<td width="50%" align='left' class='Titulo'>Estado : <b><%=clienteid%></b></td>
			</tr>
			<tr bgcolor='#FFFFFF' class="Texbuche">
				<td class='texto' align='left'><%=msje%></td>
				<td class='texto' align='left'>&nbsp;</td>
		  </tr>
			<tr bgcolor='#FFFFFF' class="SloganTrebu">
			  <td class='texto' align='left'>&nbsp;</td>
			  <td class='texto' align='left'>&nbsp;</td>
		  </tr>
</table>

	<!--*********************************************** -->

	<!-- Se traen el historial de la solicitud y se imprime -->
	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
  <tr bgcolor='#FFFFFF'><td width='10' bgcolor="#3399FF"><a href='#' onclick=manejoVisualizacion('1')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha1" width="18" height="18" border='0' id=flecha1></a></td>
	<td class='Titulo'>Datos Personales:</td>
	</tr><tr><td id='capa1' style='display:none'>&nbsp;</td>
	<td colspan='4' style='display:none' id='capa1' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
<tr bgcolor="#CCCCCC" class="trebuche">
<td width="50%" align="left" class="Texbuche"><div align="right">Apellidos: <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Apellido").Value)%></span></div></td>
<td width="50%" align="left" bgcolor="#CCCCCC" class="Texbuche"><div align="left">Nombres:<span class="texto"><%=HTMLEncode(rs.Fields("Cli_Nombre").Value)%></span></div></td>
</tr>
<tr bgcolor="#F0F0F0" class="trebuche">
<td width="50%" align="left" class="Texbuche"><div align="right">Tipo Documento:<span class="texto"><%=HTMLEncode(rs.Fields("Cli_TipoDNI").Value)%></span></div></td>
<td width="50%" align="left" bgcolor="#F0F0F0" class="Texbuche"><div align="left">N&uacute;mero:<span class="texto">
  <%=HTMLEncode(rs.Fields("Cli_DniNro").Value)%></span></div></td>
</tr>
<tr align="left" bgcolor="#CCCCCC" class="trebuche">
  <td width="50%" class="Texbuche"><div align="right">
      G&eacute;nero: <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Genero").Value)%></span></div></td>
  <td width="50%" bgcolor="#CCCCCC" class="texto">
    <div align="left">
        <span class="Texbuche">Nacimiento: <%=HTMLEncode(rs.Fields("Cli_Nacimiento").Value)%> </span>
  	  </div>  </td>
</tr>
</table></td></tr></table>

	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
      <tr bgcolor='#FFFFFF'>
        <td width='10' bgcolor="#3399FF"><a href='#' onclick=manejoVisualizacion('2')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha2" width="18" height="18" border='0' id=flecha2></a></td>
        <td class='Titulo'>Domicilio:</td>
      </tr>
	  <tr>
	    <td id='capa2' style='display:none'>&nbsp;</td>
	    <td colspan='4' style='display:none' id='capa2' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td width="50%" align="left" class="Texbuche"><div align="right">Calle: <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Calle").Value)%></span></div></td>
              <td width="50%" align="left" class="texto"><div align="left"><span class="Texbuche">N&uacute;mero: </span>
                <%=HTMLEncode(rs.Fields("Cli_CalleNro").Value)%> <span class="Texbuche">Piso:              </span>
                <%=HTMLEncode(rs.Fields("Cli_Piso").Value)%> <span class="Texbuche">Depto.:              </span>
                <%=HTMLEncode(rs.Fields("Cli_Depto").Value)%></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td colspan="2" align="left" class="Texbuche"><div align="center">Ampliaci&oacute;n: <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Ampli").Value)%></span></div>              </td>
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
        <td width='10' bgcolor="#3399FF"><a href='#' onclick=manejoVisualizacion('3')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha3" width="18" height="18" border='0' id=flecha3></a></td>
        <td class='Titulo'>Datos de Contacto:</td>
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
                  <textarea name="textarea" cols="80" rows="4" readonly="readonly" class="textbox"><%=HTMLEncode(rs.Fields("Cli_Comentario").Value)%></textarea>
              </div></td>
            </tr>
        </table></td>
      </tr>
</table>
	<table width='90%' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
      <tr bgcolor='#FFFFFF'>
        <td width='10' bgcolor="#3399FF"><a href='#' onclick=manejoVisualizacion('4')><img src='gif/flecha.gif' alt="Ver Detalles" name="flecha4" width="18" height="18" border='0' id=flecha4></a></td>
        <td class='Titulo'>Datos Comerciales:</td>
      </tr>
      <tr>
        <td id='capa4' style='display:none'>&nbsp;</td>
        <td colspan='4' style='display:none' id='capa4' align='left'><table width="90%" border="0" align="center" cellpadding="5" cellspacing="0">
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Es cliente de <span class="texto">la empresa: <%=HTMLEncode(rs.Fields("Emp_Corto").Value)%></span></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td align="left" class="Texbuche"><div align="center">Contactado por ultima vez el <span class="texto"><%=FormatMediumDate(rs.Fields("Cli_Fechaultimo").Value)%></span> por <%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%>&nbsp;<%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%></div></td>
            </tr>
            <tr bgcolor="#F0F0F0" class="trebuche">
              <td align="left" bgcolor="#CCCCCC" class="Texbuche"><div align="center">Proximo contacto agendado para el <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Fechaprox").Value)%></span></div></td>
            </tr>
            <tr bgcolor="#CCCCCC" class="trebuche">
              <td align="left" bgcolor="#F0F0F0" class="Texbuche"><div align="center">Cantidad de Contactos: <span class="texto"><%=HTMLEncode(rs.Fields("Cli_Eventos").Value)%></span></div></td>
            </tr>
        </table></td>
      </tr>
</table>
	<div align="center">
	  <!-- Se traen los posibles caminos a seguir -->
	  <span class="subtitulorep">
	  <input name=button1 type="button" class="btn" id=button1 onClick="javascript:location.href='listar.asp'" value="Volver">
	  <%if activo then%>
	  <input name="button2" type="button" class="btn" id=button2 onClick="javascript:location.href='editarcliente.asp?clienid=<%=Clienteid%>'" value="Modificar">
      <%end if%>
	  </span><span class="texto">
	  <input name="button3" type="button" class="btn" id="button3" onClick="javascript:location.href='nuevoevento.asp?clienid=<%=Clienteid%>'" value="Llamar">
	  </span><span class="texto">
	  <input name="button4" type="button" class="btn" id=button4 onClick="javascript:location.href='verhistorico.asp?clienid=<%=Clienteid%>'" value="Ver Historial">
	  </span><br>
</div>
	<center>
<%
rs.close
set rs=nothing
Session("VerCliente_rs")=""
conn.close
set conn=nothing
Response.Write(piepagina)
%>
	</center>

</body>
</html>
