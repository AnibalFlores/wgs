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
'// WG Sistemas (c) 2006
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
'Verifico si la consulta actual es la correcta sino la creo
If IsObject(Session("VerCliente_rs")) Then
    Set rs = Session("VerCliente_rs")
Else
sql="SELECT * From Productos WHERE Prod_Vigencia=true"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
End If
%>
<table width='90%' border='0' cellspacing='2' cellpadding='4' height='30' align='center'>
		  	<tr bgcolor='#FFFFFF' class="Titulo">
				<td><div align="center">Productos Vigentes </div></td>
			</tr>
			<tr bgcolor='#FFFFFF' class="Texbuche">
				<td align='left' class='Texbuche'>Los precios detallados no incluyen IVA. </td>
		  </tr>
</table>

	<!--*********************************************** -->
<%If rs.EOF Then
   Response.Write(sindatos)
Else
On Error Resume Next
rs.MoveFirst
count = 1
do while Not rs.eof 
%>
	<!-- Se traen el historial de la solicitud y se imprime -->
	<table width='700' height="25" border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#999999">
  <tr bgcolor='#FFFFFF'><td width='3%' height="25"><details><summary></td>
	<td width="24%" height="25" bgcolor="#FFFFFF" class='SloganTrebu'> Código:<span class="texto"><%=HTMLEncode(rs.Fields("Prod_Codigo").Value)%></span></td>
	<td width="73%" height="25" bgcolor="#FFFFFF" class='Titulo'><span class="texto"><%=HTMLEncode(rs.Fields("Prod_Nombre").Value)%></span></td>
  </tr></summary><tr><td width="18" id='capa<%=count%>' style='display:none'>&nbsp;</td>
	<td colspan='5' style='display:none' id='capa<%=count%>' align='left'><table width="660" border="0" align="center" cellpadding="5" cellspacing="2">
<tr class="trebuche">
<td colspan="4" align="left" bgcolor="#336699" class="Texbuche"><div align="center" class="SloganTrebu">Descripción</div>  </td>
</tr>
<tr class="trebuche">
  <td colspan="4" align="left" bgcolor="#FFFFFF" class="Texbuche"><span class="texto"><%=HTMLEncode(rs.Fields("Prod_Descrip").Value)%></span></td>
</tr>
<tr align="left" bgcolor="#CCCCCC" class="trebuche">
  <td width="25%" bgcolor="#CCCCCC" class="Texbuche">
      <div align="right">Alta:</div>
      <div align="left"></div>      <div align="right"></div>      <div align="left"></div>
      <div align="left"></div></td>
  <td width="23%" bgcolor="#CCCCCC" class="Texbuche"><div align="left"><span class="texto"><%=HTMLEncode(rs.Fields("Prod_Alta").Value)%></span></div></td>
  <td width="23%" bgcolor="#336699" class="Texbuche"><div align="right"><span class="SloganTrebu">Precio: $</span></div></td>
  <td width="29%" bgcolor="#336699" class="Texbuche"><div align="left"><span class="SloganTrebu"><span class="texto"><%=HTMLEncode(rs.Fields("Prod_Precio").Value)%></span></span></div></td>
</tr>
<tr align="left" bgcolor="#CCCCCC">
  <td bgcolor="#CCCCCC" ><div align="right" class="Texbuche">Baja: </div></td>
  <td bgcolor="#CCCCCC" class="Texbuche"><div align="left" class="Texbuche"><%=HTMLEncode(rs.Fields("Prod_Baja").Value)%></div></td>
  <td bgcolor="#336699" class="Texbuche"><div align="right" class="SloganTrebu">Minutos Libres:</div></td>
  <td bgcolor="#336699"><div align="left" class="Texbuche"><span class="SloganTrebu"><span class="texto"><%=HTMLEncode(rs.Fields("Prod_MinLib").Value)%></span></span></div></td>
</tr>
<tr align="left" bgcolor="#CCCCCC">
  <td bgcolor="#CCCCCC" ><div align="right" class="Texbuche">Rango:</div></td>
  <td bgcolor="#CCCCCC" ><div align="left" class="Texbuche"><span><%=HTMLEncode(rs.Fields("Prod_Rango").Value)%></span></div></td>
  <td bgcolor="#CCCCCC" class="SloganTrebu"><div align="right">Valor por minuto: $ </div></td>
  <td bgcolor="#CCCCCC" class="SloganTrebu"><span class="SloganTrebu"><%=HTMLEncode(rs.Fields("Prod_Valorminuto").Value)%></span></td>
</tr>
</table></td></tr></table></details>
<%
rs.MoveNext
count= count+1
loop
End If%>

	<div align="center">
	  <!-- Se traen los posibles caminos a seguir --><br>
</div>
	<center>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write(piepagina)
%>
	</center>

</body>
</html>
