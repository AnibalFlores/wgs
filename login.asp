<!--#include file="Connections/database.asp" --><%
'//////////// Receive Data //////////
usr=replace(request("usr"),"'","''")
pwd=replace(request("pwd"),"'","''")
application("wgsANM_connection")= MM_WGS_STRING

if len(usr)>0 and len(pwd)>0 then
	set conn=server.createobject("ADODB.Connection")
	'//conn.Mode=adModeReadWrite
	conn.open MM_WGS_STRING

	'/// Check if there's a Default Administrator, otherwise create one
	psql="select * from Usuarios where Usr_Nivel='Administrador'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open psql, conn, adOpenKeyset, adLockPessimistic
	if rs.eof then
		rs.addnew
		rs("Usr_Nombres")="Señor"
		rs("Usr_Apellidos")="Administrador"
		rs("Usr_Usuario")="admin"
		rs("Usr_Pass")="admin"
		rs("Usr_email")="admin@wgs.com.ar"
		rs("Usr_Observacion")="Administrador por Defecto"
		rs("Usr_Nivel")="Administrador"
		rs.update
	end if
	rs.close
	set rs=nothing

	'/// Check Entered username and password
	psql="select * from Usuarios where Usr_Usuario='"&usr&"' and Usr_Pass='"&pwd&"'"
	' // set rs=conn.execute(psql)
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open psql, conn, adOpenKeyset, adLockPessimistic
	if not(rs.eof) then
		session("wgsANM_usr")=rs("ID_Usuarios")
		session("wgsANM_lvl")=rs("Usr_Nivel")
		if rs("Usr_Estado")<>"Logueado" then
	    	rs("Usr_Log")=Now()
			rs("Usr_Passtime")=Now()
			rs("Usr_Estado")="Logueado"
	    	rs.update
	  	end if
	else
		message="Usuario y/o contraseña invalidos"
	end if
	rs.close
	set rs=nothing
	conn.close
    set conn=nothing
	if session("wgsANM_usr")<>"" and message="" then response.redirect "menu.asp"
end if
session("wgsANM_usr")=""
session("wgsANM_lvl")=""
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Style-Type" content="text/css">
<meta name="DESCRIPTION" content="Administrador de Contactos">
<meta name="AUTHOR" content="Anibal Flores">
<meta name="COPYRIGHT" content="Septiembre 2006">
<script language="javascript" type="text/javascript">
<!--
var isNS=document.layers?true:false;
var isIE=(document.all!=null)||(navigator.userAgent.indexOf('MSIE')!=-1);
var isDom2=document.getElementById;
var fVers=parseFloat(navigator.appVersion);
if ((isNS && fVers<4)||(isIE && fVers<4))
    alert("Su Navegador es muy viejo. Por favor actualizelo si desea ver esta pagina correctamente.");
//-->
</script>
<title>WG Sistemas</title>
<link rel="stylesheet" type="text/css" href="wgs.css">
</head>
<body text="#000000" bgcolor="#999999" link="#0000FF" alink="#FF0000" vlink="#800080">
<center>
<form name="form1" method="post" action="login.asp">
  <table width="296" border="0" cellpadding="0" cellspacing="0" class="SloganTrebu">
    <tr class="SloganTrebu">
      <td width="2" height="36">&nbsp;</td>
      <td colspan="2" align="center" valign="middle"><div align="center"><span class="Titulo">WG Sistemas</span><br>
          <span class="SloganTrebu">Ingreso al Administrador de Contactos        </span></div>
      </label></td>
      <td width="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="22">&nbsp;</td>
      <td colspan="2" bgcolor="#CCCCCC" class="Msjerror"><div align="center"><%=message%></div></td>
      <td width="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="27">&nbsp;</td>
      <td width="80" bgcolor="#CCCCCC" class="trebuche"><div align="right" class="Texbuche">Usuario:</div></td>
      <td width="171" bgcolor="#CCCCCC">
        <div align="left">
          <input name="usr" type="text" class="textbox" id="usr" size="28">
        </div></td>
      <td width="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="27">&nbsp;</td>
      <td width="80" bgcolor="#CCCCCC" class="trebuche"><div align="right" class="Texbuche">Contraseña:</div></td>
      <td bgcolor="#CCCCCC">
        <div align="left">
          <input name="pwd" type="password" class="textbox" id="pwd" size="28" autocomplete="off">
        </div></td>
      <td width="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="28">&nbsp;</td>
      <td colspan="2" bgcolor="#CCCCCC">
        <div align="center">
          <input name="Submit" type="submit" class="btn" value="Login">
        </div></td>
      <td width="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="37">&nbsp;</td>
      <td colspan="2" class="SloganTrebu"><div align="center"><%=piepagina%></div></td>
      <td width="2">&nbsp;</td>
    </tr>
  </table>
  <label></label>
  <p>
    <label></label>
    <label></label></p>
  </form>
<%
    set conn=server.createobject("ADODB.Connection")
	conn.open MM_WGS_STRING
	sql = "SELECT * FROM Novedades WHERE Novedades.Nov_Caduca>Now()"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, adOpenStatic, adLockOptimistic
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="5">
  
  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">Novedades Importantes </div></td>
  </tr>
<tr class="Msjerror">

<td align="center">
<%If rs.EOF Then
   Response.Write(sindatos)
Else
%>
<TABLE width="90%" BORDER=0 cellpadding="4" CELLSPACING=1 BGCOLOR=#336699 class="Texbuche">
        <THEAD>
          <TR bordercolor="#CCCCCC" class="SloganTrebu">
            <TH width="14%" height="21" >Publicada</TH>
            <TH width="86%" height="21" ><div align="center">Info</div></TH>
            </TR>
        </THEAD>
        <TBODY>
          <%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof
 %>
          <TR bordercolor="#CCCCCC" bgcolor="#FFFFFF" class="Texbuche">
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Nov_Vigencia").Value)%>            </TD>
            <TD height="24"  ALIGN=center><%=Server.HTMLEncode(rs.Fields("Nov_Info").Value)%>			</TD>
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
'Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
</center>
</BODY>
</HTML>