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
<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo"> Mi Perfil </div></td>
  </tr>
<tr class="Msjerror">

  <td align="center">EN CONSTRUCCION </td>
</tr>
  <tr class="Texbuche">
    <td height="33"><div align="center">
      <input type=button class="btn" onClick="history.go(-1)" value="Volver">
      </p>
    </div>
    <div align="center">
<%
Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
</BODY>
</HTML>