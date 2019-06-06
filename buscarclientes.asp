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
<style type="text/css">
<!--
.Estilo1 {font-size: 12px}
-->
</style>
<script language="JavaScript">
var sNumeric     = "1234567890";
function ValidateString(theField, checkOK)
{
  var checkStr = theField.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
	sMsg = "debe poseer caracteres válidos";
    return (false);
  }
  else
  {
    return (true);
  }
}

function ValidateRequired(theField)
{
  if (theField.value == "")
  {
	sMsg = "es obligatorio";
    return (false);
  }
  else
  {
    return (true);
  }
}

function ValidateLen(theField, nMinLen, nMaxLen)
{
  var checkStr = theField.value;

  if (checkStr.length < nMinLen)
  {
	sMsg = "debe poseer al menos " + nMinLen + " caracteres";
    return (false);
  }
  else
  {
	  if (checkStr.length > nMaxLen)
	  {
		sMsg = "debe poseer menos de " + nMaxLen + " caracteres";
		return (false);
	  }
	  else
	  {
		return (true);
	  }
  }
}

function validar_form()
{
	if (!ValidateRequired(document.Filtro.txt_Nro))
		{
			alert('El campo Número de Cliente ' + sMsg );
			document.Filtro.txt_Nro.focus();
			return (false);
		}
		
	if (!ValidateString(document.Filtro.txt_Nro,sNumeric))
		{
			alert('El campo Número de Cliente ' + sMsg );
			document.Filtro.txt_Nro.focus();
			return (false);
		}
		
	if (!ValidateLen(document.Filtro.txt_Nro,10,10))
		{
			alert('El campo Número de Cliente ' + sMsg );
			document.Filtro.txt_Nro.focus();
			return (false);
		}
			
	return (true)
}
</script>
</HEAD>
<BODY bgcolor="#999999">
<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">Buscar Clientes </div></td>
  </tr>
<tr class="Msjerror">

  <td align="center"><p>EN CONSTRUCCION </p>
    <form action="listaresul.asp" method="post" name="Filtro" id="Filtro" onSubmit="return validar_form();">
      <table width="90%" border="0" cellspacing="0" cellpadding="5">
        <tr bgcolor="#CCCCCC" class="Titulo">
          <td colspan="4"><div align="center">
            <p>Filtros de consulta</p>
            <p class="Estilo1"> (por ahora solo funciona buscar por n&uacute;mero) </p>
          </div></td>
          </tr>
        <tr bgcolor="#F0F0F0">
          <td> <div align="right">Numero:            </div></td>
          <td><input name="txt_Nro" type="text" class="textbox" id="txt_Nro" size="12"></td>
          <td><div align="right">Apellido:            </div></td>
          <td><input name="txt_Apellido" type="text" class="textbox" id="txt_Apellido" size="40"></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td><div align="right">Estado: </div></td>
          <td><select name="lst_Estado" class="textbox" id="lst_Estado">
            <option value="*">Todos</option>
                                        </select></td>
          <td><div align="right">Localidad:            </div></td>
          <td><select name="lst_Localidad" class="textbox" id="lst_Localidad">
            <option value="*">Todas</option>
            </select>            </td>
        </tr>
        <tr bgcolor="#F0F0F0">
          <td><div align="right">Empresa:</div></td>
          <td><select name="lst_Empresa" class="textbox" id="lst_Empresa">
            <option value="*">Todas</option>
                                        </select></td>
          <td><div align="right">Asesor:</div></td>
          <td><select name="lst_Asesor" class="textbox" id="lst_Asesor">
            <option value="*">Todos</option>
                              </select></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td colspan="4"><div align="center">
            <input name="button2" type=reset class="btn" value="Limpiar">            
            <input name="button" type=submit class="btn" value="Buscar">
          </div></td>
          </tr>
      </table>
        </form>
    </td>
</tr>
  <tr class="Texbuche">
    <td height="33"><div align="center"></p>
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