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
<!-- calendar stylesheet -->
<link rel="stylesheet" type="text/css" media="all" href="jscalendar/calendar-blue2.css"/>
<!-- main calendar program -->
  <script type="text/javascript" src="jscalendar/calendar.js"></script>

  <!-- language for the calendar -->
  <script type="text/javascript" src="jscalendar/lang/calendar-es2.js"></script>
  <script type="text/javascript" src="jscalendar/calendar-setup.js"></script>
<script language="JavaScript">

function ValidateRequired(theField)
{
  if (theField.value == "")
  {
	sMsg = "es obligatoria";
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
	if (!ValidateRequired(document.DataNoCon.f_proximo) && document.DataNoCon.lst_Estados.value != 5)
	{
		alert('La fecha para el proximo contacto ' + sMsg );
		document.DataNoCon.f_proximo.focus();
		return (false);
	}
	
	if (!ValidateLen(document.DataNoCon.txt_nota,0,200))
		{
			alert('La Nota del evento ' + sMsg);
			document.DataNoCon.txt_nota.focus();
			return (false);
		}
	return (true)
}
</script>
</HEAD>
<BODY bgcolor="#999999">
<%
clienteid=request("clienid")

    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
%>

<table width="100%"  border="0" cellspacing="0" cellpadding="5">

  <tr class="SloganTrebu">
    <td height="30" align="left"><div align="center" class="Titulo">No Contacto para el <%=clienteid%> </div></td>
  </tr>
<tr class="Msjerror">

  <td align="center">&nbsp;</td>
</tr>

  <tr class="Texbuche">  </tr>
    <tr class="Texbuche">
    <td height="33"><form action="carganocontacto.asp?clienid=<%=clienteid%>" method="post" name="DataNoCon" id="DataNoCon" onSubmit="return validar_form();">
      <div align="center" class="Texbuche"><span class="texto">
        </span>
        <table width="400" border="0" cellspacing="3" cellpadding="0">
          <tr>
            <td><div align="right"><span class="texto">Resultado: </span></div></td>
            <td><div align="left">
              <%
Set rsX = Server.CreateObject("ADODB.Recordset")
sQuery = "SELECT * FROM Estados WHERE Est_Perfil = 'Operador' AND Est_Camino=False"
rsX.Open sQuery, conn, 3 ,3
If rsX.EOF Then
    Response.Write "&nbsp;No hay estados.<BR>"
Else
    Response.Write "<SELECT NAME='lst_Estados' class='textbox'>"
    Do Until rsX.EOF
        Response.Write "<OPTION VALUE=""" & rsX("ID_Estados") & _
            """>" & rsX("Est_Nombre") & "</OPTION>"
        rsX.MoveNext
    Loop
    Response.Write "</SELECT>"
End If
rsX.Close
Set rsX = Nothing
%>
            </div></td>
            </tr>
          <tr>
            <td><div align="right">Nota: </div></td>
            <td><span class="texto">
              <textarea name="txt_nota" cols="40" rows="4" class="textbox" id="txt_nota"></textarea>
            </span></td>
            </tr>
          <tr>
            <td><div align="right">Pr&oacute;ximo llamado: </div></td>
            <td><input name="f_proximo" type="text" class="textbox" id="f_proximo" size="20" readonly="true">
		<script type="text/javascript">
    Calendar.setup({
        inputField     :    "f_proximo",     // id of the input field
        ifFormat       :    "%d/%m/%Y %k:%M", // formatos "%d/%m/%Y" y/o "%d/%m/%Y %k:%M:%S"
     	button		   :    "f_proximo",  // trigger for the calendar (button ID)
        align          :    "CR",           // alignment (defaults to "Bl")
        weekNumbers    :    false,
		firstDay       :    0,
		showsTime      :    true,
		timeFormat     :    "24",
		cache          :    false,
		singleClick    :    false
    });
            </script></td>
            </tr>
        </table>
        <span class="texto">
        <input name="button32" type="reset" class="btn" id="button32" value="Limpiar">
        <input name="button3" type="submit" class="btn" id="button3" value="Continuar">
      </span></div>
    </form></td>
  </tr>
  <tr class="Texbuche">
    <td height="33"><div align="center">
      <%
conn.close
set conn=nothing
Response.Write(piepagina)
%>
    </div></td>
  </tr>
</table>
</BODY>
</HTML>