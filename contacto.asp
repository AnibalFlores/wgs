<!--#include file="Connections/database.asp" -->
<%
If session("wgsANM_usr")="" Then
    Response.Redirect "Logout.asp"
End If
Dim rsMain
Dim rsMain_numRows

Set rsMain = Server.CreateObject("ADODB.Recordset")
rsMain.ActiveConnection = MM_WGS_STRING
rsMain.Source = "SELECT *  FROM Marcas ORDER BY Mar_Nombre"
rsMain.CursorType = 0
rsMain.CursorLocation = 2
rsMain.LockType = 1
rsMain.Open()

rsMain_numRows = 0

Dim rsSub
Dim rsSub_numRows

Set rsSub = Server.CreateObject("ADODB.Recordset")
rsSub.ActiveConnection = MM_WGS_STRING
rsSub.Source = "SELECT *  FROM Aparatos WHERE Ap_Vigencia=True ORDER BY Ap_Modelo ASC"
rsSub.CursorType = 0
rsSub.CursorLocation = 2
rsSub.LockType = 1
rsSub.Open()

rsSub_numRows = 0
%>
<html>
<head>
	<title>Editar Cliente</title>
<!-- Dynamic Dependent List box Code for *** VBScript *** Server Model //-->
<script language="JavaScript">
<!--

var arrDynaList = new Array();
var arrDL1 = new Array();

arrDL1[1] = "lst_Marca";		// Name of parent list box
arrDL1[2] = "DatClie";		// Name of form containing parent list box
arrDL1[3] = "lst_Modelo";		// Name of child list box
arrDL1[4] = "DatClie";		// Name of form containing child list box
arrDL1[5] = arrDynaList;	// No need to do anything here
  
<%
Dim txtDynaListRelation, txtDynaListLabel, txtDynaListValue, oDynaListRS

txtDynaListRelation = "Ap_Marca" 	' Name of recordset field relating to parent
txtDynaListLabel = "Ap_Modelo" 			' Name of recordset field for child Item Label
txtDynaListValue = "Id_Aparato" 			' Name of recordset field for child Value
Set oDynaListRS = rsSub						' Name of child list box recordset
  
Dim varDynaList
varDynaList = -1

Dim varMaxWidth
varMaxWidth = "1"

Dim varCheckGroup
varCheckGroup = oDynaListRS.Fields.Item(txtDynaListRelation).Value

Dim varCheckLength
varCheckLength = 0

Dim varMaxLength
varMaxLength = 0

While (NOT oDynaListRS.EOF)

 If (varCheckGroup <> oDynaListRS.Fields.Item(txtDynaListRelation).Value) Then
  If (varCheckLength > varMaxLength) Then
   varMaxLength = varCheckLength
  End If
  varCheckLength = 0
 End If
%>
 arrDynaList[<%=(varDynaList+1)%>] = "<%=(oDynaListRS.Fields.Item(txtDynaListRelation).Value)%>"
 arrDynaList[<%=(varDynaList+2)%>] = "<%=(oDynaListRS.Fields.Item(txtDynaListLabel).Value)%>"
 arrDynaList[<%=(varDynaList+3)%>] = "<%=(oDynaListRS.Fields.Item(txtDynaListValue).Value)%>"
<%
 If (len(oDynaListRS.Fields.Item(txtDynaListLabel).Value) > len(varMaxWidth)) Then
  varMaxWidth = oDynaListRS.Fields.Item(txtDynaListLabel).Value
 End If
 varCheckLength = varCheckLength + 1
 varDynaList = varDynaList + 3
 oDynaListRS.MoveNext()
Wend

If (varCheckLength > varMaxLength) Then
 varMaxLength = varCheckLength
End If
%>
//-->
</script>

<!-- End of object/array definitions, beginning of generic functions -->

<script language="JavaScript">
<!--
function setDynaList(arrDL){

 var oList1 = document.forms[arrDL[2]].elements[arrDL[1]];
 var oList2 = document.forms[arrDL[4]].elements[arrDL[3]];
 var arrList = arrDL[5];
 
 clearDynaList(oList2);
 
 if (oList1.selectedIndex == -1){
  oList1.selectedIndex = 0;
 }

 populateDynaList(oList2, oList1[oList1.selectedIndex].value, arrList);
 return true;
}
 
function clearDynaList(oList){

 for (var i = oList.options.length; i >= 0; i--){
  oList.options[i] = null;
 }
 
 oList.selectedIndex = -1;
}
 
/*This is a modified function from the original MM script. Mick White 
added the first line of oList so there would be an initial selection.
Needed this if there is only 1 child menu item, otherwise, the single
child menu item would be already hihghlighted and you can not select 
it. Also good for validation purposes so you can set the .js 
validation to not allow the first selection.
*/

function populateDynaList(oList, nIndex, aArray){
/*oList[oList.length]= new Option("Seleccione uno");*/
 for (var i = 0; i < aArray.length; i= i + 3){
  if (aArray[i] == nIndex){
   oList.options[oList.options.length] = new Option(aArray[i + 1], aArray[i + 2]);
  }
  //oList.size=oList.length //You need to comment out this line of the function if you use this mod
 }

//A quick mod here, I changed the ==0 to ==1 so that the length 
//takes into account the Please select option from above.
 if (oList.options.length == 0){
  oList.options[oList.options.length] = new Option("Ninguno",0);
 }
 oList.selectedIndex = 0; 
}

//-->
</script>
<script language="JavaScript">
var sAlpha     = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
var sAlphaExt  = "ÁÉÍÓÚÑÄËÏÖÜÂÊÎÔÛ áéíóúñäëïöüâêîôû .";
var sAlphaTE   = "() -";
var sAlphaMail = "@.-";
var sSpace     = " ";
var sAt        = "@";
var sDot       = ".";
var sDash      = "/";
var sSinNum    = "SNsn/";
var sNumeral   = "#";
var sUnderscore  = "_";
var sNumeric     = "1234567890";
var sNumericExt  = ".+-/*";
var sEsimo       ="º";    
var sMsg         = "";
var sHyphen		 = "-";
var sComa	= ",";
var sUnder	= "_";
var sAll  = sAlpha + sAlphaExt + sNumeric + sNumericExt + sAlphaTE + sAlphaMail + sComa + "!¡¿?"
var sAlphaNum = sAlpha + sAlphaExt + sNumeric + sUnder + sComa + "!¡¿?"
var sCalle = sAlpha + sAlphaExt + sNumeric + sAlphaTE + sEsimo + sNumeral+ sComa

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

function ValidateEmail(theField)
{
  var checkOK = sAlpha + sNumeric + sAlphaMail;
  var checkStr = theField.value;
  var allValid = true;
  var ExisteAt = false;

  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    if (ch == "@") ExisteAt = true;
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  
  if ((!allValid) || (!ExisteAt))
  {
	sMsg = "es inválido";
    return (false);
  }
  else
  {
    return (true);
  }
}

function copiarNombre()
{
  if (document.DatClie.checkbox.checked)
  {
	document.DatClie.txt_contacto.value = document.DatClie.txt_Nombres.value + ", " + document.DatClie.txt_Apellidos.value;
	}
  else
  {
    document.DatClie.txt_contacto.value = "";
	}
	return
}

function validar_form()
	{  /* produ = document.DatClie.lst_Producto.value;
		marca = document.DatClie.lst_Marca.value;
		modelo = document.DatClie.lst_Modelo.value;*/		
	    if  (document.DatClie.lst_Estados.value == 1)
		{ pac = true 
		}else
		{ pac = false
		}
		
		if  (document.DatClie.lst_Estados.value == 2)
		{ nega = true 
		}else
		{ nega = false
		}
		
		if  (document.DatClie.lst_Estados.value == 6)
		{ pendi = true 
		}else
		{ pendi = false
		}
		
		if (!ValidateRequired(document.DatClie.txt_Apellidos))
		{
			alert('El campo Apellidos/Razón Social del Cliente ' + sMsg );
			document.DatClie.txt_Apellidos.focus();
			return(false);
		}
		
		if (!ValidateRequired(document.DatClie.txt_Nombres))
		{
			alert('El campo Nombres del Cliente ' + sMsg );
			document.DatClie.txt_Nombres.focus();
			return (false);
		}
		
		if (!ValidateRequired(document.DatClie.txt_NroDNI) && pac)
		{
			alert('El campo Nro. Documento ' + sMsg );
			document.DatClie.txt_NroDNI.focus();
			return (false);
		}
		
		if (!ValidateRequired(document.DatClie.txt_calle) && pac)
		{
			alert('El campo Calle del Cliente ' + sMsg );
			document.DatClie.txt_calle.focus();
			return (false);
		}
		
		if (!ValidateRequired(document.DatClie.txt_calleNro) && pac)
		{
			alert('El campo Altura de la calle del Cliente ' + sMsg );
			document.DatClie.txt_calleNro.focus();
			return (false);
		}
		
		if (!ValidateRequired(document.DatClie.txt_contacto) && pac)
		{
			alert('El campo Persona de contacto ' + sMsg );
			document.DatClie.txt_contacto.focus();
			return (false);
		}
		
		if (!ValidateRequired(document.DatClie.f_Proximo) && !nega)
		{
			alert('El campo Proximo Llamado/Contacto con el Cliente ' + sMsg );
			document.DatClie.f_Proximo.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_Apellidos,sAlpha+sAlphaExt))
		{
			alert('El campo Apellidos del Cliente ' + sMsg );
			document.DatClie.txt_Apellidos.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_Nombres,sAlpha+sAlphaExt))
		{
			alert('El campo Nombres del Cliente ' + sMsg );
			document.DatClie.txt_Nombres.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_NroDNI,sNumeric))
		{
			alert('El campo Nro. documento del Cliente ' + sMsg );
			document.DatClie.txt_NroDNI.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_calle,sCalle))
		{
			alert('El campo calle del Cliente ' + sMsg);
			document.DatClie.txt_calle.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_calleNro,sNumeric+sSinNum))
		{
			alert('El campo Altura de la calle del Cliente ' + sMsg +'. O ingrese S/N');
			document.DatClie.txt_calleNro.focus();
			return (false);
		}
		
		if (!ValidateString(document.DatClie.txt_fijo,sNumeric))
		{
			alert('El campo Nro. Telefono fijo del Cliente ' + sMsg +'. Ej: 3434373048');
			document.DatClie.txt_fijo.focus();
			return (false);
		}

		if (!ValidateString(document.DatClie.txt_contacto,sAlpha+sAlphaExt+sComa))
		{
			alert('El campo Persona de contacto ' + sMsg );
			document.DatClie.txt_contacto.focus();
			return (false);
		}
		
		if (pac && produ == 1 && marca == 96 && modelo == 0)
		{
			alert('Para un Positivo debe al menos seleccionar un producto y/o aparato');
			document.DatClie.lst_Producto.focus();
			return (false);
		}
		
		if (pendi && produ == 1 && marca == 96 && modelo == 0)
		{
			alert('Para un Pendiente debe al menos seleccionar un producto y/o aparato');
			document.DatClie.lst_Producto.focus();
			return (false);
		}
		
		if (!ValidateEmail(document.DatClie.txt_email) && document.DatClie.txt_email.value!="")
		{
			alert('El campo correo electronico ' + sMsg );
			document.DatClie.txt_email.focus();
			return (false);
		}
		
		if (!ValidateLen(document.DatClie.txt_Apellidos,0,40))
		{
			alert('El campo Apellidos del Cliente ' + sMsg );
			document.DatClie.txt_Apellidos.focus();
			return (false);
		}
		
		if (!ValidateLen(document.DatClie.txt_Nombres,0,40))
		{
			alert('El campo Nombres del Cliente ' + sMsg );
			document.DatClie.txt_Nombres.focus();
			return (false);
		}
		
		if (!ValidateLen(document.DatClie.txt_calle,0,50))
		{
			alert('El campo calle del Cliente ' + sMsg);
			document.DatClie.txt_calle.focus();
			return (false);
		}
		
		if (!ValidateLen(document.DatClie.txt_email,0,60))
		{
			alert('El campo email del Cliente ' + sMsg);
			document.DatClie.txt_email.focus();
			return (false);
		} 
		
		if (!ValidateLen(document.DatClie.txt_Observa,0,200))
		{
			alert('El campo Observaciones del Cliente ' + sMsg);
			document.DatClie.txt_Observa.focus();
			return (false);
		}
		
		if (!ValidateLen(document.DatClie.txt_nota,0,200))
		{
			alert('La Nota del evento ' + sMsg);
			document.DatClie.txt_nota.focus();
			return (false);
		}
		alert('Positivo: '+ pac +' Nega: '+ nega+' Pendi: '+ pendi +' Produ: '+ produ +' Marca: '+ marca +' Modelo: '+ modelo);
  	return (false);
	}
</script>
<!-- main calendar program -->
  <script type="text/javascript" src="jscalendar/calendar.js"></script>

  <!-- language for the calendar -->
  <script type="text/javascript" src="jscalendar/lang/calendar-es2.js"></script>
  <script type="text/javascript" src="jscalendar/calendar-setup.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="wgs.css" rel="stylesheet" type="text/css">
<!-- calendar stylesheet -->
<link rel="stylesheet" type="text/css" media="all" href="jscalendar/calendar-blue2.css"/>
</head>

<body bgcolor="#999999" text="#000000" leftmargin="0" topmargin="20" marginwidth="0" marginheight="0" onLoad="setDynaList(arrDL1)">
<%
clienteid=request("clienid")
    Set conn = Server.CreateObject("ADODB.Connection")
    Conn.Open(MM_WGS_STRING)
sql = "SELECT Clientes.ID_Clientes, Clientes.Cli_Numero, Clientes.Cli_Apellido, Clientes.Cli_Nombre, Clientes.Cli_TipoDNI, Clientes.Cli_DNInro, Clientes.Cli_Nacimiento, Clientes.Cli_Genero, Clientes.Cli_Email, Status.Sta_Nombre, Empresas.Emp_Corto, Clientes.Cli_Contacto, Clientes.Cli_Fechaultimo, Clientes.Cli_Fechaprox, Clientes.Cli_Comentario, Clientes.Cli_Telefono2, Clientes.Cli_CPNuevo, Clientes.Cli_Eventos, Clientes.Cli_Valor, Usuarios.Usr_Apellidos, Usuarios.Usr_Nombres, Localidades.Loc_Localidad, Localidades.Loc_Provincia, Localidades.Loc_CP, Localidades.Loc_CPAprov, Clientes.Cli_Inicio, Clientes.Cli_Calle, Clientes.Cli_CalleNro, Clientes.Cli_Piso, Clientes.Cli_Depto, Clientes.Cli_Ampli FROM Localidades INNER JOIN ((Empresas INNER JOIN (Status INNER JOIN Clientes ON Status.ID_Status = Clientes.Cli_Status) ON Empresas.ID_Empresas = Clientes.Cli_Empresa) INNER JOIN Usuarios ON Clientes.Cli_Usuario = Usuarios.ID_Usuarios) ON Localidades.Id_Localidad = Clientes.Cli_Localidad WHERE (((Clientes.Cli_Numero)="& clienteid & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3
%>
<table width='700' border='0' cellspacing='0' cellpadding='5' height='30' align='center'>
		  	<tr bgcolor='#FFFFFF' class="Titulo">
				<td width="50%">Linea : <b><%=HTMLEncode(rs.Fields("Cli_Numero").Value)%></b></td>
				<td width="50%" align='left' class='Titulo'><span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Apellido").Value)%>, </b></span><span class="texto"><b><%=HTMLEncode(rs.Fields("Cli_Nombre").Value)%></b></span></td>
			</tr>
			<tr bgcolor='#FFFFFF' class="Texbuche">
				<td class='texto' align='left'><b><span class="Texbuche">Estado : <b><%=HTMLEncode(rs.Fields("Sta_Nombre").Value)%></b></span></b></td>
				<td align='left' class='Texbuche'><strong>Empresa: <%=HTMLEncode(rs.Fields("Emp_Corto").Value)%></strong></td>
		  </tr>
			<tr bgcolor='#FFFFFF' class="SloganTrebu">
			  <td class='texto' align='left'>Inicializado: <%=FormatMediumDate(rs.Fields("Cli_Inicio").Value)%></td>
			  <td class='texto' align='left'><span class="SloganTrebu">Contactos: <%=HTMLEncode(rs.Fields("Cli_Eventos").Value)%></span></td>
  </tr>
			<tr bgcolor='#FFFFFF'>
			  <td class='Texbuche' align='left'>Último Evento: <%=HTMLEncode(rs.Fields("Usr_Nombres").Value)%>&nbsp;<%=HTMLEncode(rs.Fields("Usr_Apellidos").Value)%> el <%=FormatMediumDate(rs.Fields("Cli_Fechaultimo").Value)%>. </td>
			  <td class='Texbuche' align='left'>Proximo Evento: <%=HTMLEncode(rs.Fields("Cli_Fechaprox").Value)%></td>
		  </tr>
</table>
<form action="cargaevento.asp?clienid=<%=clienteid%>" method="post" name="DatClie" id="DatClie" onSubmit="return validar_form();">
	<table width="700" border="0" align="center" cellpadding="5" cellspacing="0">
      <tr bgcolor="#336699" class="trebuche">
        <td colspan="2" align="left" class="Texbuche"><span class="Titulo">Datos Personales:</span></td>
      </tr>
      <tr bgcolor="#CCCCCC" class="trebuche">
        <td width="50%" align="left" class="Texbuche"><div align="right">Apellidos/Razón Social:</div></td>
        <td width="50%" align="left" bgcolor="#CCCCCC" class="Texbuche"><div align="left"><span class="texto">
          <input name="txt_Apellidos" type="text" class="textbox" id="txt_Apellidos" value="<%=HTMLEncode(rs.Fields("Cli_Apellido").Value)%>" size="40">
        </span>[x]</div></td>
      </tr>
      <tr bgcolor="#F0F0F0" class="trebuche">
        <td width="50%" align="left" class="Texbuche"><div align="right">Nombres:</div></td>
        <td width="50%" align="left" bgcolor="#F0F0F0" class="Texbuche"><div align="left"><span class="texto">
          <input name="txt_Nombres" type="text" class="textbox" id="txt_Nombres" value="<%=HTMLEncode(rs.Fields("Cli_Nombre").Value)%>" size="40">
        </span>[x]</div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Tipo Documento:</div></td>
        <td bgcolor="#CCCCCC" class="texto"><div align="left">
          <%i = rs.Fields("Cli_TipoDNI").Value%>
		  <select name="lst_TipoDNI" size="1" class="textbox" id="lst_TipoDNI">
            <option<% If i="DNI" Then Response.Write " SELECTED"%>>DNI</option>
            <option<% If i="LE" Then Response.Write " SELECTED"%>>LE</option>
            <option<% If i="LC" Then Response.Write " SELECTED"%>>LC</option>
            <option<% If i="CI" Then Response.Write " SELECTED"%>>CI</option>
            <option<% If i="CUIT" Then Response.Write " SELECTED"%>>CUIT</option>
            <option<% If i="PAS" Then Response.Write " SELECTED"%>>PAS</option>
		  </select>
          <span class="Texbuche">[x]</span></div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right"><span class="texto">N&uacute;mero: </span></div></td>
        <td bgcolor="#F0F0F0" class="texto"><div align="left">
          <input name="txt_NroDNI" type="text" class="textbox" id="txt_NroDNI" value="<%=HTMLEncode(rs.Fields("Cli_DniNro").Value)%>" size="10">
          <span class="Texbuche">[x]</span></div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right"><span class="texto">G&eacute;nero:</span></div></td>
        <td bgcolor="#CCCCCC" class="texto"><div align="left" class="Texbuche">
          <%i = rs.Fields("Cli_Genero").Value%>
		  <select name="lst_Genero" size="1" class="textbox" id="lst_Genero">
            <option<% If i="Masculino" Then Response.Write " SELECTED"%>>Masculino</option>
            <option<% If i="Femenino" Then Response.Write " SELECTED"%>>Femenino</option>
            <option<% If i="Juridico" Then Response.Write " SELECTED"%>>Juridico</option>
          </select>
        [x]</div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Nacimiento:</div>        </td>
        <td bgcolor="#F0F0F0" class="texto"><div align="left">
          <input name="f_nacim" type="text" class="textbox" id="f_nacim" value="<%=HTMLEncode(rs.Fields("Cli_Nacimiento").Value)%>" size="10" readonly="true">
          <script type="text/javascript">
    Calendar.setup({
        inputField     :    "f_nacim",     // id of the input field
        ifFormat       :    "%d/%m/%Y", // formatos "%d/%m/%Y" y/o "%d/%m/%Y %k:%M:%S"
     	button		   :    "f_nacim",  // trigger for the calendar (button ID)
        align          :    "CR",           // alignment (defaults to "Bl")
        weekNumbers    :    false,
		firstDay       :    0,
		showsTime      :    false,
		timeFormat     :    "24",
		cache          :    false,
		singleClick    :    false
    });
            </script>
        </div></td>
      </tr>
      <tr align="left" bgcolor="#336699" class="trebuche">
        <td colspan="2" class="Texbuche"><span class="Titulo">Domicilio:</span></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Calle:
          
        </div></td>
        <td bgcolor="#CCCCCC" class="texto"><div align="left" class="Texbuche">
          <input name="txt_calle" type="text" class="textbox" id="txt_calle" value="<%=HTMLEncode(rs.Fields("Cli_Calle").Value)%>" size="40">
        [x]</div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">N&uacute;mero:</div></td>
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="left">
          <input name="txt_calleNro" type="text" class="textbox" id="txt_calleNro" Value="<%=HTMLEncode(rs.Fields("Cli_CalleNro").Value)%>" size="5">
          Piso:
          <input name="txt_piso" type="text" class="textbox" id="txt_piso" Value="<%=HTMLEncode(rs.Fields("Cli_Piso").Value)%>" size="3">
          Depto.:
  <input name="txt_dpto" type="text" class="textbox" id="txt_dpto" Value="<%=HTMLEncode(rs.Fields("Cli_Depto").Value)%>" size="3">
        [x]</div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Ampliaci&oacute;n: </div></td>
        <td bgcolor="#CCCCCC" class="texto"><div align="left">
          <input name="txt_ampli" type="text" class="textbox" id="txt_ampli" value="<%=HTMLEncode(rs.Fields("Cli_Ampli").Value)%>" size="40">
        </div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Localidad:</div></td>
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="left"><b><%=HTMLEncode(rs.Fields("Loc_Localidad").Value)%></b></div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#CCCCCC" class="Texbuche"><div align="right">Provincia:</div></td>
        <td bgcolor="#CCCCCC" class="Texbuche"><div align="left"><b><%=HTMLEncode(rs.Fields("Loc_Provincia").Value)%></b></div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Código Postal Viejo:</div>
          <div align="right"></div>
          <div align="right"></div></td>
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="left"><b><%=HTMLEncode(rs.Fields("Loc_CP").Value)%></b></div>
          <div align="left"></div>
          <div align="left"></div>        </td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#CCCCCC" class="Texbuche"><div align="right"><span class="Texbuche">C&oacute;digo Postal Nuevo: </span></div>
          <div align="center"></div>          <div align="center"></div></td>
        <td bgcolor="#CCCCCC" class="Texbuche"><div align="left">
          <input name="txt_cpa" type="text" class="textbox" id="txt_cpa" value= "<%=HTMLEncode(rs.Fields("Cli_CPNuevo").Value)%>" size="10">
          <input name="Submit2" type="button" class="btn" onClick="javascript:window.open('http://www.correoargentino.com.ar/consulta_cpa/cons2.php?codprov=<%=HTMLEncode(rs.Fields("Loc_CPAprov").Value)%>')" value="CPA">
        </div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td colspan="2" bgcolor="#6699CC" class="Texbuche"><div align="center">Valide la calle y el número  en la p&aacute;gina de c&oacute;digos postales de correo argentino y luego ingrese los mismos datos en este formulario.- </div></td>
      </tr>
      <tr align="left" bgcolor="#336699" class="trebuche">
        <td colspan="2" class="Texbuche"><span class="Titulo">Datos de Contacto:</span></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">
 Persona de Contacto: </div></td>
        <td bgcolor="#CCCCCC" class="texto"><input name="txt_contacto" type="text" class="textbox" id="txt_contacto" value="<%=HTMLEncode(rs.Fields("Cli_Contacto").Value)%>" size="50" >
          <input name="checkbox" type="checkbox" onClick="copiarNombre()" value="copia">
          <span class="Texbuche">[x]</span></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Correo electronico:</div></td>
        <td bgcolor="#F0F0F0" class="texto"><input name="txt_email" type="text" class="textbox" id="txt_email" value= "<%=HTMLEncode(rs.Fields("Cli_Email").Value)%>" size="50"></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Telefono Fijo: </div></td>
        <td bgcolor="#CCCCCC" class="texto"><input name="txt_fijo" type="text" class="textbox" id="txt_fijo" value= "<%=HTMLEncode(rs.Fields("Cli_Telefono2").Value)%>" size="13"></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td colspan="2" bgcolor="#F0F0F0" class="Texbuche"><div align="center">Observaciones</div>
          <div align="center">
            <textarea name="txt_Observa" cols="80" rows="4" class="textbox" id="txt_Observa"><%=HTMLEncode(rs.Fields("Cli_Comentario").Value)%></textarea>
        </div></td>
      </tr>
      <tr align="left" bgcolor="#336699" class="trebuche">
        <td colspan="2" class="Texbuche"><span class="Titulo">Datos del Evento:</span></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right"><span class="texto">Resultado:</span></div></td>
        <td bgcolor="#CCCCCC" class="texto"><div align="left"><span class="Texbuche">
          <%
Set rsX = Server.CreateObject("ADODB.Recordset")
sQuery = "SELECT * FROM Estados WHERE Est_Perfil = 'Operador' AND Est_Camino=True"
rsX.Open sQuery, conn, 3 ,3
If rsX.EOF Then
    Response.Write "&nbsp;No hay estados.<BR>"
Else
    Response.Write "<SELECT NAME=""lst_Estados"" class=""textbox"">"
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
        </span></div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Producto: </div></td>
        <td bgcolor="#F0F0F0" class="Texbuche">
<%
Set rsX = Server.CreateObject("ADODB.Recordset")
sQuery = "SELECT * FROM Productos WHERE Prod_Vigencia = True"
rsX.Open sQuery, conn, 3 ,3
If rsX.EOF Then
    Response.Write "&nbsp;No hay productos.<BR>"
Else
    Response.Write "<SELECT NAME=""lst_Producto"" class=""textbox"">"
	Response.Write "<option value=""" & 1 & """selected>Ninguno</option>"
    Do Until rsX.EOF
        Response.Write "<OPTION VALUE=""" & rsX("ID_Producto") & _
            """>" & rsX("Prod_Codigo") &", "& rsX("Prod_Nombre") & "</OPTION>"
        rsX.MoveNext
    Loop
    Response.Write "</SELECT>"
End If
rsX.Close
Set rsX = Nothing
%>
         [x]</td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Aparato: </div></td>
        <td class="Texbuche">Marca
          <select name="lst_Marca" class="select-type1" id="lst_Marca" onChange="setDynaList(arrDL1)">
		<option Value="96" selected>Ninguna</option>
			    <%
While (NOT rsMain.EOF)
%>
			    <option value="<%=(rsMain.Fields.Item("Id_Marcas").Value)%>"><%=(rsMain.Fields.Item("Mar_Nombre").Value)%></option>
			    <%
  rsMain.MoveNext()
Wend
If (rsMain.CursorType > 0) Then
  rsMain.MoveFirst
Else
  rsMain.Requery
End If
%>
		      </select>
          Modelo
          <select name="lst_Modelo" class="select-type1" id="lst_Modelo">
		      </select>[x]</td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td bgcolor="#F0F0F0" class="Texbuche"><div align="right">Nota:</div></td>
        <td bgcolor="#F0F0F0" class="Texbuche"><span class="texto">
          <textarea name="txt_nota" cols="40" rows="3" class="textbox" id="txt_nota"></textarea>
        </span></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td class="Texbuche"><div align="right">Pr&oacute;ximo llamado o contacto:</div></td>
        <td class="Texbuche"><input name="f_Proximo" type="text" class="textbox" size="20" readonly="true">
        <script type="text/javascript">
    Calendar.setup({
        inputField     :    "f_Proximo",     // id of the input field
        ifFormat       :    "%d/%m/%Y %k:%M", // formatos "%d/%m/%Y" y/o "%d/%m/%Y %k:%M:%S"
     	button		   :    "f_Proximo",  // trigger for the calendar (button ID)
        align          :    "CR",           // alignment (defaults to "Bl")
        weekNumbers    :    false,
		firstDay       :    0,
		showsTime      :    true,
		timeFormat     :    "24",
		cache          :    false,
		singleClick    :    false
    });
            </script>
        [x]</td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td colspan="2" bgcolor="#6699CC" class="Texbuche"><div align="center" class="Texbuche">
          <p>Atención: Para positivos los campos marcados con [x] son obligatorios. </p>
          </div></td>
      </tr>
      <tr align="left" bgcolor="#CCCCCC" class="trebuche">
        <td width="50%" bgcolor="#336699" class="Texbuche"><div align="right">
          <input name="button222" type="reset" class="btn" value="Limpiar">
        </div></td>
        <td width="50%" bgcolor="#336699" class="texto"><div align="left">
          <input name="button22" type="submit" class="btn" value="Continuar">
        </div></td>
      </tr>
</table>
</form>
	<!--*********************************************** -->

	<!-- Se traen el historial de la solicitud y se imprime -->
	<div align="center">
	  <!-- Se traen los posibles caminos a seguir --><br>
</div>
	<center>
<%
rsMain.Close
Set rsMain = Nothing
rsSub.Close
Set rsSub = Nothing
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write(piepagina)
%>
	</center>

</body>
</html>
