<%
sindatos="No Hay Datos Disponibles"
piepagina="WG Sistemas &copy; 2006"
database="db/BaseWG.mdb"
Remitente="floresa@arnet.com.ar"
Receptor="floresanibal@yahoo.com.ar"

//MM_WGS_STRING="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& server.mappath(database)
MM_WGS_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\anibal\desktop\wgs\db\BaseWG.mdb"
//MM_WGS_STRING = "DSN=wgsistemas"

// Funciones del sistema

function FormatMediumDate(DateValue)
    Dim strYYYY
    Dim strMM
    Dim strDD
    Dim strHMS
if Not IsNull(DateValue) then
        strYYYY = CStr(DatePart("yyyy", DateValue))

        strDD = CStr(DatePart("d", DateValue))
        If Len(strDD) = 1 Then strDD = "0" & strDD

	    strMM = CStr(DatePart("m", DateValue))
        If Len(strMM) = 1 Then strMM = "0" & strMM

		strHMS = FormatDateTime(DateValue,3) & " "

        FormatMediumDate = strDD & "/" & strMM & "/" & strYYYY & " " & strHMS
    else
	    FormatMediumDate = ""
	end if	
End Function

function HTMLEncode(instring)
  If Not IsNull(instring) then
    HTMLEncode=Server.HTMLEncode(instring)
  else
    HTMLEncode=""
  end if
End function

function NroDiscar(Nro,Prefix)
	NroDiscar = Prefix & "-15-" & Right(Nro,Len(Nro)-Len(Prefix))
End function
%>
