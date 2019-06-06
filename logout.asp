<!--#include file="Connections/database.asp" -->
<%If session("wgsANM_usr")<>"" then
set conn=server.createobject("ADODB.Connection")
	conn.open MM_WGS_STRING
    psql="select * from Usuarios where ID_Usuarios= " & session("wgsANM_usr")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open psql,conn,1,2
    if not(rs.eof) then
		rs("Usr_Estado")="Off Line"
	    rs("Usr_Passtime")=Now()
		rs.update
	else
		message="Usuario y/o Contraseña Invalidos"
	end if
rs.close
set rs=nothing
conn.close
set conn=nothing	
session("wgsANM_usr")=""
session("wgsANM_lvl")=""
session.abandon
end if
%>
<script language="JavaScript">
top.location='login.asp'
</script>