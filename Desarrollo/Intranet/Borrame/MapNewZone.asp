<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosMap.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 09/01/2009                         -->
'<!--               Pagina     : MapNewZone.ASP                     -->
'<!--               Descripcion: Alta de Zonas para ToepferMaps     -->
'<!------------------------------------------------------------------->
dim cdZone, cn, rs, sql, msgE, msgS, cnMax, rsMax, MaxId
dim myOnLoad
cdZone = GF_Parametros7("cdZone", "", 6)
MaxId = 1
if len(cdZone) > 0 then
	sql = "Select * from Groups where dsGroup='" & cdZone & "'"
	call GF_BD_Control_Map (rs, cn, "OPEN", sql)
	if not rs.eof then 
		msgE = "Ya existe un grupo con ese nombre!"
	else
		call GF_BD_Control_Map (rsMax, cnMax, "OPEN", "select max(idGroup) + 1 as MaxId from Groups")
		if not isNull(rsMax("MaxId")) then MaxId = rsMax("MaxId")
		sql = "Insert into Groups values(" & MaxId & ", '" & cdZone & "')"
		call GF_BD_Control_Map(rs, conn, "EXEC", sql)
		msgS = "La zona '" & cdZone & "' se dio de alta satisfactoriamente!"
	end if
	call GF_BD_Control_Map(rs, conn, "CLOSE", sql)
else
	msgE = "Debe ingresar un nombre para la zona!"
end if	
if len(msgS) > 0 then myOnLoad = "sendCombo()"
%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<title>Intranet ActiSA - Nueva Zona</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="Scripts/iwin.js"></script>  
</head>
<script language="JavaScript">
	function sendCombo(){
		window.parent.addToCombo(document.getElementById("idZone").value, document.getElementById("cdZone").value);
	}
</script>
<body onload="<%=myOnLoad%>">
<form name="frmMain" method="post">
<table width="100%" border=0 valign="center" align="center" class="reg_header45">
	<tr>
		<td><b><%=GF_Traducir("Nombre")%></b></td>
		<td>
			<input type="text" id="cdZone" name="cdZone" value="<%=cdZone%>">
			<input type="hidden" id="idZone" name="idZone" value="<%=MaxId%>">
		</td>
		<td align="right">
			<input type="submit" name="Aceptar" value="Aceptar">
		</td>
	</tr>
	<% if len(msgE)>0 then %>
	<tr class="TDERROR">	
		<td colspan="3">
			<% Response.Write msgE %>		
		</td>
	</tr>
	<% end if %>
	<% if len(msgS)>0 then %>
	<tr class="TDSUCCESS">	
		<td colspan="3">
			<% Response.Write msgS %>		
		</td>
	</tr>
	<% end if %>
	
</table>	
</form>
</body>
</html>