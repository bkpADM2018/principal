<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<%
dim myDescomposicion, puerto
cdProducto = GF_Parametros7("cdProducto", "", 6)
puerto = Request("Pto")
if cdProducto <> "" then
	strSQL="Select top 5 CDCOSECHA from COSECHAS where CDCOSECHA <> 0 and CDPRODUCTO = " & cdProducto & " order by CDCOSECHA desc "
	Call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
	if (not rs.eof) then 
		listOfHarvest = rs.GetString(2,,, ";")
		listOfHarvest = left(listOfHarvest, Len(listOfHarvest)-1) 'Saco el último ';'
	else 
		listOfHarvest = "SIN COSECHAS"
	end if
end if
Response.Write listOfHarvest
%>