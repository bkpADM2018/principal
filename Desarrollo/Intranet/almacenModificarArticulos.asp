<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<%
dim myParam
dim mySql, myRs, myCn
myParam = GF_Parametros7("param","",6)
Call initAccessInfo(RES_ADM_AL)

'Tiene acceso a administracion, se verifica que tenga acceso a modificar articulos.
mySql = "SELECT MODIFICAARTICULOS AS VALOR FROM TBLREGISTROFIRMAS WHERE CDUSUARIO='" & session("Usuario") & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, myRs, oConn, "OPEN", mySql)
if not myRs.eof then
	if (myRs("VALOR") <> 1) then response.redirect "comprasAccesoDenegado.asp"
else
	response.redirect "comprasAccesoDenegado.asp"
end if
Call executeQueryDB(DBSITE_SQL_INTRA, myRs, oConn, "CLOSE", mySql)
if myParam = "CHG" then 'Cambio de unidad
	Response.Redirect "almacenChangeUnits.asp"
end if	
%>