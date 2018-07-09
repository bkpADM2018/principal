<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
dim idDivision, fechaCierre, tipoCierre, cotizacionDolar, idCierre, strSQL, rs, con, rtrn, anio, mes, idCierreEXPO
idDivision = GF_Parametros7("idDivision", 0, 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
cotizacionDolar = GF_Parametros7("cotizacionDolar", "", 6)
tipoCierre = GF_Parametros7("tipoCierre", "", 6) 
anio = left(fechaCierre,4)
mes = mid(fechaCierre,5,2)
idCierreEXPO = 0
idCierre = getIdCierre2(anio, mes, idDivision, tipoCierre)
if idCierre > 0 then
	strSQL ="UPDATE TBLCIERRESASIENTOS2 SET IMPORTEDOLARES=IMPORTEPESOS / " & replace(cotizacionDolar,",",".") & ", TIPOCAMBIO=" & replace(cotizacionDolar,",",".") & " WHERE IDCIERRE=" & idCierre
	Response.Write "Actualizacion tipo cambio " & strSQL		
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)

	strSQL ="update TBLCIERRESASIENTOS2 set ccostos='' where (cdcuenta like '185%' or cdcuenta like '114%' or cdcuenta ='" & CUENTA_PROVISIONES & "') and IDCIERRE=" & idCierre
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	Response.Write "Actualizacion cuenta en blanco " & strSQL

end if
Response.Write "Hecho..."
Response.End 
%>