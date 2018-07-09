<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<%
dim myIdAjuste, myIdCotizacion
dim strSQL, rs, conn, rtrn
myIdAjuste = GF_Parametros7("idAjuste",0,6)
myIdCotizacion = GF_Parametros7("idCotizacion",0,6)
'Se elimina el ajuste
strSQL = "DELETE FROM TBLCTZAJUSTES where IDAJUSTE=" & myIdAjuste
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)
'Se eliminan las firmas del ajuste
strSQL = "DELETE FROM TBLCTZAJUSTESFIRMAS where IDAJUSTE=" & myIdAjuste
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)

'Analizar si hay que cambiar el estado de la ctoizacion
strSQL = "SELECT COUNT(*) AS CANTIDAD FROM TBLCTZAJUSTES WHERE IDCOTIZACION = " & myIdCotizacion & " AND APLICADO='" & TIPO_NEGACION & "'"
'Response.Write strSQL
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if cint(rs("CANTIDAD")) = 0 then
	'ACTUALIZAR CABECERA PIC ESTADO
	strSQLAux = "UPDATE TBLCTZCABECERA SET ESTADO='" & CTZ_FIRMADA & "' WHERE IDCOTIZACION=" & myIdCotizacion
	'Response.Write strSQLAux
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQLAux)
end if

%>