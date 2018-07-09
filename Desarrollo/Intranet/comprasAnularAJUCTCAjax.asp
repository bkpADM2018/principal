<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
dim myIdAjuste, idContrato
dim strSQL, rs, rtrn

myIdAjuste = GF_Parametros7("idAjuste",0,6)
idContrato = GF_Parametros7("idContrato",0,6)

'Se elimina el ajuste
strSQL = "DELETE FROM TBLOBRACTCAJUSTES where IDAJUSTE=" & myIdAjuste
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
'Se eliminan las firmas del ajuste
strSQL = "DELETE FROM TBLOBRACTCAJUSTESFIRMAS where IDAJUSTE=" & myIdAjuste
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	

'Analizar si hay que cambiar el estado del contrato
strSQL = "SELECT COUNT(*) AS CANTIDAD FROM TBLOBRACTCAJUSTES WHERE IDCONTRATO = " & idContrato & " AND APLICADO='" & TIPO_NEGACION & "'"
'Response.Write strSQL
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if cint(rs("CANTIDAD")) = 0 then
	'ACTUALIZAR CABECERA PIC ESTADO
	strSQL = "UPDATE TBLOBRACONTRATOS SET ESTADO='" & ESTADO_CTC_AUTORIZADO & "' WHERE IDCONTRATO=" & idContrato
	'Response.Write strSQLAux
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
end if

%>