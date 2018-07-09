<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<%
dim myIdCotizacion, myIdPedido
dim strSQL, rs, conn, rtrn, newIdVale
myIdCotizacion = GF_Parametros7("idCotizacion",0,6)
myIdPedido = GF_Parametros7("idPedido",0,6)

strSQL = "UPDATE TBLCTZCABECERA SET ESTADO = '" & CTZ_ANULADA & "' where IDCOTIZACION=" & myIdCotizacion	
call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)

if myIdPedido <> 0 then
	'Si luego de anular el PIC no quedan ni contratos ni PICs activos se retrocede el estado del pedido.
	strSQL="Select * from TBLCTZCABECERA where IDPEDIDO=" & myIdPedido & " and ESTADO <> '" & CTZ_ANULADA & "'"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then
		strSQL="Select * from TBLOBRACONTRATOS where IDPEDIDO=" & myIdPedido & " and ESTADO <> " & ESTADO_CTC_CANCELADO
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (rs.eof) then
			strSQL = "UPDATE TBLPCTCABECERA SET ESTADO = '" & ESTADO_PCT_ADJUDICADO & "' where IDPEDIDO=" & myIdPedido
			call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	end if
end if
'-----------------------------------------------------------------------------------------
%>