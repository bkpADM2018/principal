<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<%
Dim idContrato, strSQL, rs, ret, myIdPedido

idContrato = GF_PARAMETROS7("idContrato", 0, 6)

ret = CTC_NO_EXISTE

'Se verifica que el contrato pueda anularse
strSQL="Select * from TBLOBRACONTRATOS where IDCONTRATO=" & idContrato
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.eof) then
	myIdPedido = rs("IDPEDIDO")
	strSQL="Update TBLOBRACONTRATOS set CDUSERCONF='" & session("Usuario") & "', MMTOCONF=" & session("MmtoDato") & ", ESTADO=" & ESTADO_CTC_CANCELADO & " where IDCONTRATO=" & idContrato
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	ret = RESPUESTA_OK
	
	'Si luego de anular el contrato no quedan ni contratos ni PICs activos se retrocede el estado del pedido.
	strSQL="Select * from TBLCTZCABECERA where IDPEDIDO=" & myIdPedido & " and ESTADO <> '" & CTZ_ANULADA & "'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then
		strSQL="Select * from TBLOBRACONTRATOS where IDPEDIDO=" & myIdPedido & " and ESTADO <> " & ESTADO_CTC_CANCELADO
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (rs.eof) then
			strSQL = "UPDATE TBLPCTCABECERA SET ESTADO = '" & ESTADO_PCT_ADJUDICADO & "' where IDPEDIDO=" & myIdPedido
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		end if
	end if
	
end if
Response.write ret
%>
