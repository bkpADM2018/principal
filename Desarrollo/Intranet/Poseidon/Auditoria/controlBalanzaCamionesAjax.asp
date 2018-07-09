<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<%

Function changeEstatusControlBZA(pPto, pIdControl, estado)
	dim strSQL, rs
	strSQL = "UPDATE CTRLBZACAMIONES SET ESTADO = " & estado & " WHERE IDCONTROL = " & pIdControl
	GF_BD_Puertos pPto, rs, "EXEC", strSQL
End function
'--------------------------------------------------------------------------------------------------

Dim accion, pto, idControl, estado

idControl = GF_PARAMETROS7("idControl", 0, 6)
pto		  = GF_PARAMETROS7("pto", "", 6)
accion	  = GF_PARAMETROS7("accion", "", 6)
estado	  =	GF_PARAMETROS7("estado", 0, 6)

if(accion = ACCION_CANCELAR)then
	call changeEstatusControlBZA(pto, idControl, estado)
	response.end
end if


'**************************************************************************

%>