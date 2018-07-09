<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<%
'--------------------------------------------------------------------------------------------------------------------
Dim idContrato,idObra,idArea,idDetalle,rs,accion

idContrato = GF_PARAMETROS7("idContrato",0,6)
idObra     = GF_PARAMETROS7("idObra",0,6)
idArea	   = GF_PARAMETROS7("idArea",0,6)
idDetalle  = GF_PARAMETROS7("idDetalle",0,6)
accion     = GF_PARAMETROS7("accion","",6)

if (accion = ACCION_BORRAR) then 
	strSQL = "DELETE FROM TBLCTCPARTIDAS WHERE IDCONTRATO = "&idContrato&" AND IDOBRA = "&idObra&_
			 " AND IDAREA = "&idArea& " AND IDDETALLE= " &idDetalle
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end if
%>