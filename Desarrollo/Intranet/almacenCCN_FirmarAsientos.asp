<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%
Dim usr, dPass, idPedido, ret, errCode, gCdUsuario
Dim secuencia

Function checkOperacion(pCierre, pSecuencia)			
	dim valor, rs, strSQL, conn
		
	checkOperacion = ERROR_AUTENTICACION
	if pSecuencia = FIRMA_ROL_RESP_CONTADURIA then
		if getRolFirma(session("Usuario"), SEC_SYS_ALMACENES) = FIRMA_ROL_RESP_CONTADURIA then checkOperacion = RESPUESTA_OK
	end if
	if pSecuencia = FIRMA_ROL_RESP_PUERTO then
		strSQL = "Select * from TBLCIERRESCABECERA2 where IDCIERRE=" & pCierre 
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then 
			if puedeFirmarAsientos(session("Usuario"),rs("IDDIVISION")) then checkOperacion = RESPUESTA_OK
		end if	
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	end if
End Function
'------------------------------------------------------------------
Function registrarFirma(pCierre, pSecuencia, llave)
	Dim strSQL, conn, rs
	
	strSQL="Update TBLCIERRESFIRMAS2 SET FECHAFIRMA=" & session("MmtoDato") & ", HKEY='" & llave & "', CDUSUARIO='" & UCASE(session("Usuario")) & "' where IDCIERRE=" & pCierre & " and SECUENCIA=" & pSecuencia	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds

	gCdUsuario = ""
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then			
			gCdUsuario = rs("CDUSUARIO")
			ret = true
		end if
	end if
	leerRegistroFirmas = ret
End Function
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
'Se llama por AJAX para verificar una credencial
idCierre = GF_PARAMETROS7("idCierre",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)

Call GP_CONFIGURARMOMENTOS()

ret = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then	
	'1º - Se controla que el usuario tenga permiso para la operación	
	ret = checkOperacion(idCierre, secuencia)	
	if (ret = RESPUESTA_OK) then Call registrarFirma(idCierre, secuencia, HK_readKey())
end if
Call HK_sendResponse(ret)

%>