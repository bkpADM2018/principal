<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Dim usr, dPass, idPedido, ret, errCode, gIdUsuario, gCdUsuario
Dim secuencia

Function checkOperacion(pIdCotizacion, pSecuencia, pCdUsuario)			
	dim valor, rs, strSQL, ctrl
		
    checkOperacion=ERROR_AUTENTICACION
	if (pCdUsuario = session("Usuario")) then		
		strSQL = "Select * from TBLCTZFIRMAS where IDCOTIZACION=" & pIdCotizacion & " and SECUENCIA=" & pSecuencia
	    'Response.Write strSQL
	    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	    if (not rs.eof) then
	        ctrl = false
	        rolUsuario = CInt(getRolFirma(pCdUsuario, SEC_SYS_COMPRAS))
		    if ((pSecuencia <= PIC_FIRMA_GTE_SECTOR)    and (UCase(rs("CDUSUARIO")) = gCdUsuario))        then ctrl = true		    
		    if ((pSecuencia = PIC_FIRMA_GTE_COMPRAS)    and (rolUsuario= FIRMA_ROL_GTE_COMPRAS))   then ctrl = true
		    if ((pSecuencia = PIC_FIRMA_SUP_PUERTOS)    and (rolUsuario= FIRMA_ROL_SUP_PUERTO))    then ctrl = true
		    if ((pSecuencia = PIC_FIRMA_CONTROLLER)     and (rolUsuario= FIRMA_ROL_CONTROLLER))    then ctrl = true		    
		    if ((pSecuencia = PIC_FIRMA_DIRECCION)      and (rolUsuario= FIRMA_ROL_DIRECTOR))      then ctrl = true		
		    if (ctrl) then checkOperacion = RESPUESTA_OK
	    end if
    end if	    
    
End Function
'------------------------------------------------------------------
Function registrarFirma(pIdCotizacion, pSecuencia, llave)
	Dim strSQL, conn, rs
	strSQL="Update TBLCTZFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", CDUSUARIO='" & session("Usuario") & "', HKEY='" & llave & "' where IDCOTIZACION=" & pIdCotizacion & " and SECUENCIA=" & pSecuencia	
	'Response.Write strSQL
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'------------------------------------------------------------------
Function NotificarPICCompleto()
	Dim strMsg, idUsuario, ds, emailToepfer, idPedido, idUser
	strSQL = "Select * from TBLCTZCABECERA where IDCOTIZACION=" & idCotizacion
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then 
		idPedido = rs("IDPEDIDO")
		cdUser = rs("CDUSUARIO")
	end if	
	if idPedido > 0 then
		Call initHeader(idPedido)
		emailToepfer = obtenerMail(CD_TOEPFER)
	else
		emailToepfer = getUserMail(cdUser)
	end if
	if (emailToepfer <> "") then		
		strMsg = "El Pedido Interno de Compra Nro:" & idCotizacion & " ha sido aprobado." & vbCrLf
		Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web - PIC:" & idCotizacion & " Aprobado"), strMsg, emailToepfer, emailToepfer)
	end if
End Function
'------------------------------------------------------------------
Function revisarEstadoCotizacion(pIdCotizacion)

	Dim rsApertura, strSQL, conn
	
	if (checkPICFinalizado(pIdCotizacion)) then	
		'Se cambia el estado de la cotizacion
		strSQL="Update TBLCTZCABECERA set ESTADO=" & CTZ_FIRMADA & " where IDCOTIZACION=" & pIdCotizacion
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		
		strSQL="Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_APROBADO & " where IDPEDIDO= (SELECT IDPEDIDO FROM TBLCTZCABECERA WHERE IDCOTIZACION=" & pIdCotizacion & ")"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		Call NotificarPICCompleto()
	else
		strSQL="Update TBLCTZCABECERA set ESTADO=" & CTZ_EN_FIRMA & " where IDCOTIZACION=" & pIdCotizacion
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		
	end if

End Function
'------------------------------------------------------------------
Function checkPICFinalizado(pIdCotizacion)
	Dim ret, conn, strSQL, rs

	ret = false
	strSQL = "Select * from TBLCTZFIRMAS where IDCOTIZACION=" & pIdCotizacion & " and HKEY is NULL"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then ret = true
	checkPICFinalizado = ret 
	
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"				
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = UCase(rs("CDUSUARIO"))
			if (session("Usuario") = gCdUsuario) then ret = true			
		else
			gCdUsuario = ""			
		end if
	end if	
	leerRegistroFirmas = ret
End Function
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
'Se llama por AJAX para verificar una credencial
idCotizacion = GF_PARAMETROS7("idCotizacion",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)

Call GP_CONFIGURARMOMENTOS()

ret = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then		
	'1º - Se controla que el usuario tenga permiso para la operación	
	ret = checkOperacion(idCotizacion, secuencia, gCdUsuario)		
	if (ret = RESPUESTA_OK) then				
		Call registrarFirma(idCotizacion, secuencia, HK_readKey())
		Call revisarEstadoCotizacion(idCotizacion)		
	end if
end if
Call HK_sendResponse(ret) 

%>