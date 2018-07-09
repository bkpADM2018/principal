<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->

<%
'------------------------------------------------------------------
Function registrarFirma(pIdProveedor, pSecuencia, llave)
	Dim strSQL, conn, rs
	strSQL="Update TOEPFERDB.TBLEMPRESASFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", CDUSUARIO='" & session("Usuario") & "', HKEY='" & llave & "' where IDEMPRESA=" & pIdProveedor & " and SECUENCIA=" & pSecuencia	
	Call executeQuery(rs, "EXEC", strSQL)	
End Function
'------------------------------------------------------------------
Function checkOperacion(pIdProveedor, pSecuencia)			
		rol = getRolFirma(gCdUsuario, SEC_SYS_PROVEEDORES)
		if (rol = FIRMA_ROL_LEGALES) then usuarioEspecial = LEGALES_USER
		strSQL = "Select * from TOEPFERDB.TBLEMPRESASFIRMAS where IDEMPRESA=" & pIdProveedor & " and SECUENCIA=" & pSecuencia
		Call executeQuery(rs, "OPEN", strSQL)
		if (not rs.eof) then		
			if ((rs("CDUSUARIO") = gCdUsuario) or (rs("CDUSUARIO")=usuarioEspecial)) then
				checkOperacion = RESPUESTA_OK
			end if
		end if
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TOEPFERDB.TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"		
		
		Call executeQuery(rs, "OPEN", strSQL)
		
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true			
		else
			gCdUsuario = ""			
		end if
	end if
	leerRegistroFirmas = ret
End Function
'-----------------------------------------------------------------------------------------
Function enviarMail(pIdProveedor)
    Dim listaMails
    
    Call loadDataDB(pIdProveedor)
    
	mensaje = "Se ha autorizado al siguiente proveedor:" &vbcrlf&vbcrlf
	mensaje = mensaje & " - Nro. Proveedor: " 	& idProveedor			&vbcrlf
	mensaje = mensaje & " - Razon Social: " 	& razsoc 				&vbcrlf
	mensaje = mensaje & " - Nom Ampliado: " 	& nomamp 				&vbcrlf
	mensaje = mensaje & " - Tipo Documento: " 	& getDsTipoDoc(tipdoc) 	&vbcrlf	
	mensaje = mensaje & " - Nro. Documento: " 	& nrodoc 				&vbcrlf	
	mensaje = mensaje & " - Tipo Proveedor: " 	& getDsTipoProv(tiprov) &vbcrlf	
	Call GP_ENVIAR_MAIL("Proveedores - Legales Autorizó un nuevo proveedor.", mensaje,  SENDER_LEGALES, SENDER_SUPPLIERS)    
    enviarMail = RESPUESTA_OK
End function
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
Dim idProveedor, rtrnd, errCode, gCdUsuario
Dim secuencia
idProveedor = GF_PARAMETROS7("idProveedor",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
Call GP_CONFIGURARMOMENTOS()
rtrnd = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then		    
	'1º - Se controla que el usuario tenga permiso para la operacion	
	rtrnd = checkOperacion(idProveedor, secuencia)
	if (rtrnd = RESPUESTA_OK) then				
		Call registrarFirma(idProveedor, secuencia, HK_readKey())
		call enviarMail(idProveedor)
	end if
end if
Response.Write HK_sendResponse(rtrnd)

%>