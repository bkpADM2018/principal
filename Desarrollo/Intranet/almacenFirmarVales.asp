<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
'------------------------------------------------------------------------------------------------
' Función:	controlarFirmaAJS
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	12/11/2013
' Objetivo:	
'			Controla las firmas del vale de Ajuste de Sotck, si el vale tiene que firmarlo el Director y esta listo para que lo haga, envia un mail a auditoria
' Parametros:
'			pIdVale	[int] 
' Devuelve:
'			-
'--------------------------------------------------------------------------------------------
Function controlarFirmaAJS(pIdVale)
	Dim strSQL, msg, totalRegistros
	'Obtengo todas las firmas que faltan firmar		
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLVALESFIRMAS_GET_BY_FILTERS", pIdVale & "||||||IS NULL||SECUENCIA DESC")	
	If not rs.Eof then
		totalRegistros = CDbl(rs.RecordCount)		
	end if	
End Function
'------------------------------------------------------------------
Function registrarFirma(pIdVale, pSecuencia, llave)
	Dim strSQL, conn, rs
	strSQL="Update TBLVALESFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", CDUSUARIO='" & session("Usuario") & "', HKEY='" & llave & "' where IDVALE=" & pIdVale & " and SECUENCIA=" & pSecuencia	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
End Function
'------------------------------------------------------------------
Function checkOperacion(pIdVale, pSecuencia)			
	dim valor, rs, strSQL, conn, controlSecuenciaOk, rs2, usuarioEspecial, ROL

	checkOperacion=ERROR_AUTENTICACION
	controlSecuenciaOk = false
	if (pSecuencia = VS_FIRMA_RESPONSABLE) then
		'Si intenta firmar el responsable, es el primero que debe hacerlo, esta todo OK.		
		controlSecuenciaOk = true
	else		
		'Intenta firmar otro usuario que no es el responsable, se busca la secuencia correspondiente al responsable del vale para asegurarse que este ya firmo antes que intenten firmar el resto.
		strSQL = "Select 1 from TBLVALESFIRMAS where IDVALE=" & pIdVale & " and SECUENCIA=" & VS_FIRMA_RESPONSABLE & " and HKEY is not null"		
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
		if (not rs2.eof) then	controlSecuenciaOk = true		
	end if	
	if (controlSecuenciaOk) then		
		usuarioEspecial = ""
		rol = getRolFirma(gCdUsuario, SEC_SYS_ALMACENES)
		if (rol = FIRMA_ROL_RESP_PUERTO) then usuarioEspecial = VS_NO_USER 
		if (rol = FIRMA_ROL_AUDITOR) then usuarioEspecial = VS_AUDIT_USER
		if (rol = FIRMA_ROL_SUP_PUERTO) then usuarioEspecial = VS_PORT_SUPERVISOR_USER	
		if (rol = FIRMA_ROL_DIRECTOR) then usuarioEspecial = DIRECTOR_USER
		strSQL = "Select * from TBLVALESFIRMAS where IDVALE=" & pIdVale & " and SECUENCIA=" & pSecuencia
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
		if (not rs2.eof) then		
			if ((rs2("CDUSUARIO") = gCdUsuario) or (rs2("CDUSUARIO")=usuarioEspecial)) then
				checkOperacion = RESPUESTA_OK
			end if
		end if
		Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "CLOSE", strSQL)
	end if	
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then		
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
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
Dim idVale, rtrn, errCode, gCdUsuario
Dim secuencia
idVale = GF_PARAMETROS7("idVale",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
Call GP_CONFIGURARMOMENTOS()
rtrn = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then		    
	call initHeaderVale(idVale)
	'1º - Se controla que el usuario tenga permiso para la operacion	
	rtrn = checkOperacion(idVale, secuencia)
	if (rtrn = RESPUESTA_OK) then				
		Call registrarFirma(idVale, secuencia, HK_readKey())
		if vs_cdVale = CODIGO_VS_AJUSTE_STOCK then Call controlarFirmaAJS(idVale)
	end if
end if
Call HK_sendResponse(rtrn)
%>