<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->

<%
'------------------------------------------------------------------
Function registrarFirma(pIdAjuste, pSecuencia, llave)
	Dim strSQL, conn, rs
	strSQL="Update TBLOBRACTCAJUSTESFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", CDUSUARIO='" & session("Usuario") & "', HKEY='" & llave & "' where IDAJUSTE =" & pIdAjuste & " and SECUENCIA=" & pSecuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'------------------------------------------------------------------
Function checkOperacion(pIdAjuste, pSecuencia)			
	dim valor, rs, strSQL, controlSecuenciaOk, rs2, rolUsuario
	checkOperacion=ERROR_AUTENTICACION
	strSQL = "Select CF.* from TBLOBRACTCAJUSTESFIRMAS CF where IDAJUSTE=" & pIdAjuste & " and SECUENCIA=" & pSecuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (not rs.eof) then
        ctrl = false
        rolUsuario = CInt(getRolFirma(gCdUsuario, SEC_SYS_COMPRAS))
	    if ((pSecuencia = CTC_FIRMA_RESPONSABLE)    and (rs("CDUSUARIO") = gCdUsuario))         then ctrl = true
	    if ((pSecuencia = CTC_FIRMA_GTE_SECTOR)     and (rs("CDUSUARIO") = gCdUsuario))         then ctrl = true        
	    if ((pSecuencia = CTC_FIRMA_GTE_COMPRAS)    and (rolUsuario= FIRMA_ROL_GTE_COMPRAS))    then ctrl = true		    
	    if (ctrl) then checkOperacion = RESPUESTA_OK
    end if
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then

		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"		
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
		
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true			
		else
			gCdUsuario = ""			
		end if
	end if
	leerRegistroFirmas = ret
End Function
'---------------------------------------------------------------------------------------------------------------------------------
function esUltimaFirma(pIdAjuste)
dim strSQL, rs, con
	esUltimaFirma = false
	strSQL = "SELECT idAjuste FROM TBLOBRACTCAJUSTESFIRMAS WHERE IDAJUSTE = " & pIdAjuste & " AND (HKEY='' or HKEY is Null)"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then esUltimaFirma = true
end function
'---------------------------------------------------------------------------------------------------------------------------------
Function actualizarContrato(pIdAjuste)
    'Actualizamos el importe del contrato.
    strSQL = "Select IDCONTRATO, TIPOAJUSTE, IMPORTEPESOS, IMPORTEDOLARES from TBLOBRACTCAJUSTES where IDAJUSTE = " & pIdAjuste
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			
    while (not rs.eof)
        'Se actualiza el estado del contrato y los importes.
        strSQL="Update TBLOBRACONTRATOS set"
        if (rs("TIPOAJUSTE") = CTC_AJUSTE_GENERAL) then
            'Es un ajuste presupuestario
            strSQL= strSQL & " IMPORTEPESOS=IMPORTEPESOS + " & rs("IMPORTEPESOS") & ", IMPORTEDOLARES=IMPORTEDOLARES + " & rs("IMPORTEDOLARES") 
        else
            'Es un ajuste de valor unitario
            strSQL= strSQL & " IMPORTEUNITARIOPESOS=IMPORTEUNITARIOPESOS + " & rs("IMPORTEPESOS") & ", IMPORTEUNITARIODOLARES=IMPORTEUNITARIODOLARES + " & rs("IMPORTEDOLARES") 
        end if
        strSQL = strSQL & ", ESTADO=" & ESTADO_CTC_AUTORIZADO & " where IDCONTRATO=" & rs("IDCONTRATO")	    
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
			
        'Aplicamos todos los ajustes.
        strSQL = "Update TBLOBRACTCAJUSTES set APLICADO='" & TIPO_AFIRMACION & "' WHERE IDAJUSTE = " & pIdAjuste
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	   
        rs.MoveNext()
    wend
End Function    
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
Dim rtrn, errCode, gCdUsuario, idAjuste
Dim secuencia
idAjuste = GF_PARAMETROS7("idAjuste",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
Call GP_CONFIGURARMOMENTOS()
rtrn = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then
	'1º - Se controla que el usuario tenga permiso para la operacion	
	rtrn = checkOperacion(idAjuste, secuencia)
	if (rtrn = RESPUESTA_OK) then				
		Call registrarFirma(idAjuste, secuencia, HK_readKey())
		if (esUltimaFirma(idAjuste)) then actualizarContrato(idAjuste)		
	end if
end if
Call HK_sendResponse(rtrn)
%>