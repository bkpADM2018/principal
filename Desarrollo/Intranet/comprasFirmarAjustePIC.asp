<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
'------------------------------------------------------------------
Function checkOperacion(pIdAjuste, pSecuencia)			

	dim valor, rs, strSQL, ctrl
		
    checkOperacion=ERROR_AUTENTICACION
	strSQL = "Select * from TBLCTZAJUSTESFIRMAS where IDAJUSTE=" & pIdAjuste & " and SECUENCIA=" & pSecuencia
    'Response.Write strSQL
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    if (not rs.eof) then
        ctrl = false
        rolUsuario = CInt(getRolFirma(gCdUsuario, SEC_SYS_COMPRAS))        
	    if ((pSecuencia <= PIC_FIRMA_GTE_SECTOR)    and (rs("CDUSUARIO") = gCdUsuario))        then ctrl = true
	    if ((pSecuencia = PIC_FIRMA_GTE_COMPRAS)    and (rolUsuario= FIRMA_ROL_GTE_COMPRAS))   then ctrl = true
	    if ((pSecuencia = PIC_FIRMA_SUP_PUERTOS)    and (rolUsuario= FIRMA_ROL_SUP_PUERTO))    then ctrl = true
	    if ((pSecuencia = PIC_FIRMA_CONTROLLER)     and (rolUsuario= FIRMA_ROL_CONTROLLER))    then ctrl = true
	    if ((pSecuencia = PIC_FIRMA_DIRECCION)      and (rolUsuario= FIRMA_ROL_DIRECTOR))      then ctrl = true		
	    if (ctrl) then checkOperacion = RESPUESTA_OK
    end if    
End Function
'------------------------------------------------------------------
Function registrarFirma(pIdAjuste, pSecuencia, llave)
	Dim strSQL, conn, rs	
	strSQL="Update TBLCTZAJUSTESFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", CDUSUARIO='" & session("Usuario") & "', HKEY='" & llave & "' where IDAJUSTE = " & pIdAjuste & " and SECUENCIA=" & pSecuencia	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then

		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
		if (not rs.eof) then
			gCdUsuario = UCase(rs("CDUSUARIO"))
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
    esUltimaFirma = False
	strSQL = "SELECT idAjuste FROM TBLCTZAJUSTESFIRMAS WHERE IDAJUSTE = " & pIdAjuste & " AND (HKEY='' or HKEY is NULL)"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then esUltimaFirma = True
end function
'---------------------------------------------------------------------------------------------------------------------------------
Function actualizarPIC(pIdAjuste)

    Dim rs, rsAux, strSQL, myEstado
    
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_AJUSTES_BY_IDAJUSTE", pIdAjuste)
    
    while (not rs.eof)
        				
        'ACTUALIZAR DETALLE PIC, CARGAR NUEVO DETALLE PARA QUE COINCIDAN LOS IMPORTES								
	    strSQL = "UPDATE TBLCTZDETALLE SET IMPORTEPESOS=IMPORTEPESOS + " & rs("IMPORTEPESOS_AJU") & " , IMPORTEDOLARES=IMPORTEDOLARES + " & rs("IMPORTEDOLARES_AJU") & ", CANTIDAD= CANTIDAD + " & rs("CANTIDAD_AJU") & " WHERE IDCOTIZACION=" & rs("IDCOTIZACION") & " AND IDARTICULO=" & rs("IDARTICULO_DET") & " AND IDAREA=" & rs("IDAREA_AJU") & " AND IDDETALLE=" & rs("IDDET_AJU")
	    'Response.Write strSQL
	    call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    	
	    'Se actualiza el monto facturado.
	    if (rs("CDMONEDA") = MONEDA_PESO) then
		    strSQL="Update TBLCTZDETALLE set FACTURADO=(CANTIDAD*IMPORTEPESOSFACTURADO)/IMPORTEPESOS where IDCOTIZACION=" & rs("IDCOTIZACION") & " AND IDARTICULO=" &  rs("IDARTICULO_DET") & " AND IDAREA=" & rs("IDAREA_AJU") & " AND IDDETALLE=" & rs("IDDET_AJU")
	    else
		    strSQL="Update TBLCTZDETALLE set FACTURADO=(CANTIDAD*IMPORTEDOLARESFACTURADO)/IMPORTEDOLARES where IDCOTIZACION=" & rs("IDCOTIZACION") & " AND IDARTICULO=" &  rs("IDARTICULO_DET") & " AND IDAREA=" & rs("IDAREA_AJU") & " AND IDDETALLE=" & rs("IDDET_AJU")
	    end if
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    	
    	'ACTUALIZAR ESTADO DE LOS AJUSTES
        strSQL = "UPDATE TBLCTZAJUSTES SET APLICADO='" & TIPO_AFIRMACION & "' WHERE IDAJUSTE= " & pIdAjuste
        'Response.Write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
        
	    'ACTUALIZAR CABECERA PIC, MODIFICAR TOTALES Y ESTADO
	    strSQL="Select * from TBLCTZAJUSTES where IDCOTIZACION=" & rs("IDCOTIZACION") & " and APLICADO='" & TIPO_NEGACION & "'"
		call executeQueryDb(DBSITE_SQL_INTRA, rsAux, "OPEN", strSQL)
	    myEstado = CTZ_FIRMADA
	    if (not rsAux.eof) then myEstado = CTZ_EN_AJUSTE
	    
	    strSQL = "UPDATE TBLCTZCABECERA SET ESTADO='" &  myEstado & "', IMPORTEPESOS=IMPORTEPESOS + " & rs("IMPORTEPESOS_AJU") & ", IMPORTEDOLARES=IMPORTEDOLARES + " & rs("IMPORTEDOLARES_AJU") & " WHERE IDCOTIZACION=" & rs("IDCOTIZACION")
	    'Response.Write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    
	    rs.MoveNext()
    wend        
    
    Call NotificarAICCompleto(pIdAjuste)
        	    
End Function
'------------------------------------------------------------------
Function NotificarAICCompleto(pIdAjuste)
	Dim strMsg, destinatario
	strSQL = "Select * from TBLCTZAJUSTES where IDAJUSTE=" & pIdAjuste
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then 
	    cdUser = UCase(rs("CDUSUARIO"))
	    destinatario = getUserMail(cdUser)			
	    if (destinatario <> "") then		
		    strMsg = "PIC Nro: " & rs("IDCOTIZACION") & " - El ajuste " & pIdAjuste & " ha sido aprobado."
		    Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web - PIC:" & rs("IDCOTIZACION") & " - AJUSTE Aprobado"), strMsg, MAILTO_COMPRAS, destinatario)
	    end if
    end if	    
End Function
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
Dim myRtrn, gCdUsuario
Dim secuencia, idAjuste

idAjuste = GF_PARAMETROS7("idAjuste",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
Call GP_CONFIGURARMOMENTOS()
myRtrn = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then
	'1º - Se controla que el usuario tenga permiso para la operacion	
	myRtrn = checkOperacion(idAjuste, secuencia)
	if (myRtrn = RESPUESTA_OK) then					
		Call registrarFirma(idAjuste, secuencia, HK_readKey())		
		if (esUltimaFirma(idAjuste)) then actualizarPIC(idAjuste)
	end if
end if
Call HK_sendResponse(myRtrn)
%>