<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%
dim CdUsuario,IdUsuario,respuesta,idAfe,gSecuencia(), gCdUsuario

'---------------------------------------------------------------------------------------------
Function registarFirma(pAfeId,cdUsuario)
	Dim strSQL, conn, rs, k, sp_ret
	    
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLDATOSAFE_UPD_FIRMAR_CBTE", pAfeId & "||" & cdUsuario & "$$respuesta")
    if (CInt(sp_ret("respuesta")) = 0) then 
        Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLAFEFIRMAS_INS", pAfeId & "||" & cdUsuario & "||" & session("MmtoDato"))	
        registarFirma = RESPUESTA_OK
    else
        'Se indica el error por no esncontrar el nuevo estado del AFE.
        registarFirma = ESTADO_AFE_NO_DETERMINADO
    end if        
End Function
'---------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = UCase(rs("CDUSUARIO"))			
			if (gCdUsuario = session("Usuario")) then ret = true			
		else
			gCdUsuario = ""
		end if
	end if
	leerRegistroFirmas = ret
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   sendMailNextSignerAFE
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   27/09/2013
' Objetivo:
'			   Busca el siguiente firmante del Afe y le notifica por Mail que tiene una firma pendiente de AFE 
'			   En caso de que sea el director, no se le notifica(como lo hace comprasAutorizaciones)	
' Parametros:
'			   pIdAfe			[int]	ID Afe
' Devuelve:	   
'                -
' Modificacion :	  
'               CNA - Ajaya Nahuel
' Fecha Modificacion :	  
'               27-10-2014
' Objetivo modificacion :	  
'               Se adaptó a la nueva forma de firmar (por estados de transicion), de esta manera cada rol 
'               puede tener mas de un usuario asociado a el.
'-------------------------------------------------------------------------------------------- 
Function sendMailNextSignerAFE(pIdAfe)
	Dim emailTo, emailToepfer, asunto, msg	
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLAFEFIRMAS_GET_NEXT_USER", pIdAfe)
	if (not rs.Eof) then
        emailToepfer = obtenerMail(CD_TOEPFER)
        asunto = "AFE - ALERTA FIRMA PENDIENTE "
        msg = "Usted tiene disponible para firmar el AFE : "& getCdAFE(pIdAfe)
        while not rs.Eof
            if (rs("IDROL") <> FIRMA_ROL_DIRECTOR) then
                emailTo = getUserMail(rs("CDUSUARIO"))
		        Call GP_ENVIAR_MAIL(asunto, msg, emailToepfer, emailTo )
		    end if
	        rs.MoveNext()
        wend
    end if
End Function
'**********************************************************	
'*****************  INICIO DE PAGINA  *********************
'**********************************************************
Dim aux,auxKey, ret 


idAfe = GF_Parametros7("IDAFE",0,6)

Call GP_CONFIGURARMOMENTOS()

ret = LLAVE_NO_CORRESPONDE
if (idAfe <> "") then
    if (leerRegistroFirmas()) then		
    	'1º - Se controla que el usuario tenga permiso para la operación	    	
        ret = registarFirma(idAfe,gCdUsuario)
    	Call sendMailNextSignerAFE(idAfe)
    end if
else
    respuesta = CODIGO_VACIO
end if
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(ret)


%>
