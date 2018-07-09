<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosunificador.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<% 
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pNroLote, pFechaLote)
	Dim rs,rsApl,rsEst, flagFirmar, auxNroLote, flagFinalizado
    
    flagFirmar = false
    'Verifico si el estado actual del lote es el autorizado por el jefe de sector, si es asi debo llamar al programa de AS400 para que se ejecute 
    Call executeSP(rsApl, "EJIFL.TBLPROVISIONESCANE_GET_BY_PARAMETERS", pNroLote &"||"& pFechaLote &"||"& PROVISCIONES_ESTADO_AUTORIZADO &"||1||0")
    if (not rsApl.Eof ) then 
        'Ejecuta el programa AS400 volcando datos en la EJIFL.TPE100F1, este proceso ademas actualiza el estado del lote a APLICADO o ERROR        
        auxNroLote = GF_nDigits(pNroLote,9)
        Call executeSP(rs, "STORED.PGM_TPE631", CStr(auxNroLote))
        'Una vez que se ejecuta el programa de AS400 verifico que no haya dado errores para poder continuar con las firmas( debe estar en estado aplicado)
        Call executeSP(rsFir, "EJIFL.TBLPROVISIONESCANE_GET_BY_PARAMETERS", pNroLote &"||"& pFechaLote &"||"& PROVISCIONES_ESTADO_APLICADO &"||1||0")
        if (not rsFir.Eof) then flagFirmar = true
        flagFinalizado = true
    else
        'Al no ser el estado esperado para llamar al procesos de AS400 , actualizo el estado del lote normalmente
        set rs_Ret = executeSP(rs, "EJIFL.TBLPROVISIONESCANE_UPD_FIRMAR_CBTE", pNroLote &"||"& pFechaLote &"||"& Session("Usuario"))
        if (rs_Ret(SP_IDERROR) = ESTADO_ACTIVO) then flagFirmar = true
        flagFinalizado = false
    end if

    'Si todo el proceso anterior esta correcto, agrego la firma 
    if (flagFirmar) then
        Call executeSP(rs, "EJIFL.TBLPROVISIONESFIRMAS_INS", pNroLote &"||"& pFechaLote &"||"& Session("Usuario") &"||"& Session("MmtoDato"))
        Call sendMailNextSignatory(pNroLote ,pFechaLote, flagFinalizado)
        registrarFirma = RESPUESTA_OK
    else
        registrarFirma = ERROR_AUTORIZACION_PROVISIONES
    end if

End Function
'---------------------------------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds	
	ret = false	
	if (HK_isKeyReady()) then
		strSQL = "Select * from TOEPFERDB.TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
        Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true
		else
			gCdUsuario = ""
		end if
	end if		
	leerRegistroFirmas = ret
End Function
'---------------------------------------------------------------------------------------------------------------------
'Se encarga de enviar mail al proximo firmante que tiene el lote, en caso de ser el ultimo envia a los primero autorizantes informando que se aplico
Function sendMailNextSignatory(pNroLote ,pFechaLote, pFlagFinalizado)
    Dim rs, mailMsg, mailOrigen, mailDestino, mailAsunto
    'El sotre procedure devuelve el/los usuarios que deberan ser notificados por la alerta de mail de provisiones
    Call executeSP(rs, "EJIFL.TBLPROVISIONESFIRMAS_GET_NEXT_SIGNATORY_BY_PARAMETERS", pNroLote &"||"& pFechaLote &"||"& Session("Usuario"))
    if (not rs.Eof) then
        'Obtenemos la casilla de mail que enviara el mail de alerta para la tarea provisiones
        mailOrigen = getTaskMailList(TASK_EJE_PROVISIONS, MAIL_TASK_SENDER)
        mailAsunto = "Sistema Provisiones - Alerta de firma"
        if (pFlagFinalizado) then 
            'Provision aplicada, envia mail informandolo a los firmantes que participaron en el proceso 
            mailMsg = "La siguiente provisión se aplicó correctamente: "& vbcrlf
            mailMsg = mailMsg & "Nro.Lote: "& pNroLote & vbcrlf
            mailMsg = mailMsg & "Fecha Lote: "& GF_FN2DTE(pFechaLote) & vbcrlf
        else
            'Provision en firma, envia mail al siguiente firmante indicando que tiene para firmar la provision
            mailMsg = "Tiene pendiente para autorizar la siguiente provisión: "& vbcrlf
            mailMsg = mailMsg & "Nro.Lote: "& pNroLote & vbcrlf
            mailMsg = mailMsg & "Fecha Lote: "& GF_FN2DTE(pFechaLote) & vbcrlf
        end if
        while(not rs.Eof)
            mailDestino = getUserMail(Trim(rs("CDUSUARIO")))
            Call GP_ENVIAR_MAIL(mailAsunto, mailMsg, mailOrigen, mailDestino)
            rs.MoveNext()
        wend
    end if
End Function
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim nroLote,fechaLote,gCdUsuario,gsecuencia

nroLote = GF_PARAMETROS7("nroLote", 0, 6)
fechaLote = GF_PARAMETROS7("fechaLote", 0, 6)
    
Call GP_CONFIGURARMOMENTOS()

respuesta = LLAVE_NO_CORRESPONDE
if (CDbl(nroLote) <> 0)and(CDbl(fechaLote) <> 0) then
	if (leerRegistroFirmas()) then respuesta = registrarFirma(nroLote,fechaLote)
else
	respuesta = CODIGO_VACIO
end if	
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
