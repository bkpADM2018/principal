<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="poseidonAjusteAutorizacionPrint.asp"-->
<% 
'----------------------------------------------------------------------------------------------
' Función:	  sendEmailDraftAuthorized
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/11/2013
' Objetivo:	  Enviar mail notificando que el Director ha firmado el Draft Survey, en esta instancia se cierra el proceso 
'			  de autorización. La lista que recibirán el mail se encuentra en una Lista de Correo, la constante
'			  para identificar a la lista es:
'			  		LISTA_DRAFT_AUTORIZADO
'			  Además en el mail va a ir ajunto el archivo PDF autorizado
' Parametros:
'			  [int]		pIdAjuste - Id Ajuste	
'			  [string]	pCdAjuste - Codigo Ajuste
'			  [int]		pPto	  - Puerto
' Devuelve:
'			  -
'----------------------------------------------------------------------------------------------
Function sendEmailDraftAuthorized(pIdAjuste, pCdAjuste, pPto)
	Dim pathPDF, asunto, mensaje, listaMails
	
	pathPDF = armadoPDF(pPto,pIdAjuste, "", "", pCdAjuste, PDF_FILE_MODE,0)	
	asunto  = GF_TRADUCIR(getDsCodigoAjustePuerto(pCdAjuste) & " Autorizado.")
	mensaje = GF_TRADUCIR("El ajuste por "& getDsCodigoAjustePuerto(pCdAjuste) &" - N°: "& pIdAjuste &" fué autorizado por todos los responsables, se adjunta el informe.")
	listaMails = pCdAjuste & "-" & getLetraPuerto(pPto)
	Call SendMail(TASK_POS_ADM_AJUSTES, listaMails, asunto, mensaje, pathPDF)    
	
End Function
'---------------------------------------------------------------------------------------------------------------------
Function sendNotifyEmail(pIdAjuste, pCdAjuste, pCdUsuario, pPto)
    Dim emailToepfer,rs,strTo
    'Esl siguiente store procedure se encarga de obtener el proximo usuario que debera firmar el ajuste o lote, este store se da cuenta por si solo el rol que debera firmar
    Call executeProcedureDb(pPto, rs, "TBLAJUSTES_GET_NEXT_SIGNATORY_BY_PARAMETERS", pCdUsuario &"||"& pCdAjuste &"||"& pIdAjuste )
    if not rs.Eof then
        emailToepfer = getTaskMailList(TASK_POS_ADM_AJUSTES, MAIL_TASK_SENDER)
        strTo = ""
        while (not rs.eof)			
            strTo = strTo & getUserMail(rs("CDUSUARIO")) & ";"
            rs.MoveNext()
        wend
        msg = "Está disponible para su firma un ajuste de stock realizado en el puerto de " & pPto & vbCrLf
        msg = msg & "Motivo del Ajuste: " & getDsCodigoAjustePuerto(pCdAjuste) & vbCrLf
        Call GP_ENVIAR_MAIL("POSEIDON - Sistema de Puertos - Ajuste de Stock", msg, emailToepfer, strTo)
    end if
End function
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pCdUsuario,pIdAjuste,pCdAjuste,pPto)
	Dim rs
    Call executeProcedureDb(pPto, rs, "TBLAJUSTESFIRMAS_INS", pCdUsuario &"||"& pIdAjuste &"||"& HK_readKey() &"||"& Session("MmtoDato"))
    Call executeProcedureDb(pPto, rs, "TBLAJUSTES_UPD_FIRMAR_CBTE", pCdUsuario &"||"& pCdAjuste &"||"& pIdAjuste)

    'if ((pCdAjuste = AJUSTE_DRAFT_SURVEY)OR(pCdAjuste = AJUSTE_CALIDAD)OR(pCdAjuste = AJUSTE_MANIPULEO)) then
        'Si es un ajuste de Draft o Calidad o Manipuleo controlo si fue la ultima firma
        strSQL = "SELECT * FROM TBLAJUSTES WHERE IDAJUSTE = "& pIdAjuste &" AND ESTADO = " & AJUSTE_ESTADO_AUTORIZADO
        Call executeQueryDb(pPto, rs, "OPEN",strSQL)
        if (not rs.Eof) then
            if (pCdAjuste = AJUSTE_DRAFT_SURVEY) then
                'Actualizo el draft
                strSQL = "UPDATE TBLEMBARQUESDRAFTSURVEY SET CDESTADO = " & ESTADO_AUTORIZADO & " WHERE IDDRAFT = " & rs("IDORIGEN")
		        Call executeQueryDb(pPto, rs, "EXEC",strSQL)
                Call sendEmailDraftAuthorized(pIdAjuste, pCdAjuste, pPto)
            end if
        else
            'No es ultima firma, informo por mail el proximo firmante
            Call sendNotifyEmail(pIdAjuste, pCdAjuste, pCdUsuario, pPto)
        end if
	'end if
End Function
'---------------------------------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim strSQL, rs, ret, km, ds	
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
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim gCdUsuario,gsecuencia,respuesta

idAjuste = GF_PARAMETROS7("idAjuste",0,6)
cdAjuste = GF_PARAMETROS7("cdAjuste","",6)
pto = GF_PARAMETROS7("pto","",6)


Call GP_CONFIGURARMOMENTOS()
respuesta = LLAVE_NO_CORRESPONDE

if (leerRegistroFirmas()) then			
    respuesta = RESPUESTA_OK
    Call registrarFirma(gCdUsuario,idAjuste,cdAjuste,pto)
end if

if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
