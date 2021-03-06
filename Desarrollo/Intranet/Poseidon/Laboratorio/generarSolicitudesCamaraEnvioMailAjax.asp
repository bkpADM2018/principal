<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../../Includes/procedimientosuser.asp"-->
<!--#include file="../../Includes/procedimientosLaboratorio.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosLog.asp"-->
<!--#include file="../../includes/procedimientosMail.asp"-->

<%
'-------------------------------------------------------------------------------------------------------------------
Function enviarMailError(pPto,pMsg,pAsunto)
    Dim strMsg, auxDestino,emailOrigen
    emailOrigen = getTaskMailList(TASK_POS_INFO_ANALISIS, MAIL_TASK_SENDER)
    strMsg = "Error al intentar enviar los archivos de Analisis de Camara." & vbcrlf &_
             pMsg & vbcrlf &_
             "	Puerto:" &pPto & vbcrlf
    auxDestino = getTaskMailList(TASK_POS_INFO_ANALISIS, MAIL_TASK_ERROR_LIST)
    logMig.info("ERROR: " & pMsg)
    logMig.info("Enviando mail de error a : " & auxDestino)
    Call GP_ENVIAR_MAIL(pAsunto, strMsg, emailOrigen, auxDestino)
End Function
'------------------------------------------------------------------------------------------------------------------
Function enviarMailAnalisisCamara(pDestino,pStrAtt,pAsunto, pMsg)
    Dim emailOrigen    
    emailOrigen = getTaskMailList(TASK_POS_INFO_ANALISIS, MAIL_TASK_SENDER)
    if (session("Usuario") = "JAS") then pDestino = "scalisij@toepfer.com"
    logMig.info("Enviando mail a : " & pDestino)
    Call GP_ENVIAR_MAIL_ATTACHMENT(pAsunto,pMsg, emailOrigen, pDestino,pStrAtt)
    'response.Write "Mail Enviado a: " & pDestino & "<br>"
End Function
'******************************************************************************************
'				INICIO DE PAGINA
'******************************************************************************************
Dim pto,valParameterPath,strNamePathCabecera,strNamePathDetalle,strNamePathCuenta,nameFileZip,auxDestino,respuesta
Dim strAtt,procesar, strNamePathErrCam, flagSent,logMig

pto = GF_Parametros7("pto","",6)

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
fileNameLogExp = "EXPORTACION_CAMARA_"& Ucase(pto) &"_" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)
logMig.fileName = fileNameLogExp

logMig.info("Iniciando envio de mail")

Set fso = CreateObject("Scripting.FileSystemObject")
'------------------------------------------------------------------------------------
valParameterPath = Server.MapPath(".") & "\Archivos\Solicitudes"
'------------------------------------------------------------------------------------
respuesta = FILE_MISSING

strNamePathCabecera = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_CABECERA
strNamePathReport  = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_REPORTE    
if (pto <> DBSITE_BAHIA) then
    strNamePathDetalle  = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_ANALISIS
    strNamePathCuenta   = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_CUENTAYORDEN
end if

strFileError = ""

if (fso.FolderExists(valParameterPath)) then
	if (not fso.FileExists(strNamePathCabecera)) then strFileError = "Falta Archivo: " & strNamePathCabecera & VbCrLf
	if (not fso.FileExists(strNamePathReport)) then strFileError = strFileError & "Falta Archivo: " & strNamePathReport & VbCrLf
    if (pto <> DBSITE_BAHIA) then        
        if (not fso.FileExists(strNamePathDetalle)) then strFileError = strFileError & "Falta Archivo: " &  strNamePathDetalle & VbCrLf
	    if (not fso.FileExists(strNamePathCuenta)) then strFileError = strFileError & "Falta Archivo: " & strNamePathCuenta & VbCrLf
    end if
	if (strFileError = "") then			
		auxDestino = getTaskMailList(TASK_POS_INFO_ANALISIS, MAIL_TASK_INFO_LIST)		
		strAtt = strNamePathCabecera 
		strAtt = strAtt &";"& strNamePathReport
        if (pto <> DBSITE_BAHIA) then 
            strAtt = strAtt &";"& strNamePathDetalle &";"& strNamePathCuenta
        end if
		if (auxDestino <> "") then Call enviarMailAnalisisCamara(auxDestino,strAtt,"Poseidon - Exportacion Analisis de C�mara", "Se generaron los archivos de analisis para la C�mara. Se adjuntan los mismos.")				
		'fso.DeleteFile(strNamePathCabecera)
		'fso.DeleteFile(strNamePathDetalle)
		'fso.DeleteFile(strNamePathCuenta)
		respuesta = auxDestino
	else
		Call enviarMailError(pto,strFileError,"Poseidon - Exportacion Analisis de C�mara - ERROR")
	end if
else
	Call enviarMailError(pto, "No existe el directorio: " & valParameterPath,"Poseidon - Exportacion Analisis de C�mara - ERROR")
end if

logMig.info("Finalizando envio de mail")

%>
