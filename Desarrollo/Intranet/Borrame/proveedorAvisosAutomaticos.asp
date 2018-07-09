<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="comprasReporteFacturaDupdoDeProveedorPrintXLS.asp"-->
<!--#include file="comprasReportePagoDupdoAProveedorPrintXLS.asp"-->
<!--#include file="interfacturas/interfacturasReportePrintXLS.asp"-->
<%
'******************************************************************************************************************
'
'	Pagina que se ejecuta via una tarea programada de manera de actualizar el estado de todos los proyectos del 
'	sector de Proveedores.
'	Esto sirve para mantener el sistema actualizado en caso de que nadie ingrese al sistema y se cumplan plazos 
'	específicos.
'
'******************************************************************************************************************

Function enviarMail(pTitulo,pDestino,pMensaje,pAdjunto) 
Dim rtrn, sender
rtrn=false
if (pDestino <> "") then		
	Call logInfo("Mail enviado a: " & pDestino)
	sender = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_SENDER)	
	Call GP_ENVIAR_MAIL_ATTACHMENT(pTitulo, pMensaje, sender, pDestino, pAdjunto)			
	rtrn = true
end if	
enviarMail = rtrn
End Function
'------------------------------------------------------------------------------------------------------------------
Function enviarReportePagoDuplicado()
Dim fechaActual, fechaDesde ,rss, lista
fechaActual = left(session("MmtoSistema"),8) 
fechaDesde = GF_DTEADD(fechaActual, -7, "D")
fechaActual = GF_DTEADD(fechaActual, -1, "D")

Call logInfo("##########################################")
Call logInfo("GENERANDO REPORTE PAGOS DUPLICADOS EN EL PERIODO DE: - Fecha Desde: "&GF_FN2DTE(fechaDesde)&" - Fecha Hasta: "&GF_FN2DTE(fechaActual))
Call logInfo("##########################################")

strMsg = "Se adjunta reporte EXCEL de pagos duplicados a proveedores en el periodo de: "&GF_FN2DTE(fechaDesde)&" - "&GF_FN2DTE(fechaActual)
nameFile = loadReportPagoDuplicado(XLS_FILE_MODE,fechaDesde,fechaActual)
strPathAttachment = Server.mapPath("temp/" & nameFile)
Call logInfo("Hay pagos ducplicados")
        
lista = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_INFO_LIST)  
IF NOT enviarMail("Reporte pagos duplicados a proveedores",lista,strMsg,strPathAttachment) THEN
    Call logInfo("ERROR - No se pudo enviar el REPORTE a la dirección: " & lista)
    lista = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_ERROR_LIST)         
    Call enviarMail("ERROR en Reporte Pago Duplicado A Proveedores", lista, "Error enviando el mail de Reporte Pagos Duplicados, por favor consulte el log para mayores datos.", "")
ELSE
    Call logInfo("E-MAIL ENVIADO. CON EXISTO A :" & lista & "- ADJUNTO :" & strPathAttachment)
END IF
Call logInfo("----------------FIN DE LA BUSQUEDA-------------")	
End Function

'------------------------------------------------------------------------------------------------------------------
Function enviarReporteFacturaDuplicada()
Dim fechaActual, fechaDesde, lista 
fechaActual = left(session("MmtoSistema"),8)
fechaDesde = GF_DTEADD(fechaActual, -7, "D")
fechaActual = GF_DTEADD(fechaActual, -1, "D")

Call logInfo("##########################################")
Call logInfo("GENERANDO REPORTE FACTURAS DUPLICADOS DE PROVEEDOR EN EL PERIODO DE: - Fecha Desde: "&GF_FN2DTE(fechaDesde)&" - Fecha Hasta: "&GF_FN2DTE(fechaActual))
Call logInfo("##########################################")

strMsg = "Se adjunta reporte EXCEL de facturas duplicadas de proveedores en el periodo de: "&GF_FN2DTE(fechaDesde)&" - "&GF_FN2DTE(fechaActual)
    
nameFile = loadReportFacturasDuplicadas(XLS_FILE_MODE,right(fechaDesde, 6), right(fechaActual, 6))
strPathAttachment = Server.mapPath("temp/" & nameFile)
Call logInfo("Hay facturas ducplicadas")

lista = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_INFO_LIST)         
IF NOT enviarMail("Reporte facturas duplicadas de proveedores",lista,strMsg,strPathAttachment) THEN
    Call logInfo("ERROR - No se puedo enviar el REPORTE a la dirección: " & lista)
    lista = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_ERROR_LIST)         
    Call enviarMail("ERROR en Reporte facturas duplicadas de proveedores", lista, "Error enviando el mail de Reporte facturas duplicadas de proveedores, por favor consulte el log para mayores datos.", "")
ELSE
    Call logInfo("E-MAIL ENVIADO. CON EXISTO A :" & lista & "- ADJUNTO :"& strPathAttachment)
END IF

Call logInfo("----------------FIN DE LA BUSQUEDA-------------")	
End Function 
'------------------------------------------------------------------------------------------------------------------
Function enviarReporteFacturasNegativas()
Dim fechaActual, fechaDesde, lista 
fechaActual = left(session("MmtoSistema"),8)
fechaDesde = GF_DTEADD(fechaActual, -7, "D")
fechaActual = GF_DTEADD(fechaActual, -1, "D")

Call logInfo("##########################################")
Call logInfo("GENERANDO REPORTE FACTURAS EMITIDAS CON VALOR NEGATIVO EN EL PERIODO DE: - Fecha Desde: "&GF_FN2DTE(fechaDesde)&" - Fecha Hasta: "&GF_FN2DTE(fechaActual))
Call logInfo("##########################################")

strMsg = "Se adjunta reporte EXCEL de facturas emitidas con valor negativo en el periodo de: "&GF_FN2DTE(fechaDesde)&" - "&GF_FN2DTE(fechaActual)
    
nameFile = loadReportFactura(fechaDesde,fechaActual, 1,XLS_FILE_MODE) 
strPathAttachment = Server.mapPath("temp/" & nameFile)
Call logInfo("Hay facturas con valor negativo.")

lista = getTaskMailList(TASK_FAC_MAIL_ALERT, MAIL_TASK_INFO_LIST)         
IF NOT enviarMail("Reporte facturas emitidas con valor negativo",lista,strMsg,strPathAttachment) THEN
    Call logInfo("ERROR - No se puedo enviar el REPORTE a la dirección: " & lista)
    lista = getTaskMailList(TASK_FAC_MAIL_ALERT, MAIL_TASK_ERROR_LIST)         
    Call enviarMail("ERROR en Reporte facturas emitidas con valor negativo", lista, "Error enviando el mail de Reporte facturas emitidas con valor negativo, por favor consulte el log para mayores datos.", "")
ELSE
    Call logInfo("E-MAIL ENVIADO. CON EXISTO A :" & lista & "- ADJUNTO :"& strPathAttachment)
END IF

Call logInfo("----------------FIN DE LA BUSQUEDA-------------")	
End Function 
'------------------------------------------------------------------------------------------------------------------
Function enviarReporteFacturasCero()
Dim fechaActual, fechaDesde, lista 
fechaActual = left(session("MmtoSistema"),8)
fechaDesde = GF_DTEADD(fechaActual, -7, "D")
fechaActual = GF_DTEADD(fechaActual, -1, "D")

Call logInfo("##########################################")
Call logInfo("GENERANDO REPORTE FACTURAS EMITIDAS CON VALOR CERO EN EL PERIODO DE: - Fecha Desde: "&GF_FN2DTE(fechaDesde)&" - Fecha Hasta: "&GF_FN2DTE(fechaActual))
Call logInfo("##########################################")

strMsg = "Se adjunta reporte EXCEL de facturas emitidas con valor cero en el periodo de: "&GF_FN2DTE(fechaDesde)&" - "&GF_FN2DTE(fechaActual)
    
nameFile = loadReportFactura(fechaDesde,fechaActual, 0,XLS_FILE_MODE) 
strPathAttachment = Server.mapPath("temp/" & nameFile)
Call logInfo("Hay facturas con valor cero.")

lista = getTaskMailList(TASK_FAC_MAIL_ALERT, MAIL_TASK_INFO_LIST)         
IF NOT enviarMail("Reporte facturas emitidas con valor cero",lista,strMsg,strPathAttachment) THEN
    Call logInfo("ERROR - No se puedo enviar el REPORTE a la dirección: " & lista)
    lista = getTaskMailList(TASK_FAC_MAIL_ALERT, MAIL_TASK_ERROR_LIST)         
    Call enviarMail("ERROR en Reporte facturas emitidas con valor cero", lista, "Error enviando el mail de Reporte facturas emitidas con valor cero, por favor consulte el log para mayores datos.", "")
ELSE
    Call logInfo("E-MAIL ENVIADO. CON EXISTO A :" & lista & "- ADJUNTO :"& strPathAttachment)
END IF

Call logInfo("----------------FIN DE LA BUSQUEDA-------------")	
End Function 
'------------------------------------------------------------------------------------------------------------------


'******************************************************************************************************************
'********************************************	INICIO DE PAGINA   ************************************************
'******************************************************************************************************************
Dim logMig

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "PROVEEDORES_AVISOS_AUTOMATICOS_" & left(session("MmtoDato"),8)

'Seteo de Valores necesarios par ala session
Call GP_ConfigurarMomentos()
session("Usuario") = "SYNC"

'Se ejecutan las alertas.
Call enviarReportePagoDuplicado()
Call enviarReporteFacturaDuplicada()
Call enviarReporteFacturasCero()
Call enviarReporteFacturasNegativas()
%>