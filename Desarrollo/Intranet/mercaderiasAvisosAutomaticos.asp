<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Poseidon/reporteCuposResolucion25Print.asp"-->
<%

'******************************************************************************************************************
'
'	Pagina que se ejecuta via una tarea programada de manera de actualizar el estado de todos los proyectos del 
'	sector de Mercaderia curso.
'	Esto sirve para mantener el sistema actualizado en caso de que nadie ingrese al sistema y se cumplan plazos 
'	específicos.
'
'	NOTA: se incluye la página que genera el reporte de Cupos Resolución 25-13 (reporteCuposResolucion25Print.asp),  
'		  de esta manera solo se llama a la función que crea el pdf y en esta página se lo adjunta al mail.
'	
'******************************************************************************************************************

Function enviarMail(pTitulo,pDestino,pMensaje,pAdjunto) 
	Dim rtrn
	rtrn=false
	if (pDestino <> "") then		
		Call logInfo("Mail enviado a: " & pDestino)			
		Call GP_ENVIAR_MAIL_ATTACHMENT(pTitulo, pMensaje, SENDER_MERCADERIAS, pDestino, pAdjunto)		
		rtrn = true
	end if	
	enviarMail = rtrn
End Function
'--------------------------------------------------------------------------------------------------------------
Function buscarAlertasCuposResolucion(pPto)	
	Dim cuposProveedor,fecha,flagCupo, mailTo, cupoArr, cupoTra, emailToepfer	
	'Informa los cupos del dia siguiente al actual
	fecha = GF_DTEADD(left(session("MmtoSistema"),8),1,"D")	
	Call logInfo("##########################################")
	Call logInfo("BUSCANDO ALERTAS CUPOS RESOLUCION 25/13 - Fecha : " & GF_FN2DTE(fecha))
	Call logInfo("##########################################")
		
	'Armo el PDF y obtengo la ruta del archivo	
	pathPDF = armarPDF(fecha, pPto, PDF_FILE_MODE)
	if (pathPDF <> "") then
	    Call logInfo("Se encontraron cupos en " & pPto)
	    strMsg = "Se adjuntan los cupos asignados correspondiente a la fecha " & GF_FN2DTE(fecha) & vbCrLf & vbCrLf
	    strMsg = strMsg & "Empresa: " & getDsClienteByCUIT(CUIT_TOEPFER) & vbCrLf
	    Set fs = Server.CreateObject("Scripting.FileSystemObject")
	    mailDestino = SENDER_CUPOS_RESOLUCION_25_13
	    Call enviarMail(GF_TRADUCIR("Planilla de informacón de espacio físico - "& pPto),mailDestino ,strMsg & "Puerto: " & pPto, pathPDF)
	    'Luego de enviar el mail, borro el archivo del directorio
	    Call fs.deleteFile(pathPDF, true)	
    else
    	 Call logInfo("No se generò archivo para enviar.")   
    end if	    
	Call logInfo("------------------FIN DE LA BUSQUEDA----------------")    
End function
'-----------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'********************************************	INICIO DE PAGINA   ************************************************
'******************************************************************************************************************

pto = GF_PARAMETROS7("pto", "", 6)
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
'Seteo de Valores necesarios par ala session
Call GP_ConfigurarMomentos()
session("Usuario") = "JAS"
'Se ejecutan las alertas.
Call buscarAlertasCuposResolucion(pto)

%>