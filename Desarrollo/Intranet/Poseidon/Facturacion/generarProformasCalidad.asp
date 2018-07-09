<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosFechas.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosLog.asp"-->
<!--#include file="../../Includes/procedimientosFacturacionCalidad.asp"-->
<%

'--------------------------------------------------------------------------------------------------------------
'Modifica la fecha de Facturacion dependiendo si es para Tercero o propios de la empresa.
Function modificarFechaFacturacion(pCliente,pFecha,pPto)
	if (pCliente = FACT_ACOND_DESCARGA_3ROS) then
		Call updateValueParametro(PARAM_FACT_FECHA_3ROS, pFecha, pPto)
		logMig.info("Se actualizo el parametro DTULTFACTACONDE con la fecha " & pFecha)
	else
		Call updateValueParametro(PARAM_FACT_FECHA_PROPIAS, pFecha, pPto)
		logMig.info("Se actualizo el parametro DTULTFACTACONDP con la fecha " & pFecha)
	end if	
End Function
'-----------------------------------------------------------------------------------------------------------------
' Función:	
'			obtenerFechaFacturacion
' Autor: 	
'			CNA - Ajaya Nahuel
' Fecha: 	
'			04/07/2014
' Objetivo:	
'			Obtener la ultima fecha de migracion que se prosesó, dependiendo si es tercero o propio.
' Parametros:
'			pTipoDescarga   [string]
'			pPto	  		[string] 
' Devuelve:
'			Ultima fecha migrada
'--------------------------------------------------------------------------------------------
Function obtenerFechaFacturacion(pTipoDescarga,pPto)	
	Dim auxFecha	
	if (pTipoDescarga = FACT_ACOND_DESCARGA_3ROS) then		
		auxFecha = getValueParametro(PARAM_FACT_FECHA_3ROS,pPto)		
	else
		auxFecha = getValueParametro(PARAM_FACT_FECHA_PROPIAS,pPto)		
	end if
	obtenerFechaFacturacion = GF_DTEADD(auxFecha, 1, "D")
End Function
'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'****************************************************

Dim pto, myHoy, logMig, tipoDescarga, flagBach, myHasta, myDesde, myUltimo

'On Error Resume Next

Call GP_CONFIGURARMOMENTOS

flagBach = false

if (session("usuario") = "") then session("usuario") = "SYNC"
myHoy = Left(session("MmtoDato"), 8)
	
pto = GF_PARAMETROS7("pto", "", 6)
tipoDescarga = GF_PARAMETROS7("td", "", 6)
if (tipoDescarga = "") then tipoDescarga = FACT_ACOND_DESCARGA_PROPIAS
myDesde = GF_PARAMETROS7("fd", "", 6)
myHasta = GF_PARAMETROS7("fh", "", 6)
if (myDesde = "") then
	myDesde = obtenerFechaFacturacion(tipoDescarga, pto)
	if (CLng(myDesde) > CLng(myHoy)) then myDesde = myHoy
	myHasta = myDesde
	flagBach = true
end if

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "FACT-ACOND-SYNC-" & myHoy

logMig.info("####################################################")
logMig.info("		-PUERTO       :  "& pto)
logMig.info("		-MOMENTO      :  "& GF_FN2DTE(Left(session("MmtoSistema"),8)))
logMig.info("		-TIPO DESCARGA:  "& tipoDescarga )
logMig.info("		-USUARIO      :  "& session("Usuario"))	
logMig.info("		-FECHA MIGRADA:  "& GF_FN2DTCONTABLE(myDesde) & " a " & GF_FN2DTCONTABLE(myHasta))
logMig.info("####################################################")	

while (myDesde <= myHasta)
	logMig.info("PROCESANDO FECHA:  "& GF_FN2DTCONTABLE(myDesde))	
	Call migrarMermasAFacturar(pto, myDesde, "", "", TIPO_TRANSPORTE_CAMVAG, tipoDescarga, logMig)
	myUltimo = myDesde
	myDesde = GF_DTEADD(myDesde, 1, "D")	
wend	
'Solo se actualiza la fecha de migracion si se esta llamando al proceso de manera bachera.
if (flagBach) then Call modificarFechaFacturacion(tipoDescarga,myUltimo,pto)
logMig.info("####################################################")
logMig.info("---               FIN DEL PROCESO                ---")
logMig.info("####################################################")	

%>