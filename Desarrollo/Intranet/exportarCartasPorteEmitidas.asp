<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->

<%
'------------------------------------------------------------------------------------------------------------------
Function getSqlCdProductoTercero()
getSqlCdProductoTercero = "(select cdtercero from tblconversiones where nucuitcliente = '" & CONV_KEY_AFIP & "'"&_
			 " and tipodato = '" & CONV_KEY_PRODUCTO & "' and cdpropio = cc.cdproducto)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlCuitEntregador()
getSqlCuitEntregador = "(select nucuit from entregadores e where e.cdentregador = cc.cdentregador)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlPesoNeto()
getSqlPesoNeto = "(select (select vlpesada from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " & PESADA_BRUTO & "  " &_
			 " and sqpesada = (select max(sqpesada) from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_BRUTO & ")) "&_             
             " - (select vlpesada from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_TARA & " "&_                          
             " and sqpesada = (select max(sqpesada) from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_TARA & ")))"
end function
'------------------------------------------------------------------------------------------------------------------
Function obtenerCartasPorteEmitidas(pFecha,pPto)
	Dim strSQL,myWhere
	Dim sqlPesoNeto, sqlCdProductoTercero, sqlCuitEntregador
	myWhere = " where cc.CDESTADO in (5, 6) "
	Call mkWhere(myWhere, "cc.DtContable", pFecha, "=", 3)
	Call mkWhere(myWhere, "cc.cdtipocamion", CIRCUITO_CAMION_CARGA, "=", 1)
	
	sqlCdProductoTercero = getSqlCdProductoTercero()
	sqlCuitEntregador = getSqlCuitEntregador()
	sqlPesoNeto = getSqlPesoNeto()	
				
		
		
	strSQL = "select do.nctipotransporte as TipoTransporte, "&_
			 " do.nctipocartaporte AS TipoCartaPorte, "&_
			 " ltrim(rtrim(do.nccartaporte)) as NroCartaPorte, "&_
			 " ltrim(rtrim(do.ncau)) as NroCEE, "&_
			 " ltrim(rtrim(do.ncctg)) AS NroCTG, "&_
			 " CAST( (right(do.dcarga,2) + substring(cast(do.dcarga as varchar),6,2) + left(do.dcarga,4)) as varchar) as FechaCarga, "&_
			 " case when do.ncuitremitente is null then '00000000000' else do.ncuitremitente end as CuitTitular, "&_
			 " case when do.ncuitc1 is null then '00000000000' else do.ncuitc1 end as CuitIntermediario, "&_
			 " case when do.ncuitc2 is null then '00000000000' else do.ncuitc2 end as CuitRteComercial, "&_
             " case when do.ncuitcorre is null then '00000000000' else do.ncuitcorre end as CuitCorredor, " &_
             " case when " & sqlCuitEntregador & " is null then '00000000000'  else ltrim(rtrim(" & sqlCuitEntregador & ")) end as CuitRepteEntregador, "&_                          
             " case when do.ncuitdestinatario is null then '00000000000' else ltrim(rtrim(do.ncuitdestinatario)) end as CuitDestinatario, "&_             
             " case when do.nccuitdest is null then '00000000000' else do.nccuitdest end as CuitEstablDestino, "&_             
             " case when do.ncuittransportista is null then '00000000000' else do.ncuittransportista end as CuitTransportista, "&_                          
             " case when do.ncuitchofer is null then '00000000000' else do.ncuitchofer end as CuitChofer, "&_
			 " substring(cast(cdcosecha as varchar(8)),3,2) + '-' + right(cast(cdcosecha as varchar(8)),2) as Cosecha, "&_             
             " replicate('0', 3 - len(" & sqlCdProductoTercero & ")) + cast (" & sqlCdProductoTercero & " as varchar) as CodigoEspecie, "&_
			 " '00' as TipoGrano, "&_      
             " '00000000000000000001' as Contrato, "&_
			 " 1 as TipoPesado, " &_
             " replace(replicate('0', 11-len( " & sqlPesoNeto & ")) + cast(" & sqlPesoNeto & " as varchar),'.',',') as PesoNeto, " &_
             " replicate('0', 6-len(do.ncestableproce)) + cast(ltrim(rtrim(do.ncestableproce)) as varchar) as CdEstableProcedencia, " &_
             " replicate('0', 5-len(do.nclocproce)) + cast(ltrim(rtrim(do.nclocproce)) as varchar) as CdLocalidadProcedencia, "&_             
             " replicate('0', 6-len(do.ncestabledest)) + cast(ltrim(rtrim(do.ncestabledest)) as varchar) as CdEstableDestino, "&_             
             " replicate('0', 5-len(ltrim(rtrim(do.nclocdest)))) + ltrim(rtrim(cast(do.nclocdest as varchar))) as CdLocalidadDestino, "&_             
             " replicate('0', 4-len(ltrim(rtrim(do.kmrecorrer)))) + ltrim(rtrim(cast(do.kmrecorrer as varchar))) as KmRecorrer, "&_             
             " ltrim(rtrim(cc.cdchapacamion)) + replicate(' ', 11-len(ltrim(rtrim(cc.cdchapacamion)))) as PatenteCamion, "&_       
             " ltrim(rtrim(cc.cdchapaacoplado)) + replicate(' ', 11-len(ltrim(rtrim(cc.cdchapaacoplado)))) as PatenteAcoplado, "&_       
             " replace(replicate('0', 8-len(do.fitarifa)) + cast(do.fitarifa as varchar),'.',',') as TarifaTonelada, "&_
             " case when replace(replicate('0', 8-len(do.tarifaref)) + cast(do.tarifaref as varchar),'.',',') is null then '00000,00' else replace(replicate('0', 8-len(do.tarifaref)) + cast(do.tarifaref as varchar),'.',',') end as TarifaReferencia "&_             			 
			 " from datosoncca do inner join (select hc.DTCONTABLE,hc.IDCAMION,CDCHAPACAMION,CDCHAPAACOPLADO,CDTIPOCAMION,CDTRANSPORTISTA,DSNOMBRECONDUCTOR, "&_
			 " DSAPELLIDOCONDUCTOR,CDTIPODOC,NUDOCUMENTO,DTINGRESO,DTEGRESO,CDESTADO,CDCIRCUITO,CDFILA,CDSILO,CDPLATAFORMA,ICTRANSMITIDO,SQCAMION, "&_
			 " CDPRODUCTO,NUAUTSALIDA,SQTURNO,ICCUPO,DSCUPO,ICCONTRATOESP,IDCUPOASIGNADO,NUCUITREM,NUCUPO,CDEMPRESA,CDCLIENTE,CDCORREDOR,CDDESTINATARIO,CDVENDEDOR, "&_
			 " CDENTREGADOR,CDDESTINO,CDMOTIVOEGRESO,CDCOSECHA,ICINSPECCION,ICREINSPECCION,ICRECHAZO,VLKGSACARGAR,NUINFOCARGA,NUCARTAPORTE,NURECIBO,NUCTAPTEDIG "&_
			 " from hcamiones hc inner join hcamionescarga hcc on hc.idcamion = hcc.idcamion and hc.dtcontable = hcc.dtcontable)cc on do.nccartaporte = cc.nucartaporte "&_
			 myWhere
			 
			'logMig.info(strSQL)   

	Call executeQueryDb(pPto, rs, "OPEN", strSQL)	
	

	Set obtenerCartasPorteEmitidas = rs
End Function
'---------------------------------------------------------------------------------------------------------
Function armarregistroDatos(myRs, myFile, pto)
    
	Dim myArrayAux
    
    On Error Resume Next    
    
    logMig.info("Archivo listo para trabajar.")        
	myArrayAux = myRs.GetRows		
	registro = ""
	For row = 0 To UBound(myArrayAux, 2)
		For col = 0 To UBound(myArrayAux, 1) 
			registro = registro & CStr(myArrayAux(col, row))	
		Next		
		if registro <> "" then myFile.WriteLine registro
		registro = ""		
	Next		
	logMig.info("Finalizado armado de registros para puerto " & pto)	
End Function
'---------------------------------------------------------------------------------------------------------
Function verificarIntegridadCartasPorte(rsCartasPorte, fecha, pto, listaErrores)
Dim sinErrores, myStrErrores
sinErrores = true
myStrErrores = ""
	if (not rsCartasPorte.eof) then
		while (not rsCartasPorte.Eof)
		if (isNull(rsCartasPorte("NroCartaPorte"))) then
			myStrErrores = myStrErrores & "Carta de Porte: sin datos ONCCA " & vbCrLf 
			sinErrores = false
		else
		
			if (len(Trim(rsCartasPorte("NroCEE"))) <> 14) and (IsNumeric(rsCartasPorte("NroCEE"))) then 
				myStrErrores = myStrErrores & "Numero CEE: " & CStr(rsCartasPorte("NroCEE")) & ", longitud incorrecta " & vbCrLf 				
				sinErrores = false
			end if			
			if (isNull(rsCartasPorte("CodigoEspecie"))) then
				myStrErrores = myStrErrores & "Codigo Producto: incorrecto o inexistente para datos oncca " & vbCrLf 				
				sinErrores = false
			end if
			if (isNull(rsCartasPorte("CuitEstablDestino"))) or (rsCartasPorte("CuitEstablDestino") = "00000000000") then
				myStrErrores = myStrErrores & "Cuit Establecimiento Destino: incorrecto o inexistente para datos oncca " & vbCrLf 				
				sinErrores = false
			end if
			if (isNull(rsCartasPorte("PesoNeto"))) then
				myStrErrores = myStrErrores & "Peso Neto: incorrecto. Notificar al personal de sistemas " & vbCrLf 				
				sinErrores = false
			end if			
			'Controles propios de Camiones
			if (CInt(rsCartasPorte("TipoTransporte")) = TIPO_TRANSPORTE_CAMION) then
				if isNull(rsCartasPorte("TarifaReferencia")) then 
					myStrErrores = myStrErrores & "Tarifa de Referencia: valor incorrecto " & vbCrLf 				
					sinErrores = false
				elseif (CDbl(replace(rsCartasPorte("TarifaReferencia"),",",".")) <= 0) or (CDbl(replace(rsCartasPorte("TarifaReferencia"),",",".")) > 5000) then
					myStrErrores = myStrErrores & "Tarifa de Referencia: " & CStr(rsCartasPorte("TarifaReferencia")) & ", valor fuera de rango " & vbCrLf 				
					sinErrores = false
				end if
				if isNull(rsCartasPorte("TarifaTonelada")) then 
					myStrErrores = myStrErrores & "Tarifa por Tonelada: valor incorrecto " &  vbCrLf 				
					sinErrores = false
				elseif (CDbl(replace(rsCartasPorte("TarifaTonelada"),",",".")) <= 0) or (CDbl(replace(rsCartasPorte("TarifaTonelada"),",",".")) > 5000) then
					myStrErrores = myStrErrores & " Tarifa por Tonelada: " & CStr(rsCartasPorte("TarifaTonelada")) & ", valor fuera de rango " & vbCrLf 				
					sinErrores = false
				end if
				if isNull(rsCartasPorte("KmRecorrer")) then 
					myStrErrores = myStrErrores & "Kilometros a Recorrer: valor incorrecto " &  vbCrLf 				
					sinErrores = false
				elseif (CInt(rsCartasPorte("KmRecorrer")) <= 0)then
					myStrErrores = myStrErrores & " Kilometros a Recorrer: debe ser mayor a cero " & vbCrLf 				
					sinErrores = false
				end if
			end if
		end if
			
			'logMig.Info("salida de control: " & myStrErrores)
			if myStrErrores <> "" then listaErrores.Add CStr(rsCartasPorte("NroCartaPorte")), myStrErrores			
			myStrErrores = ""
			
			rsCartasPorte.MoveNext()
		wend
	else
		logMig.info ("algo salio mal, recordSet en verificarIntegridadDatos esta vacio")
	end if

verificarIntegridadCartasPorte = sinErrores
End function
'---------------------------------------------------------------------------------------------------------
Function enviarMailsCartasPorte(pFecha, pFileAttachment, listaErroresArroyo, listaErroresTransito,listaErroresPiedraBuena)
    Dim strBody, strSubject,fs
    strBody = ""    
    strSubject = "Cartas de Porte Emitidas de " & GF_FN2DTE(pFecha) 
    	
	Set fsAux = Server.CreateObject("Scripting.FileSystemObject")		
    if (fsAux.FileExists(pFileAttachment) and fsAux.GetFile(pFileAttachment).Size > 0) then    
        strBody = strBody & enviarMailConErroresParaPuerto(TERMINAL_ARROYO, pFecha,listaErroresArroyo)
		strBody = strBody & enviarMailConErroresParaPuerto(TERMINAL_TRANSITO, pFecha,listaErroresTransito)
		strBody = strBody & enviarMailConErroresParaPuerto(TERMINAL_PIEDRABUENA, pFecha,listaErroresPiedraBuena)
		if strBody = "" then						
			strBody = "Se envia adjunto el archivo con las Cartas de Porte" & vbCrLf & vbCrLf & strBody 
			logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )
			Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_INFO_LIST, strSubject, strBody, pFileAttachment)
			'solo se actualiza la fecha de ultima ejecucion del archivo .lck, si la fecha no se ha pasado por parametro
			if not flagForzarEjecucion then flagActualizarFechaUltimaEjecucion = true
		end if
    else					
		strSubject = "Sin Cartas de Porte Emitidas de " & GF_FN2DTE(pFecha) 
		strBody = "No se registraron Cartas de Porte Emitidas para la fecha " & GF_FN2DTE(pFecha) & "."
		logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )
		Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_INFO_LIST, strSubject, strBody, "")
		'solo se actualiza la fecha de ultima ejecucion del archivo .lck, si la fecha no se ha pasado por parametro			
		if not flagForzarEjecucion then flagActualizarFechaUltimaEjecucion = true
    end if
	Set fsAux = Nothing
End Function
'---------------------------------------------------------------------------------------------------------
Function enviarMailConErroresParaPuerto(pto, pFecha,listaErrores)
	Dim strBody, strSubject, myLista, letraPuerto
    strBody = ""
	       
    if listaErrores.Count > 0 then
		letraPuerto = getLetraPuerto(pto)
		logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )    
		strSubject = "Cartas de Porte Emitidas con errores de " & GF_FN2DTE(pFecha) & " " & pto 
		strBody = strBody & GF_FN2DTE(pFecha) & vbCrLf & "Cartas Porte con errores en " & pto & vbCrLf & vbCrLf
		logMig.info(strBody & " " & strSubject)
		for each cartaPorteConError in listaErrores.Keys
			strBody = strBody & vbTab & "Nro CCPP : " & cartaPorteConError & vbCrLf & vbTab & vbTab & listaErrores(cartaPorteConError)
		Next
		strBody = strBody & vbCrLf
		logMig.info(strBody & " " & strSubject)
		Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_ERROR_LIST & letraPuerto, strSubject, strBody, "")		
	end if
	enviarMailConErroresParaPuerto = strBody
end function
'---------------------------------------------------------------------------------------------------------
Function getLastExecDate()
    Dim fso, myFile, myFilename, myData
    
    myFilename = server.MapPath(".") & "\servicioEnvioCartasPorteEmitidas.lck"
	'logMig.info("nombre archivo: " & myFilename)
    Set fso = CreateObject("scripting.filesystemobject")        
	if (fso.FileExists(myFilename)) then
		'logMig.info("archivo existe " & myFilename)
	    Set myFile = fso.OpenTextFile(myFilename,1,false)
        if (not myFile.AtEndOfStream) then
	        getLastExecDate = myfile.ReadLine	               
        else
            getLastExecDate = ""
        end if	        
	    myFile.Close
	    Set myFile = Nothing
	else
		logMig.info("archivo NO existe " & myFilename)
	    getLastExecDate = ""
	end if
	Set fso = Nothing
End Function
'---------------------------------------------------------------------------------------------------------
Function updateLastExecDate(pData)
    Dim fso, myFile, myFilename    
    myFilename = server.MapPath(".") & "\servicioEnvioCartasPorteEmitidas.lck"
    Set fso = CreateObject("scripting.filesystemobject")    
	if (fso.FileExists(myFilename)) then fso.DeleteFile(myFilename)
    Set myFile = fso.CreateTextFile(myFilename, true)
    myFile.WriteLine(pData)
    myFile.Close
	Set myFile = Nothing
	Set fso = Nothing	
End Function
'---------------------------------------------------------------------------------------------------------
Function procesarCartasPorte (rsCartasPorte, fecha, pto,listaErrores, pFile)
	if (not rsCartasPorte.eof) then
			if verificarIntegridadCartasPorte(rsCartasPorte, fecha, pto, listaErrores) then
				rsCartasPorte.MoveFirst()								
				Call armarRegistroDatos(rsCartasPorte, pFile, pto)
			else
				logMig.info("Cartas Porte con errores en " & pto)
				for each cartaPorteConError in listaErrores.Keys
					logMig.info("Nro CCPP : " & cartaPorteConError & vbCrLf & listaErrores(cartaPorteConError))
				Next
			end if
		else
			logMig.info("No se encontraron Cartas de Porte para exportar para puerto " & pto & " en la fecha " & fechaContable)
		end if
End function
'---------------------------------------------------------------------------------------------------------
'                                   ***** COMIENZA PAGINA *****
'---------------------------------------------------------------------------------------------------------

Dim fecha, fechaContable, logMig, strSQL, registro, myFilename, g_strPuerto
Dim rsCartasPorteArroyo, rsCartasPorteTransito, rsCartasPortePiedraBuena
Dim listaErroresArroyo,listaErroresTransito,listaErroresPiedraBuena
Dim fs, myFile, errMsg, myCartaPorte
Dim myArrayAux
Dim flagActualizarFechaUltimaEjecucion, flagForzarEjecucion
set listaErroresArroyo = Server.CreateObject("Scripting.Dictionary")
set listaErroresTransito = Server.CreateObject("Scripting.Dictionary")
set listaErroresPiedraBuena = Server.CreateObject("Scripting.Dictionary")

flagActualizarFechaUltimaEjecucion = false
flagForzarEjecucion = false

fecha = GF_PARAMETROS7("fd", "", 6)

Call GP_ConfigurarMomentos()
session("usuario") = "SYNC"


Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "EXPORTACION_CARTASPORTE_EMITIDAS" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)

if (fecha = "") then
	fecha=getLastExecDate()
	if fecha = "" then
		'fecha = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")		
		logMig.info("Error en tomar la fecha del archivo auxiliar de fecha servicioEnvioCartasPorteRecibidas.lck")
	else		
		logMig.info("Fecha tomada del archivo auxiliar de fecha servicioEnvioCartasPorteRecibidas.lck " & fecha)
		'fecha = GF_DTEADD(fecha, 1, "D")
		'logMig.info("Fecha actualizada " & fecha)
	end if
else
	'se ha pasado fecha por parametro, se debe ejecutar proceso sin actualizacion de fecha
	flagForzarEjecucion = true
end if


if flagForzarEjecucion or (fecha < Left(session("MmtoDato"),8)) then
'se ejecuta el proceso unicamente: Si se ha pasado fecha como parametro. O si la fecha de ultima ejecucion es menor a de hoy
	fechaContable = GF_FN2DTCONTABLE(fecha)
	myFilename = server.MapPath(".\Temp") & "\CARTASPORTE_EMITIDAS" & fecha & ".txt"

	logMig.info("------------ INCIANDO EXPORTACION CARTAS DE PORTE EMITIDAS------------------")	
	logMig.info(" ---> FECHA : " & fechaContable)
	logMig.info("-----------------------------------------------------------------------------")	



		
		'On Error Resume Next    
		
		'Se abre el archivo de datos    
		logMig.info("Inicializando archivo de datos: " & pFilename)
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		If fs.FileExists(myFilename) Then  Call fs.deleteFile(myFilename, true)
		Set myFile = fs.OpenTextFile(myFilename, 2, true)
		logMig.info("Archivo listo para trabajar.") 
		

		Set rsCartasPorteArroyo = obtenerCartasPorteEmitidas(fechaContable, TERMINAL_ARROYO)
		Set rsCartasPorteTransito = obtenerCartasPorteEmitidas(fechaContable, TERMINAL_TRANSITO)
		Set rsCartasPortePiedraBuena = obtenerCartasPorteEmitidas(fechaContable, TERMINAL_PIEDRABUENA)	
		
		Call procesarCartasPorte(rsCartasPorteArroyo, fecha, TERMINAL_ARROYO,listaErroresArroyo, myFile)
		Call procesarCartasPorte(rsCartasPorteTransito, fecha, TERMINAL_TRANSITO,listaErroresTransito, myFile)
		Call procesarCartasPorte(rsCartasPortePiedraBuena, fecha, TERMINAL_PIEDRABUENA,listaErroresPiedraBuena, myFile)
		

		myFile.Close()
		
		Set myFile = Nothing
		Set fs = Nothing
			
		Call enviarMailsCartasPorte(fecha, myFilename, listaErroresArroyo, listaErroresTransito,listaErroresPiedraBuena)
		'solo actualiza la fecha de ultima ejecucion en el archivo auxiliar si:
		'1: la ejecucion es programada y la fecha no se ha pasado por parametro, sino se la ha tomado del archivo auxiliar
		'2. si no hubo errores en la verificacion de datos.
		'de este modo la proxima ejecucion programada se va a ejecutar sobre la misma fecha, hasta que no haya errores.
		if flagActualizarFechaUltimaEjecucion then 
			fecha = GF_DTEADD(fecha, 1, "D")
			logMig.info("Fecha actualizada " & fecha)
			updateLastExecDate(fecha)
		end if
		

	logMig.info("--------------------------- FIN PROCESO ---------------------------")	
else
	logMig.info("No hay mas Cartas de Porte para procesar. La fecha de ultima ejecucion exitosa es igual a la actual")
end if
%>