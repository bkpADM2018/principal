<!--#include file="interfacturasPrintCommon.asp"-->
<!--#include file="../includes/procedimientosSeguridad.asp"-->
<!--#include file="../includes/procedimientosLog.asp"-->
<!--#include file="../includes/procedimientosMail.asp"-->
<%

'-------------------------------------------------------------------------------------------------------------------
'Obtiene todos los numeros de comprobantes de una determinada lista de numero de registro
Function getComprobantesByNroReg(pListReg)
	Dim strSQL,rs,strCbte	
	strSQL = "SELECT * FROM TFFL.TF100F1 WHERE FCRGNR IN ("&pListReg&")"
	Call executeQuery(rs, "OPEN", strSQL)
	strCbte = ""	
	while not rs.Eof		
		strCbte = strCbte & GF_nDigits(rs("FCCMDV"),4) &"-"& GF_nDigits(rs("FCCMNR"),8) & vbcrlf
		rs.MoveNext()
	wend 
	'strCbte = left(strCbte,Len(strCbte)-1)
	getComprobantesByNroReg = strCbte
End Function
'-------------------------------------------------------------------------------------------------------------------
Function actualizarMmtoEnvioFactura(pList,pDestino)
	Dim strSQL 
	strSQL = "UPDATE TFFL.TF900 SET FLMLMM = "& session("MmtoDato") &", FLURLM = '"& Left(pDestino, 100) &"' WHERE FLRGNR IN("& pList &")"
	Call executeQuery(rsUpd, "EXEC", strSQL)
	Call logMig.info("Se actualizo el registros: " & pList & " en el TFFL.TF900")
End Function
'-------------------------------------------------------------------------------------------------------------------
Function getMailFacturacionProveedores(pIdProveedor)
	Dim auxDestino,rsMail
	auxDestino = ""
	if (pIdProveedor <> "") then	    
		Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsMail, "TBLTAREAMAILS_GET_BY_PARAMETERS", TASK_FAC_MAIL_ALERT & "||" & idProveedor )
	    while (not rsMail.Eof)
		    auxDestino = auxDestino & Trim(rsMail("MAIL")) & "; "
		    rsMail.MoveNext()
	    wend
	    if (Len(auxDestino) > 0) then auxDestino = Left(auxDestino, Len(auxDestino)-2)
    end if	    
	getMailFacturacionProveedores = auxDestino
End Function
'-------------------------------------------------------------------------------------------------------------------
Function getAsunto(pAccion)
    if (pAccion = ACCION_CONFIRMAR) then
        getAsunto = getDescripcionProveedor(CD_TOEPFER) & " - Facturación - Aviso de Vencimiento"
    else
        getAsunto = GF_TRADUCIR(getDescripcionProveedor(CD_TOEPFER) & " - Facturación")
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Function getCuerpoMail(pAccion, pProveedor, pListReg)
	    if (pAccion = ACCION_CONFIRMAR) then
	        getCuerpoMail = "Señor/es: "& vbcrlf & getDescripcionProveedor(pProveedor) & vbcrlf & vbcrlf &_
					 "Nos dirijimos a Usted a efectos de comunicarle que el/los siguientes comprobantes está próximos a vencer: " & vbcrlf & vbcrlf & pListReg & vbcrlf & vbcrlf &_
					 "RECUERDE: La mora en el pago devengará a favor de Toepfer un interes punitorio equivalente a la tasa cobrada por el Banco de la Nación Argentina para adelantos en cuenta corriente no solicitados." & vbcrlf & vbcrlf &_
					 "Atentamente."&vbcrlf&vbcrlf&_
					 "Departamento de Tesoreria"&vbcrlf& getDescripcionProveedor(CD_TOEPFER)&vbcrlf&"Tel (011) 4317-0000"    
	    else
              getCuerpoMail = "Señor/es: "& vbcrlf & getDescripcionProveedor(pProveedor) & vbcrlf & vbcrlf &_
	             "Nos dirijimos a Usted a efectos de hacerle llegar adjuntos los comprobantes:" & vbcrlf & vbcrlf & pListReg & vbcrlf & vbcrlf &_					 
	             "Por favor confirmar recepcion de este mail." &vbcrlf&vbcrlf&_
	             "Los comprobantes adjuntos son confeccionados en cumplimiento de lo estipulado por la AFIP en su Resolucion General 3571/12. " & vbcrlf &_
	             "El contenido del presente mensaje es privado, estrictamente confidencial y exclusivo para sus destinatario, pudiendo "  & vbcrlf &_
	             "contener informacion protegida por normas legales y de secreto profesional. Bajo ninguna circunstancia su contenido "  & vbcrlf &_
	             "puede ser transmitido o relevado a terceros ni divulgado en forma alguna. En consecuencia de haberlo recibido por error "  & vbcrlf &_
	             "solicitamos contactar al remitente y eliminarlo de su sistema. "&vbcrlf&vbcrlf&_
	             "Atentamente."&vbcrlf&vbcrlf&_
	             "Departamento de Tesoreria"&vbcrlf& getDescripcionProveedor(CD_TOEPFER)&vbcrlf&"Tel (011) 4317-0000"    
        end if	             
    End Function
'-------------------------------------------------------------------------------------------------------------------
Function enviarMailByProveedorLista(pAccion, pIdProveedor,pFileAtt,pListReg)
	Dim auxDestino,auxAsunto,auxProveedor,rs,rsMail, myLista, mySender
	auxProveedor = pIdProveedor
	'Obtengo las direcciones de mail del Proveedor-Lista (puede que hallan mas de un mail para ambos registros)	
	auxDestino = getMailFacturacionProveedores(auxProveedor)
	
	auxAsunto  = getAsunto(pAccion)
	auxMensaje = getCuerpoMail(pAccion, auxProveedor, pListReg)
	
	'En caso de que el proveedor no figura en la tabla de Mail, se obtiene el correo interno notificando el error
	if (Trim(auxDestino) = "") then
	    auxAsunto  = GF_TRADUCIR("FACTURACION - El proveedor no tiene direcciones definidas: " & getDescripcionProveedor(pIdProveedor))									 
	    auxMensaje = "Hay facturas que presentarle al proveedor " & pIdProveedor & "-" & getDescripcionProveedor(pIdProveedor) & vbcrlf
	    auxMensaje = auxMensaje & " que no pudieron ser enviadas ya que no tenemos mails del proveedor. Se adjuntan los comprobantes."
	    Call logMig.info("El proveedor no tiene direcciones definidas, se le envía el mail a Tesorería.")		
    else
        auxDestino = auxDestino & ";"	    
	end if	
	auxDestino = auxDestino & getMailFacturacionProveedores(CD_TOEPFER)
	mySender = getMailFacturacionProveedores(MAIL_TASK_SENDER)
	auxDestino = "scalisij@toepfer.com"		
	Call GP_ENVIAR_MAIL_ATTACHMENT(auxAsunto, auxMensaje, mySender, auxDestino, pFileAtt)
	Call logMig.info("Mail Enviado a: " & auxDestino)	
	'if (pAccion <> ACCION_CONFIRMAR) then Call actualizarMmtoEnvioFactura(listNroRegistro, auxDestino)
	
End Function
'-------------------------------------------------------------------------------------------------------------------
Function enviarMailError(pIdProveedor, pRegs, pFile)
    Dim strMsg, auxDestino, mySender
    
    strMsg = "Error al intentar enviar un archivo de facturas." & vbcrlf &_
             "Datos:" & vbcrlf &_
             "  Proveedor   : " & pIdProveedor & vbcrlf &_  
             "  Archivo     : " & pFile & vbcrlf &_
             "  Lista Mail  : " & pIdProveedor & vbcrlf &_
             "  Registros   : " & pRegs
    auxDestino = getMailFacturacionProveedores(FACTURACION_LISTA_MAIL_ERROR)    
	mySender = getMailFacturacionProveedores(MAIL_TASK_SENDER)
    Call GP_ENVIAR_MAIL("Sistema de Facturacion - ERROR", strMsg, mySender, auxDestino)
    Call logMig.info(strMsg)
End Function
'******************************************************************************************
'										INICIO DE PAGINA
'******************************************************************************************
Dim listNroRegistro,logMig,rsMail,auxFileAtt,fso, myNroReg, flagMAT, listCbtes, action
Dim strSQL

myNroReg = GF_Parametros7("registro", "",6)
action = GF_Parametros7("accion", "",6)

Call GP_ConfigurarMomentos

if (session("Usuario") = "") then session("Usuario") = "JAS"

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "FACTURACION-ELECT-"& left(session("MmtoDato"),8)
Call logMig.info("--------------------------------------------------------------")
'Controlo el origen de la llamada a la pagina
if (action = ACCION_CONFIRMAR) then	
    'fechaVto = GF_DTEADD(Left(session("MmtoDato"), 8), 2, "D")
    'dteVto = GF_FN2DTCONTABLE(fechaVto)
    'while (Weekday(dteVto, 2) >= 6)
    '    fechaVto = GF_DTEADD(fechaVto, 2, "D")
    '    dteVto = GF_FN2DTCONTABLE(fechaVto)
    'wend
    'Call logMig.info("Buscando Comprobante con vencimiento " & GF_FN2DTE(fechaVto))
	'strSQL="Select * from FAC001A where [date] = '" & dteVto & "'"    
else
    Call logMig.info("Verificando si hay comprobantes para enviar....")
	strSQL="Select YEAR(feccbt)*10000 + Month(feccbt)*100 + Day(feccbt) ctrlfecha, * from FAC001A where "
    'if (myNroReg <> "") then 
		strSQL= strSQL & " [guid] = '" & myNroReg & "'"
	'else
	'	strSQL= strSQL & " ndjve = '' "
	'end if
end if    
Call executeQueryDb(DBSITE_SQL_MAGIC, rsPendiente, "OPEN", strSQL)
'Primero obtengo todos los registros de las facturas que estan procesadas.
if (not rsPendiente.eof) then
    while not rsPendiente.Eof
	    idProveedor =  getCUITProveedor(rsPendiente("cliente"))
	    logMig.info(" - PROVEEDOR ID  : " & rsPendiente("cliente"))
		logMig.info(" - PROVEEDOR CUIT: " & idProveedor)
	    flagMAT = false
	    listNroRegistro = ""
	    listCbtes = ""
		myFile = ""
	    'while (idProveedor = cdbl(rsPendiente("cliente"))) 
	        'Guardo el registro en la lista del proveedor y del sector.
			if (CLng(rsPendiente("ctrlfecha")) > 20170914) then 
				listNroRegistro = listNroRegistro & rsPendiente("GUID") & "', '"
			else
				myFile2 = myFile2 & Server.MapPath(".") & "\archive\" & Right(rsPendiente("GUID"), 8) & ".pdf;"
			end if
		    'Guardo el registro en la lista del proveedor y del sector.
	        auxTipo = "Factura"
	        if (CInt(rsPendiente("tipcbt")) = 2 ) then auxTipo= "N. Debito"
			if (CInt(rsPendiente("tipcbt")) = 3 ) then auxTipo= "N. Credito"
		    listCbtes = listCbtes & auxTipo & " " & GF_nDigits(rsPendiente("succbt"), 4) & "-" & GF_nDigits(rsPendiente("nrocbt"), 8) & ", vto: " & GF_FN2DTE(rsPendiente("fecvto")) & ", Importe: " & getSimboloMoneda(rsPendiente("codmone")) & " " & GF_EDIT_DECIMALS(CDbl(rsPendiente("imptotcbt"))*100, 2) & vbcrlf & vbcrlf
		    logMig.info(" - " & listCbtes)
		    'Para el MAT se debe incluir solo una factura por archivo.
		    'if (Cdbl(rsPendiente("cliente")) = PROVEEDOR_ESPECIAL_MAT) then flagMAT = true
		    rsPendiente.MoveNext()
	    'wend
		if (myFile2 <> "") then myFile2 = Left(myFile2, Len(myFile2)-1)
	    if (Len(listNroRegistro) > 0) then
		    listNroRegistro = left(listNroRegistro,Len(listNroRegistro)-4)
		    logMig.info(" - NRO. DE REGISTROS: " & listNroRegistro)		    
		    'Se genera la factura.
            myFile = generarPDF(listNroRegistro,PDF_FILE_MODE) 
			myFile = Server.MapPath("..\temp\" & myFile)
        end if
		if (myFile2 <> "") then myFile = myFile2 & myFile
		if (myFile <> "") then            
            Call logMig.info("	Archivo  : " & myFile)
            Set fso = CreateObject("Scripting.FileSystemObject")
            rtrn = myFile
            if (fso.FileExists(rtrn)) then
                Call enviarMailByProveedorLista(action, idProveedor,rtrn, listCbtes)                
                Set fso = CreateObject("Scripting.FileSystemObject")
                fso.DeleteFile(rtrn)
            else
                Call logMig.info("No se encuentra el archivo a enviar: " & rtrn)
                Call enviarMailError(idproveedor, registro, rtrn)
            end if
			
	    end if
    wend                 

else
    Call logMig.info("No hay comprobantes.")
end if
%>