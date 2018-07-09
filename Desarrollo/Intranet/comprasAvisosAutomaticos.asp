 <!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosRoles.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->

<%
Const DIV_EXPORTACION = 1
Const DIV_ARROYO      = 2
Const DIV_PIEDRABUENA = 3
Const DIV_TRANSITO    = 4
'Alertas Vto Minutas - Legales
'Codigo de cuenta relacionadas a Seguros
Const COD_CTA_MINUTA_SEGUROS_1 = "11339102"
Const COD_CTA_MINUTA_SEGUROS_2 = "11334603"
Const COD_CTA_MINUTA_SEGUROS_3 = "11353001" 
'************************************************************
'	Pagina que se ejecuta via una tarea programada de manera
'	de actualizar el estado de todos los proyectos en curso.
'	Esto sirve para mantener el sistema actualizado en caso 
'	de que nadie ingrese al sistema y se cumplan plazos 
'	específicos.
'************************************************************
'------------------------------------------------------------
Function enviarMail(pTitulo, pSender, pEmail,pMensaje) 
	Dim rtrn
	rtrn=false	
	'if ((InStr(pEmail, "AlemanM@Toepfer.com") = 0) and (InStr(pEmail, "GalarzaP@Toepfer.com") = 0)) then
		if (pSender <> "") and (pEmail <> "") then		
			Call myLog.info("Mail enviado a: " & pEmail)
			'mail_config_Type = MAIL_TYPE_HTML
			Call GP_ENVIAR_MAIL(pTitulo, pMensaje, pSender, pEmail)		
			rtrn = true
		end if	
	'end if
	enviarMail = rtrn
End Function
'------------------------------------------------------------
Function actualizarEstadoPedidos()
	Call myLog.info("##########################################")
	Call myLog.info("ACTUALIZACION DE ESTADO DE PEDIDOS")
	Call myLog.info("##########################################")
	Dim myWhere, strSQL, conn, rs,rs2
	Call GP_ConfigurarMomentos
	
	Call mkWhere(myWhere, "ESTADO", ESTADO_PCT_PUBLICADO, ">=", SQL_WHERE_INTEGER)
	Call mkWhere(myWhere, "ESTADO", ESTADO_PCT_ABIERTO, "<=", SQL_WHERE_INTEGER)
	strSQL = "Select * from TBLPCTCABECERA " & myWhere
	'response.write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	while (not rs.eof)
		Call myLog.info("Actualizando pedido ID=" & rs("IDPEDIDO") & " (" & rs("CDPEDIDO") & ")") 
		Call initHeader(rs("IDPEDIDO"))		
		Call actualizarExtensible()
		Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "OPEN", "Select * from TBLPCTCABECERA WHERE idpedido =" & rs("idpedido"))
		Call actualizarEstado(rs2)	
		Call myLog.info("Pedido Actualizado.")
		rs.MoveNext()
	wend	
	Call myLog.info("Proceso terminado")

End Function
'------------------------------------------------------------
Function comprobarPedidosAVencer()
	Dim strSQL, strSQL2, rsADMINS, rs2, conn, conn2, mensaje, cdAdmin, mailAdmin
	Call GP_ConfigurarMomentos()
	Call myLog.info("##########################################")
	Call myLog.info("COMPROBACION DE PEDIDOS")
	Call myLog.info("##########################################")
	Call myLog.info("Obteniendo Pedidos proximos a cerrar.")
	'OBTENGO LA LISTA DE PERSONAS QUE ADMINISTRAN PEDIDOS QUE CIERRAN HOY O MAÑANA
	strSQL = "SELECT DISTINCT CDSOLICITANTE FROM TBLPCTCABECERA WHERE FECHACIERRE = " & left(session("MmtoSistema"),8) & " or FECHACIERRE= " & GF_DTEADD(left(session("MmtoSistema"),8),1,"D")
    Call executeQueryDb(DBSITE_SQL_INTRA, rsADMINS, "OPEN", strSQL)
	if (not rsADMINS.eof) then
		Call myLog.info("Enviando mails de notificaciones..." )
	else
		Call myLog.info("No se encontraron pedidos a cerrar el día de mañana" )
	end if
	while not rsADMINS.eof		
		cdAdmin = Trim(rsADMINS("CDSOLICITANTE"))
		Call myLog.info("Procesando Pedidos de:" & cdAdmin)
		mailAdmin = SENDER_LICITACIONES & "; " & MAILTO_COMPRAS 'getUserMail(cdAdmin)	
		strSQL2 = "SELECT * FROM TBLPCTCABECERA WHERE CDSOLICITANTE ='" & cdAdmin & "' AND FECHACIERRE = " & left(session("MmtoSistema"),8) & " or FECHACIERRE= " & GF_DTEADD(left(session("MmtoSistema"),8),1,"D")
		Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL2)
		mensaje = "Los siguientes pedidos estan próximos a finalizar:" & vbcrlf & vbcrlf
		while not rs2.eof
			mensaje = mensaje & rs2("CDPEDIDO") & " - " & rs2("TITULO") & vbcrlf & vbcrlf
			rs2.movenext
		wend
		Call enviarMail("Sistema de Compras Web - Alerta pedidos proximos a cerrar", MAILTO_COMPRAS, mailAdmin ,Mensaje)
		rsADMINS.movenext
	wend	
	Call myLog.info("Proceso terminado")
	
End Function
'------------------------------------------------------------------------------------------------
' Función:	controlarIdentificacionUsuario
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	--/--/----
' Objetivo:	
'			Definir los usuarios que tienen documentos para firmar, a medida que no sean duplicados los va guardando
'			en el objeto Dictionay			
' Parametros:
'			pcdUsuraio 	[string], pidAlmacen   [int],	 pTipo [string]  	
' Devuelve:
'			-
' Modificacion: 
'			27/11/2012	
'--------------------------------------------------------------------------------------------
Function controlarIdentificacionUsuario(pcdUsuario,pidAlmacen,pTipo,pPto)
	Dim strSQL, pDivision, cdUsuario,rtrn,myWhere	
	
	'Si está pendiente la firma de un vale de ajuste o reclasificacion (o sus anulaciones), para un usuario generico se analiza quienes tiene los permisos para firmarlo.
	if((pcdUsuario = VS_NO_USER)or(pcdUsuario = VS_AUDIT_USER)or(pcdUsuario = VS_PORT_SUPERVISOR_USER)or(pcdUsuario = VS_PORT_GERENTE_USER)or(pcdUsuario = DIRECTOR_USER)or(pcdUsuario = CONTROLLER_USER))then
	    'DEPENDIENDO DEL TIPO DOC. OBTENGO EL ROL EQUIVALENTE PARA PODER UTILIZAR LA TABLA REGISTRO FIRMAS	
	    myWhere = "rolfirma ="
		select case pTipo
			case AUTH_TYPE_AJS, AUTH_TYPE_XJS
				Select case pcdUsuario
					case VS_AUDIT_USER: 
					' AUDITOR
						myWhere = myWhere & FIRMA_ROL_AUDITOR
					case DIRECTOR_USER: 
					' DIRECTOR
						myWhere = myWhere & FIRMA_ROL_DIRECTOR
					case VS_PORT_SUPERVISOR_USER: 
					'COORDINADOR DE PUERTOS
						myWhere = myWhere & FIRMA_ROL_SUP_PUERTO
					case VS_NO_USER 
					'GERENTE DE PUERTO
						myWhere = myWhere & FIRMA_ROL_RESP_PUERTO
						' PROCESO LOS USUARIOS QUE HACEN REFERENCIA A UN GRRUPO DE ROLES
		                strSQL = " SELECT iddivision FROM tblalmacenes WHERE idalmacen = " & pidAlmacen
		                Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		                ' A TRAVES DE SU ALMACEN OBTENGO LA DIVISION 
		                if(not rs.eof)then
			                select case rs("IDDIVISION")
				                case DIV_ARROYO:
					                myWhere = myWhere & " AND AJARROYO = 1 "
				                case DIV_TRANSITO:				
					                myWhere = myWhere & " AND AJTRANSITO = 1 "
				                case DIV_PIEDRABUENA:				
					                myWhere = myWhere & " AND AJPIEDRABUENA = 1 "
				                case DIV_EXPORTACION:
					                myWhere = myWhere & " AND AJEXPORTACION = 1 "
			                end select
		                end if
				End Select
			case AUTH_TYPE_VRS, AUTH_TYPE_XRS: 
			' SUPERVISOR
				myWhere = myWhere &  FIRMA_ROL_SUP_PUERTO
			case AUTH_TYPE_AJD, AUTH_TYPE_AJC, AUTH_TYPE_AJM:
				Select case pcdUsuario
					case VS_PORT_GERENTE_USER:
						myWhere = myWhere & FIRMA_ROL_RESP_PUERTO
					case CONTROLLER_USER:
						myWhere = myWhere & FIRMA_ROL_CONTROLLER
					case DIRECTOR_USER:
						myWhere = myWhere & FIRMA_ROL_DIRECTOR
				End Select				
				select case Ucase(pPto)
	                case TERMINAL_ARROYO:						
		                myWhere = myWhere & " AND AJPTOARROYO = 1 "		                
	                case TERMINAL_TRANSITO:				
		                myWhere = myWhere & " AND AJPTOTRANSITO = 1 "		                
	                case TERMINAL_PIEDRABUENA:				
		                myWhere = myWhere & " AND AJPTOPIEDRABUENA = 1 "	                
                end select
            case AUTH_TYPE_PIC, AUTH_TYPE_AIC:
				Select case pcdUsuario
					case CTZ_NO_USER: 
					' SUP. DE PUERTOS
						myWhere = myWhere & FIRMA_ROL_SUP_PUERTO
					case DIRECTOR_USER: 
					' DIRECTOR
						myWhere = myWhere & FIRMA_ROL_DIRECTOR					
				End Select
			case AUTH_TYPE_PCP, AUTH_TYPE_AFE:
				Select case pcdUsuario				
					case DIRECTOR_USER: 
					' DIRECTOR
						myWhere = myWhere & FIRMA_ROL_DIRECTOR					
				End Select
		end select		
		strSQL = " SELECT cdusuario FROM tblregistrofirmas WHERE " & myWhere
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		while(not rs.eof)
			clave = rs("CDUSUARIO")
			if (not vUser.Exists(clave)) then 	
				Call vUser.Add(clave,0)
				Call myLog.info("El usuario " & clave & " tiene firmas pendientes!")
			end if
			rs.MoveNext()
		wend			
	else
		' PROCESA LOS USUARIOS QUE YA VIENEN IDENTIFICADOS CON SU USERNAME
		if (not vUser.Exists(pcdUsuario)) then  			
			Call vUser.Add(pcdUsuario,0)
			Call myLog.info("El usuario " & pcdUsuario & " tiene firmas pendientes!")
		end if
	end if
End function
'------------------------------------------------------------------------------------------------
' Función:	enviarNotificacionesPendientes
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	--/--/----
' Objetivo:	
'			Enviar mail a los responsables de cada firma para notificarle los documentos pendientes a firmar,
'			obteniendolos del Dictionay
' Parametros:
'				-
' Devuelve:
'				-
' Modificacion: 
'			27/11/2012
'----------------------------------------------------------------------------------------------
Function enviarNotificacionesPendientes()	
	Dim mensaje,asunto,origen,destino, theKey
	mensaje = " Usted tiene disponible documentos pendientes para firmar "
	asunto  = GF_TRADUCIR("AUTORIZACIONES PENDIENTES - ALERTA FIRMA PENDIENTE ")			
	for each theKey in vUser.Keys		
		if not enviarMail(asunto, getUserMail(theKey), mensaje) then
			Call myLog.info("ERROR - No se puedo enviar el mail al usuario " & theKey)
		end if
	next
End Function
'----------------------------------------------------------------------------------------------
' Función:	  prepararAlertaCTC
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  20/05/2013
' Objetivo:	  Prepara el texto del mail quen recibirá cada usuario
' Parametros:
'				-
' Devuelve:
'				-
Function prepararAlertaCTC(usuario, rs, ByRef dic)
    
    Dim mailMsj
    
    if (not dic.Exists(usuario)) then
        mailMsj = "Los siguientes contratos estan por vencer y requieren su atención:" & vbcrlf & vbcrlf
        Call dic.add(usuario, mailMsj)		
    end if
	mailMsj = "---------------------------------------------------------------------------------------" & vbcrlf		
    mailMsj = mailMsj & "Contrato: " & rs("CDCONTRATO")& " - " & rs("TITULO") & vbcrlf    
	mailMsj = mailMsj & "Vencimiento: " & GF_FN2DTE(rs("FECHAVTO")) & vbcrlf
    mailMsj = mailMsj & "Asignado a Pedido: " & rs("CDPEDIDO") & " - " & rs("TITULO") & vbcrlf
    mailMsj = mailMsj & "Responsable: " & rs("CDRESPONSABLE") & " - " & getUserDescription(rs("CDRESPONSABLE")) & vbcrlf
    mailMsj = mailMsj & "Proveedor Elegido: " & rs("IDPROVEEDOR") & " - " & getDescripcionProveedor(rs("IDPROVEEDOR")) & vbcrlf
    mailMsj = mailMsj & "Cuit Proveedor: " & GF_STR2CUIT(getCUITProveedor(rs("IDPROVEEDOR"))) & vbcrlf
    mailMsj = mailMsj & "---------------------------------------------------------------------------------------" & vbcrlf
    
    dic.Item(usuario) = dic.Item(usuario) & mailMsj
    
End Function
'----------------------------------------------------------------------------------------------
' Función:	  buscarAlertasCTC
' Autor: 	  CNA - Ajaya Nahuel
' Autor Modificación: 	  CNA - Ajaya Nahuel
' Fecha: 	  01/03/2013
' Fecha Modificación: 	  04/11/2013
' Objetivo:	  Envia un mail de alerta cuando un CTC esta por vencer o ya vencio, solo busca Contrato de tipo General.
' Parametros:
'				-
' Devuelve:
'				-
'----------------------------------------------------------------------------------------------
Function buscarAlertasCTC()
	Dim rs, myDia, fechaHasta, diasDiff5, diasDiff2, mailTo, mailMsj
	Dim dicAlertas, key, mailCompras, myUsr
	
	Set dicAlertas = Server.CreateObject("Scripting.Dictionary")	
	myDia = Day(Now())
	Call myLog.info("##########################################")
	Call myLog.info("BUSCANDO ALERTAS CTC")
	Call myLog.info("##########################################")	
	'Primero verifico si hoy corresponde enviar alertas. (Responsables dia por medio y legales cada 5 dias)
	'Se da aviso cada 2 y 5 días, varia según a quien se le avise.
	diasDiff5 = myDia mod 5 
	diasDiff2 = myDia mod 2 
	if ((diasDiff2 = 0) or (diasDiff5 = 0)) then	
	    fechaHasta = GF_DTEADD(Left(session("MmtoSistema"),8),2,"m")
	    Set rs = buscarCTCporVencer(fechaHasta, "")
	    if (not rs.EoF)then
	        Call myLog.info("Hay Contratos/servicios con vencimientos para notificar!!")	
	        mailCompras = MAILTO_COMPRAS
		    while(not rs.EoF)
		        if (rs("CDCONTRATO") <> CONTRATO_TIPO_SERVICIO) then
			        if (diasDiff2 = 0) then
			            Call myLog.info("Notifico a Responsable de Compras y Responsable del Contrato " & rs("CDCONTRATO"))
			            Call prepararAlertaCTC(mailCompras, rs, dicAlertas)	
						'--> JAS: Habilitar luego de tener el nuevo active Directory en ADM
			            'myUsr = rs("CDRESPONSABLE")
		                'Call prepararAlertaCTC(myUsr, rs, dicAlertas)		        
						'<--
		            end if
		            if (diasDiff5 = 0) and (rs("TIPO") <> CTC_TIPO_OBRA)then
		                Call myLog.info("Notifico a Legales. Contrato " & rs("CDCONTRATO"))		        
		                Call prepararAlertaCTC(SENDER_LEGALES, rs, dicAlertas)
		            end if
                else
                    Call myLog.info("Es un servicio, no se notifica. IDCONTRATO=" & rs("IDCONTRATO"))	
                end if		            
		        rs.MoveNext()
		    wend
		    'Se envian los mails.		
		    for each key in dicAlertas.Keys
		        mailMsj = dicAlertas.item(key)		    
				'--> JAS: Habilitar luego de tener el nuevo active Directory en ADM
		        'if ((key <> mailCompras) and (key <> SENDER_LEGALES)) then					
		            'mailTo = getUserMail(key)		    
		        'else
		            mailTo = key
		        'end if
		        Call myLog.info("Se envia mail a: " & mailTo)				    
		        if not enviarMail("Sistema de Compras Web - Alerta vencimientos de Contratos ", SENDER_LEGALES, mailTo, mailMsj) then
		            Call myLog.info("ERROR - No se puedo enviar la alerta de mail de CTC a las direcciones:  " & mailTo)
	            end if
		    next				
        else		    
            Call myLog.info("No hay contrato a vencer para el periodo analizado.")
	    end if	
    else
        Call myLog.info("Hoy no es un día en que deban enviarse estas alertas.")
    end if	    
	Call myLog.info("------------------FIN DE LA BUSQUEDA----------------")	
End Function
'----------------------------------------------------------------------------------------------
' Función:	  
'				actualizarEstadosCTC
' Autor: 	
'				CNA - Ajaya Nahuel
' Fecha: 		
'				05/03/2013
' Objetivo:		
'				Actualizar el estado cuando un Contrato este mas de un mes vencido, de manera de dar tiempo a completar los pagos.
' Parametros:	-
' Devuelve:		RecordSet
'----------------------------------------------------------------------------------------------
Function actualizarEstadosCTC()	
	Dim fecha, rs
	
	'fecha = GF_DTEADD(Left(session("MmtoSistema"),8),-10,"D")	
	fecha = Left(session("MmtoSistema"),8)
	Call myLog.info("##########################################")
	Call myLog.info("ACTUALIZANDO ESTADOS CTC ")
	Call myLog.info("##########################################")		
	Set rs = buscarCTCporVencer(fecha,"")
	while(not rs.EoF)
		strSQL = "UPDATE TBLOBRACONTRATOS SET ESTADO = " & ESTADO_CTC_FINALIZADO & " WHERE IDCONTRATO = " & rs("IDCONTRATO")
		Call executeQueryDb(DBSITE_SQL_INTRA, rsX, "EXEC", strSQL)	
		Call myLog.info("Actualizando CTC: " & rs("CDCONTRATO") & " - " & rs("TITULO") & " a estado FINALIZADO ")
		rs.MoveNext
	wend 
	Call myLog.info("----------------FIN DE LA ACTUALIZACION-------------")
End Function
'----------------------------------------------------------------------------------------------
Function actualizarDolar()
    Dim tipoCambio, strSQL, rs
    
    Call myLog.info("##########################################")
	Call myLog.info("CAMBIO DE COTIZACION DEL DOLAR EN PUERTOS")
	Call myLog.info("##########################################")
	
    tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
    
    strSQL = "Update PARAMETROS Set VLPARAMETRO='" & tipoCambio & "' where CDPARAMETRO='COTIZACIONDOLAR'"    
    Call executeQueryDb(DBSITE_ARROYO	, rs, "EXEC", strSQL)
    Call executeQueryDb(DBSITE_TRANSITO	, rs, "EXEC", strSQL)
    Call executeQueryDb(DBSITE_BAHIA	, rs, "EXEC", strSQL)
    
    Call myLog.info("Nueva cotizacion: $" & tipoCambio)
    
End Function		
'**************************************************************
'********************	INICIO DE PAGINA   ********************
'**************************************************************
Dim cdUsuario,cant,vTipoDocumento,idAlmacen,tipoDoc,vUser,index,pto, myHoy
session("Usuario") = "JAS"

myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))

'if (isFormSubmit()) then	
    Set myLog = new classLog
	Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
	myLog.fileName = "AVISOS-AUTOMATICOS-" & myHoy
	'Call actualizarDolar()
	Call actualizarEstadoPedidos()	
	Call comprobarPedidosAVencer()
	Call actualizarEstadosCTC()
	Call buscarAlertasCTC()
	'Set vUser = Server.CreateObject("Scripting.Dictionary")
	'Call myLog.info("##########################################")
	'Call myLog.info("VERIFICANDO FIRMAS PENDIENTES")
	'Call myLog.info("##########################################")
	'cant = GF_PARAMETROS7("cantidad", 0, 6)
	'for i=0 to cant
	'	cdUsuario = GF_PARAMETROS7("cdUsuario_" & i , "", 6)
	'   idAlmacen = GF_PARAMETROS7("idAlmacen_" & i , "", 6)
	'	tipoDoc	  = GF_PARAMETROS7("idtipo_" & i , "", 6)
	'	pto		  = GF_PARAMETROS7("pto_" & i , "", 6)		
	'	Call controlarIdentificacionUsuario(cdUsuario,idAlmacen,tipoDoc,pto)
	'next	
	'Call enviarNotificacionesPendientes()	
'end if
%>
<html>
<head>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript">
/*
	var ch = new channel();				
	var arrUsuario = new Array();
	var arrAlmacen = new Array();
	var arrTipoDoc = new Array();	
	var arrPto	   = new Array();
	var vTipoDocumento = new Array();

	function bodyOnLoad(){
		<% if (not isFormSubmit()) then	 		
			 vTipoDocumento = getDocumentoFirmar() 
			 for i=0 to UBound(vTipoDocumento)-1 %>
				vTipoDocumento.push('<% =vTipoDocumento(i) %>');
			<% next %>			
			var tipoDoc = vTipoDocumento.pop();			
			ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" + tipoDoc + "&accion=<%=ACCION_PROCESAR%>&origen=1","CallBack_getAutorizaciones()");
			ch.send();
		<% end if %>
	}
*/		
	/* ESTA FUNCION ME PERMITE APLICAR EL PRIMER FILTRO A LOS DATOS A ENVIAR, DE ESTA MANERA SOLO PASAN AQUELLOS QUE NO SE REPITEN */	
/*	
	function comprobarDuplicados(pUsuario, pAlmacen, pTipo, pPto){
		var rtrn = false;
		if((pTipo == '<%=AUTH_TYPE_AJD%>')||(pTipo == '<%=AUTH_TYPE_AJC%>')||(pTipo == '<%=AUTH_TYPE_AJM%>')){
			for( var i in arrUsuario){
				if((arrUsuario[i] == pUsuario)&&(arrPto[i] == pPto)) rtrn = true;
			}	
		}
		else{		
			for( var i in arrUsuario){
				if((arrUsuario[i] == pUsuario)&&(arrAlmacen[i] == pAlmacen)) rtrn = true;
			}	
		}	
		return rtrn;
	}	
	
	function CallBack_getAutorizaciones(){
		var rtrn = ch.response();
		var arr = rtrn.split(";");
		if(rtrn.length > 0){
			for (i in arr) {
				var val = arr[i].split("|");
				if(val[0] != '<%= FIRMA_NO_USER %>'){
					if(!comprobarDuplicados(val[0],val[1],val[2],val[3])){
						arrUsuario.push(val[0]);
						arrAlmacen.push(val[1]);
						arrTipoDoc.push(val[2]);
						arrPto.push(val[3]);
					}
				}
			}
		}
		if (vTipoDocumento.length > 0) {
			var tipoDoc = vTipoDocumento.pop();
			ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" +  tipoDoc + "&accion=<%=ACCION_PROCESAR%>&origen=1","CallBack_getAutorizaciones()");
			ch.send();
		}
		else{			
			MostrarListaTipoDocumento();
		}	
	}
		
	function MostrarListaTipoDocumento(){
		var cont = 0;
		var myForm = document.getElementById('frmSel');			
		for(var z= 0;z<arrUsuario.length;z++){
			var cdUsuario = document.createElement("input");
			cdUsuario.id   = 'cdUsuario_' + z;
			cdUsuario.name = 'cdUsuario_' + z;
			cdUsuario.type = 'hidden';
			cdUsuario.value = arrUsuario[z];
			var idAlmacen = document.createElement("input");
			idAlmacen.id   = 'idAlmacen_' + z;
			idAlmacen.name = 'idAlmacen_' + z;
			idAlmacen.type = 'hidden';
			idAlmacen.value = arrAlmacen[z];						
			var idtipo = document.createElement("input");
			idtipo.id   = 'idtipo_' + z;
			idtipo.name = 'idtipo_' + z;
			idtipo.type = 'hidden';
			idtipo.value = arrTipoDoc[z];
			var pto = document.createElement("input");
			pto.id   = 'pto_' + z;
			pto.name = 'pto_' + z;
			pto.type = 'hidden';
			pto.value = arrPto[z];
			myForm.appendChild(cdUsuario);
			myForm.appendChild(idAlmacen);
			myForm.appendChild(idtipo);
			myForm.appendChild(pto);
		}
		var cantidad = document.createElement("input");
		cantidad.id   = 'cantidad';
		cantidad.name = 'cantidad';
		cantidad.type = 'hidden';
		cantidad.value = arrUsuario.length - 1;
		myForm.appendChild(cantidad);
		var accion = document.createElement("input");
		accion.id   = 'accion';
		accion.name = 'accion';
		accion.type = 'hidden';
		accion.value = '<%=ACCION_SUBMITIR%>';
		myForm.appendChild(accion);
		submitInfo();
	}
	
	function submitInfo() {
		document.getElementById("frmSel").submit();
	}
	
*/	
</script>
</head>
<body>
<form id="frmSel" name="frmSel">	
</form>
</body>
</html>
