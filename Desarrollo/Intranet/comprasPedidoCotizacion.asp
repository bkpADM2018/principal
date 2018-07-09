<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/MD5.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call comprasControlAccesoCM(RES_CC)

Const LICITACIONES_ARGENTINA = "Compras Argentina"

Dim idPedido, index, esModificable, controlOK, submitPage, accion, cambiaPlazo, esCancelable, controlObservaciones 
Dim rsObras, flagGuardar, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myPedidoComment
dim minCantPro, nrmName, obraFlag, rsDivision, especifTecnica, condParticulares, hayET, hayCP, cantprovExistentes
Dim fechaInicioOld,fechaCierreOld,cdSolicitanteOld,tituloPedidoOld,idObraOld,idDivisionOld,dsPedidoOld,observacionesOld,idAreaOld,idDetalleOld
Dim registroModificaciones()
redim registroModificaciones(0)
'------------------------------------------------------------------------------
'Funcion que permite determinar si el Pedido de Cotización puede modificarse
Function puedeModificar() 
	
	dim km, kc, ds	
	puedeModificar = false	
	if (pct_idPedido = 0) then	
		'Es un pedido nuevo, se pueden cargar datos.
		puedeModificar = true
	else
		'Solo lo puede modificar el usuario que lo cargo o bien el solicitante o bien un administrador	
		if (checkControlPCT()) then
			'Verifico si esta en un estado modificable.				
			if ((not isAuditor(pct_idDivision)) and ((pct_idEstado = ESTADO_PCT_PENDIENTE) or ((pct_idEstado < ESTADO_PCT_PUBLICADO) and (pct_cdSolicitante = session("Usuario"))))) then puedeModificar = true				
		end if		
	end if	
End Function
'------------------------------------------------------------------------------
Function puedeCancelar()	
	puedeCancelar = false
	'Un administrador puede cancelar en cualquier momento, un solicitante antes de que se publique el pedido.
	if (checkControlPCT() and (not isAuditor(pct_idDivision)) and (pct_idPedido > 0) and (pct_idEstado <= ESTADO_PCT_APROBADO)) then puedeCancelar = true
End Function
'------------------------------------------------------------------------------
'**************** 	MODIFICADO 	********************************
'****************  Fecha : 29/09/2015 - *******************************
'****************	Autor : Ajaya Nahuel - CNA 	**********************
Function puedeCambiarPlazo()	
	puedeCambiarPlazo = false
	'Primero verifico si el usuario tiene permiso para modificar el pedido (solicitante o usuario o administrador o Admin de la division)
    if ((pct_cdSolicitante = session("Usuario"))or(pct_usuario = session("Usuario"))or(pct_usuarioAdmin = session("Usuario"))or(isAdmin(pct_idDivision))) then
	    'Por ultimo verifico la marca de extension, si es 'S' si puede
        'if (pct_Extension = TIPO_AFIRMACION) then 
		puedeCambiarPlazo = true
	end if
End Function
'------------------------------------------------------------------------------
Function puedeConfirmar()
dim flag, rtrn
rtrn = false
if pct_idPedido > 0 then
	if ((isAdmin(pct_idDivision) or (pct_cdSolicitante = session("Usuario"))) and _
		(pct_idEstado = ESTADO_PCT_PENDIENTE)) then 
		rtrn = true	
	end if
end if	
puedeConfirmar = rtrn
end function
'------------------------------------------------------------------------------
Function puedeAgregarProveedor()
	Dim rol
	
	rol = CInt(getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)) 
	
	puedeAgregarProveedor=false	
	'if ((rol = FIRMA_ROL_GTE_COMPRAS) and (pct_idEstado < ESTADO_PCT_ABIERTO)) then
		puedeAgregarProveedor=true
	'end if		
End Function
'------------------------------------------------------------------------------
Function enviarMail(asunto, msg, email)
	Dim emailSender, obras, dsTipoCompra,i, usrAdmin, emailSolicitante
	enviarMail = false
	emailSender = getUserMail(session("Usuario"))	
	'email = "scalisij@toepfer.com;" & email
	'Response.Write "(" & email & ")<br>"
	if (email <> "") then
		emailSolicitante = getUserMail(pct_cdSolicitante)		
		msg = msg & vbCrLf & vbCrLf
		msg = msg & "Datos del Pedido" & vbCrLf
		msg = msg & String(100, "-") & vbCrLf
		msg = msg & "Codigo asignado.....: " & pct_cdPedido & vbCrLf		
		msg = msg & "Titulo..............: " & pct_tituloPedido & vbCrLf
		'Set obras = obtenerListaObras(pct_idObra, "", "","",OBRA_ACTIVA)		
		Set obras = obtenerDescripcionCompletaDetalle(pct_idObra, pct_idArea, pct_idDetalle)
		if ((not obras.eof) and (CLng(pct_idObra) <> 0)) then
			msg = msg & "Ptda. Presupuestaria: " & obras("CDOBRA") & " - " & obras("DSOBRA") & vbCrLf
			if (pct_idArea <> 0) then			
				msg = msg & "                      ----> " & obras("IDAREA") & " - " & obras("DSAREA") & vbCrLf
				if (pct_idDetalle <> 0) then
					msg = msg & "                            ----> " & obras("IDDETALLE") & " - " & obras("DSDETALLE") & vbCrLf
				end if
			end if			
		end if
		msg = msg & "Solicitante.........: " & pct_cdSolicitante & "-" & pct_dsSolicitante & vbCrLf					
		msg = msg & "Tipo de Pedido......: Pedido de Precios" & vbCrLf		
		msg = msg & "Fecha de Emisión....: " & pct_FechaInicio & vbCrLf		
		msg = msg & "Fecha de Limite.....: " & pct_FechaCierre & vbCrLf
		msg = msg & "Administra..........: " & LICITACIONES_ARGENTINA & vbCrLf
		msg = msg & "Descripcion: " & vbCrLf & pct_dsPedido & vbCrLf
		
		if (UBound(registroModificaciones)) then
			msg = msg & vbCrLf & "Se modificaron los siguientes puntos:" & vbCrLf 
			msg = msg & String(100, "-")  & vbCrLf 
			for each valor in registroModificaciones
				msg = msg & Trim(valor) & vbCrLf & vbCrLf
			next
		end if
		'response.write "<pre>" & msg  & "</pre>"
		'response.end
		Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web -" & asunto) & ": " & pct_cdPedido, msg, emailSender, email)
		enviarMail = true
		'Response.Write "Mando!<br>"
	end if	
	
End function
'------------------------------------------------------------------------------
Function enviarMailGrupoNuevo()
	Dim strMsg
	
	strMsg = "Se ha cargado un nuevo pedido de cotización." & vbCrLf & vbCrLf
	enviarMailGrupoNuevo = enviarMail("Alerta Nuevo Pedido", strMsg, obtenerMail(CDTOEPFER))		
	
End Function
'------------------------------------------------------------------------------
Function actualizarEstadoAutorizadoPCT(pIdPedido)
	dim  strSQL, rs, conn
	strSQL="Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_AUTORIZADO & " where IDPEDIDO=" & pIdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'------------------------------------------------------------------------------
Function enviarMailSolicitanteNuevo()
	Dim strMsg, emailDestino, strSQL
	if (session("Usuario") <> pct_cdSolicitante) then   
		strMsg = "Se ha cargado un nuevo pedido de cotización." & vbCrLf & vbCrLf	
		strMsg = "Se requiere que ingrese al sistema y confirme la información para proceder a presentar la solicitud de precios a los proveedores." & vbCrLf & vbCrLf
		enviarMailSolicitanteNuevo = enviarMail("Alerta Nuevo Pedido", strMsg, getUserMail(pct_cdSolicitante))						
	end if
End Function
'------------------------------------------------------------------------------
Function enviarMailGrupoUpdate()
	Dim strMsg
	
	strMsg = "Se ha modificado el pedido de cotización: " & pct_cdPedido & vbCrLf & vbCrLf
	enviarMailGrupoUpdate = enviarMail("Modificación de Pedido", strMsg, obtenerMail(CD_TOEPFER))		
	
End Function
'------------------------------------------------------------------------------
Function enviarMailSolicitanteUpdate()
	Dim strMsg, asunto
	
	strMsg = "Se ha modificado el pedido de cotización: " & pct_cdPedido & vbCrLf & vbCrLf	
	asunto = "Modificación de Pedido"
	if (pct_idEstado = ESTADO_PCT_PENDIENTE) then 
		strMsg = "Se requiere que ingrese al sistema y confirme la información para proceder a presentar la solicitud de precios a los proveedores." & vbCrLf & vbCrLf
		asunto = "Solicitud de Confirmación"
	end if
	enviarMailSolicitanteUpdate = enviarMail(asunto, strMsg, getUserMail(pct_cdSolicitante))
	
End Function

'------------------------------------------------------------------------------
Function enviarMailNuevo()
	Call enviarMailGrupoNuevo()
	Call enviarMailSolicitanteNuevo()
End Function
'------------------------------------------------------------------------------
Function enviarMailUpdate()
	Dim myConfirm
	
	if (session("Usuario") <> pct_cdSolicitante) then Call enviarMailSolicitanteUpdate()
	enviarMailGrupoUpdate()
	
	if (pct_idEstado >= ESTADO_PCT_PUBLICADO) then
		myConfirm = "<script>var resp = confirm('Se realizaron modificaciones. Desea anviar un Email a los proveedores con los cambios?');if(resp){"
		myConfirm = myConfirm & "window.open('comprasEnvioPCTMail.asp?idPedido="&idPedido&"', '_blank', 'location=no,menubar=no,statusbar=no,height=240,width=500','false');"
		myConfirm = myConfirm & "}</script>"
		response.write myConfirm
	end if
End Function
'------------------------------------------------------------------------------
Function registrarCambio(pMensaje)	
	Dim pos
		
	pos = UBound(registroModificaciones)+1
	redim preserve registroModificaciones(pos)
	registroModificaciones(pos) = pMensaje
	registrarCambio = true
	
End Function
'------------------------------------------------------------------------------
Function controlarModificacion()
	Dim rtrn, provCambio, path, fileEspecifTecnica, fileReglamento
	rtrn = false
    
	'if (pct_FechaInicio	<> fechaInicioOld) then rtrn = registrarCambio("La Fecha de Emisión fue modificada, su valor era........: " & fechaInicioOld)
    
	if (pct_FechaCierre	<> fechaCierreOld) then	
		rtrn = registrarCambio("La Fecha de Cierre fue modificada, su valor era.........: " & fechaCierreOld)
		'Se le aplica la marca de Extendido para que no lo vuelva hacer
		pct_Extension = TIPO_NEGACION		
        if (GF_DTEDIFF(GF_DTE2FN(fechaCierreOld), Left(session("MmtoSistema"),8),"D") > 0) then pct_idEstado = ESTADO_PCT_PUBLICADO
	end if	

    if (pct_cdSolicitante <> cdSolicitanteOld) then rtrn = registrarCambio("El Solicitante fue modificado, anteriormente era........: " & cdSolicitanteOld & " - " & getUserDescription(cdSolicitanteOld))
	
	if (pct_tituloPedido <> tituloPedidoOld) then rtrn = registrarCambio("El Titulo fue modificado, su valor era..................: " & tituloPedidoOld)
	
	path=""
	fileEspecifTecnica = GF_PARAMETROS7("etFile","",6)		
	if (fileEspecifTecnica <> "") then path = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & fileEspecifTecnica
	if (isFileModified(path, idPedido, PCT_BINARY_SPECIFICATION)) then rtrn = registrarCambio("Se ha cambiado la especificación técnica del pedido.")
	
	path=""
	fileReglamento = GF_PARAMETROS7("rgFile","",6)		
	if (fileReglamento <> "") then path = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & fileReglamento
	if (isFileModified(path, idPedido, PCT_BINARY_CONDITIONS)) then rtrn = registrarCambio("Se han cambiado las condiciones particulares del pedido.")		
	
	if (pct_idObra <> Cdbl(idObraOld)) then rtrn = registrarCambio("La Ptda. Presupuestaria fue modificada, su valor era....: " & idObraOld & " - " & getDescripcionObra(idObraOld))

    if (GF_PARAMETROS7("idDivision",0,6) <> CLng(idDivisionOld)) then rtrn = registrarCambio("La Division fue modificada, su valor era................: " & getDivisionDS(idDivisionOld))

	if (pct_dsPedido <> dsPedidoOld) then rtrn = registrarCambio("La Descripcion ha sido modificada.")
	
	if (pct_observaciones <> observacionesOld) then rtrn = registrarCambio("Las Observaciones fueron modificadas.")		
	
	if (pct_idArea <> CLng(idAreaOld)) then rtrn = registrarCambio("El area de la Ptda. Presupuestaria fue modificada, su valor era...:" & idAreaOld)
	
	if (pct_idDetalle <> Clng(idDetalleOld)) then rtrn = registrarCambio("El detalle de la Ptda. Presupuestaria fue modificada, su valor era:"& idDetalleOld)
	
	auxIndex = 0
    provCambio = false
    if (initProveedoresDB()) then
        while ((readNextProveedorDB())and(not provCambio))
            auxProveedorOld = GF_PARAMETROS7("supplier" & auxIndex, 0, 6)
            if (CLng(pct_idProveedor) <> CLng(auxProveedorOld)) then provCambio=true
			auxIndex = auxIndex + 1
		wend
	end if
	if (GF_PARAMETROS7("supplier" & auxIndex, 0, 6) <> 0) then provCambio = true
	if (provCambio) then rtrn = registrarCambio("Los Proveedores participantes han sido modificados.")
	controlarModificacion = rtrn
end Function
'------------------------------------------------------------------------------
Function fecthDataPCTOriginal()
    fechaInicioOld = pct_FechaInicio
    fechaCierreOld = pct_FechaCierre
    cdSolicitanteOld = pct_cdSolicitante
    tituloPedidoOld = pct_tituloPedido
    idObraOld = pct_idObra
    idDivisionOld = pct_idDivision
    dsPedidoOld =  pct_dsPedido
    observacionesOld =  pct_observaciones
    idAreaOld = pct_idArea
    idDetalleOld = pct_idDetalle
End function 
'------------------------------------------------------------------------------
Function  enviarMailConfirmacion()
	Dim ret
	
	strMsg = "Se ha confirmado el pedido " & pct_cdPedido & ", el mismo se encuentra listo para ser enviado a los proveedores." & vbCrLf	
	ret = enviarMail("Pedido Confirmado", strMsg, obtenerMail(CD_TOEPFER))
	enviarMailConfirmacion = ret
	
End Function
'--------------------------------------------------------------------------------------------------
Function showUploadFiles(pHayFile,pModificable,pIdDiv,pUpload)

'============================================================================
' 	F = Archivos | A = Abierto (Modificable) | C = Cerrado (No modificable) | L = Admin de licitaciones
' 	0 = Sin | 1 = Con
'
' Caso|	F	A	C	L	 
' ----|----------------
'	1 |	0	0	1	0 	=>	Sin Archivos
'	2 |	0	0	1	1 	=>	Uploads para Admin de Lic.
'	3 |	0	1	0	0 	=>	Uploads
'	4 |	0	1	0	1 	=>	Uploads
'   5 |	1	0	1	0 	=>	Downloads
'	6 |	1	0	1	1 	=>	Uploads para Admin de Lic.
'	7 |	1	1	0	0 	=>	Uploads
'	8 |	1	1	0	1 	=>	Uploads
'============================================================================

Dim isAdmLic,rtrn
	
	rtrn = ""
	
	isAdmLic = isUserInGroup(session("Usuario"),"licitaciones.ar")
		
	'if (pct_idEstado > ESTADO_PCT_PUBLICADO) then
	'	if ( (not pHayFile) and (not pModificable) and (not isAdmLic) ) then rtrn = "No hay archivos asociados."
'
'		if ( (pHayFile) and (not pModificable) and (not isAdmLic) ) then 
'			if (pUpload = "ET") then rtrn = "<img align='absMiddle' src='images/doc.gif'>&nbsp;<a target='_blank' href='comprasOpenArchivo.asp?idPedido="&pct_idPedido&"&fileno="&PCT_BINARY_SPECIFICATION&"'>"&especifTecnica&"</a>"
'			if (pUpload = "CP") then rtrn = "<img align='absMiddle' src='images/doc.gif'>&nbsp;<a target='_blank' href='comprasOpenArchivo.asp?idPedido="&pct_idPedido&"&fileno="&PCT_BINARY_CONDITIONS&"'>"&condParticulares&"</a>"
'		end if
'	else		

		'Caso 1
		if ( (not pHayFile) and (not pModificable) and (not isAdmLic) ) then rtrn = "No hay archivos asociados."
		
		'Caso 2
		if ( (not pHayFile) and (not pModificable) and (isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
		
		'Caso 3
		if ( (not pHayFile) and (pModificable) and (not isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
		
		'Caso 4
		if ( (not pHayFile) and (pModificable) and (isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
			
		'Caso 5
		if ( (pHayFile) and (not pModificable) and (not isAdmLic) ) then 
			if (pUpload = "ET") then rtrn = "<img align='absMiddle' src='images/doc.gif'>&nbsp;<a target='_blank' href='comprasOpenArchivo.asp?idPedido="&pct_idPedido&"&fileno="&PCT_BINARY_SPECIFICATION&"'>"&especifTecnica&"</a>"
			if (pUpload = "CP") then rtrn = "<img align='absMiddle' src='images/doc.gif'>&nbsp;<a target='_blank' href='comprasOpenArchivo.asp?idPedido="&pct_idPedido&"&fileno="&PCT_BINARY_CONDITIONS&"'>"&condParticulares&"</a>"
		end if
		
		'Caso 6
		if ( (pHayFile) and (not pModificable) and (isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
		
		'Caso 7
		if ( (pHayFile) and (pModificable) and (not isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
		
		'Caso 8
		if ( (pHayFile) and (pModificable) and (isAdmLic) ) then rtrn = "<div id='"&pIdDiv&"'></div>"
'	end if
	
	showUploadFiles = rtrn	
	
End Function
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************

call GP_ConfigurarMomentos()

idPedido = GF_PARAMETROS7("idPedido",0,6)
accion = GF_PARAMETROS7("accion","",6)
myPedidoComment = GF_PARAMETROS7("pedidoComment","",6)
especifTecnica = GF_PARAMETROS7("etFile","",6)
condParticulares = GF_PARAMETROS7("rgFile","",6)
cantprovExistentes = GF_PARAMETROS7("provExistentes",0,6)

flagGuardar = false
controlOK = false

'Se cargan los datos del Pedido para mostrar en pantalla
Call initHeader(idPedido)

'Se controla si tiene acceso a la información
if (idPedido <> 0) then
	if (not checkPointAcceso(pct_idDivision)) then response.redirect "comprasAccesoDenegado.asp"	
	'Se controla si tiene acceso a la información
	if (not checkControlPCT()) then	response.redirect "comprasAccesoDenegado.asp"
end if

flagDebeConfirmar = puedeConfirmar()

if (isFormSubmit()) then
    fechaInicioOld = GF_PARAMETROS7("fechaInicioOld","",6)
    fechaCierreOld = GF_PARAMETROS7("fechaCierreOld","",6)
    cdSolicitanteOld = GF_PARAMETROS7("cdSolicitanteOld","",6)
    tituloPedidoOld = GF_PARAMETROS7("tituloPedidoOld","",6)
    idObraOld = GF_PARAMETROS7("idObraOld",0,6)
    idDivisionOld = GF_PARAMETROS7("idDivisionOld",0,6)
    dsPedidoOld = GF_PARAMETROS7("dsPedidoOld","",6)
    observacionesOld = GF_PARAMETROS7("observacionesOld","",6)
    idAreaOld = GF_PARAMETROS7("idAreaOld",0,6)
    idDetalleOld = GF_PARAMETROS7("idDetalleOld",0,6)
	if (accion = ACCION_CONFIRMAR) then						
		call actualizarEstadoAutorizadoPCT(idPedido)
		pct_idEstado = ESTADO_PCT_AUTORIZADO
		Call enviarMailConfirmacion()
		flagDebeConfirmar = puedeConfirmar()
	else
		'Se controlan los datos.				
		controlOK = controlarPedidoCotizacion(idPedido)
		if ((accion = ACCION_GRABAR) and (controlOK)) then				
			if (idPedido = 0) then 
				flagGuardar = true						
				idPedido = grabarFormulario()
				'Call enviarMailNuevo()
			else
				if (controlarModificacion()) then					
					idPedido = grabarFormulario()
					Call enviarMailUpdate()
				end if
			end if
		end if	
	end if
else
    fecthDataPCTOriginal()
end if

esModificable   = puedeModificar()
aceptaProveedor = puedeAgregarProveedor()
cambiaPlazo     = puedeCambiarPlazo()
esCancelable    = puedeCancelar()

especifTecnica = GF_PARAMETROS7("etFile","",6)
if (especifTecnica <> "") then 
	hayET = true
else
	hayET = hayEspecifTecnica(pct_idPedido)
	if (hayET) then  especifTecnica = buildFileName(pct_idPedido, PCT_BINARY_SPECIFICATION, "")
end if

condParticulares = GF_PARAMETROS7("rgFile","",6)
if (condParticulares <> "") then 
	hayCP = true
else
	hayCP = hayCondParticulares(pct_idPedido)
	if (hayCP) then  condParticulares = buildFileName(pct_idPedido, PCT_BINARY_CONDITIONS, "")
end if
%>
<html>
<head>
<title>Pedido de Cotizacion</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/uploadManager.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var ch = new channel();
	//Constantes - Nombre de Campo	
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var provExistentes = 0;
	var idBtnGuardar = 0;
	var idBtnControl = 0;
	var up1 = new UploadHandler("fileEspec","<% =PATH_COMPRAS_TEMP %>");		
	<% if (hayET) then %>
		up1.setFile('<% =especifTecnica %>');
	<%	end if	%>
	var up2 = new UploadHandler("fileReglamento","<% =PATH_COMPRAS_TEMP %>");		
	<% if (hayCP) then %>
		up2.setFile('<% =condParticulares %>');
	<% end if	%>
	function ReloadParent(){
		//document.parentWindow.location.reload();
	}
	function abrirCuadroCacenlar(idPedido){
	    pp = new winPopUp('pp', "comprasPCTCancelacionPopUp.asp?idPedido=" + idPedido, '500', '200', 'Cancelar Pedido', '');
	}
	function agregarLineaProveedor(force) {
		var tblProveedores = document.getElementById("tblProveedores");
		var rProveedor = tblProveedores.insertRow(lastProveedores+1);
		var cCodigo = rProveedor.insertCell(0);
		var cDescripcion = rProveedor.insertCell(1);		
		var cMail = rProveedor.insertCell(2);
		cMail.align = "center";
		cMail.id = SUPPLIER_MAIL + lastProveedores;
		var cCotizacion = rProveedor.insertCell(3);
		cCotizacion.align = "center";
		var dCodigo = document.createElement('div');
		dCodigo.className = "labelStyle";
		dCodigo.id = SUPPLIER_DIV + lastProveedores;		
		cCodigo.appendChild(dCodigo);
		var iDescripcion = document.createElement('div');		
		iDescripcion.id = SUPPLIER_DESC + lastProveedores;				
		cDescripcion.appendChild(iDescripcion);
		var ms;

<%		if ((esModificable) or (aceptaProveedor))then 				%>
		force=true
<%		end if								%>

		if (force) {
			ms = new MagicSearch("", SUPPLIER_DESC + lastProveedores, 60, 2, "comprasStreamElementos.asp?tipo=empresas&linea=" + lastProveedores);
			ms.setToken(";");
			ms.minChar = 3;
			ms.onBlur = "seleccionarProveedor(" + lastProveedores + ")";
		}
		var iCodigo = document.createElement('input');
		iCodigo.type = "hidden";
		iCodigo.id = SUPPLIER_ID + lastProveedores;
		iCodigo.name = SUPPLIER_ID + lastProveedores;
		cCodigo.appendChild(iCodigo);		
		var dCotizacion = document.createElement('div');
		dCotizacion.id = SUPPLIER_CT + lastProveedores;
		cCotizacion.appendChild(dCotizacion);
		lastProveedores++;
		document.getElementById("cantProveedores").value = lastProveedores;
		return ms;
	}	
	
	function aceptarPedido(id) {
		document.getElementById("accion").value = "<% =ACCION_CONFIRMAR %>";
		document.getElementById("frmSel").submit();
	}
	function callback_Cancelar(){
	    document.getElementById("frmSel").submit();
	}
	
	function seleccionarProveedor(pLinea, pMs) {		
		var desc = "";
		if (pMs) desc = pMs.getSelectedItem();		
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById(SUPPLIER_ID + pLinea).value = arr[0];
			document.getElementById(SUPPLIER_DIV + pLinea).innerHTML = arr[0];
			pMs.setValue(arr[1]);
		} else {
			if (desc == "") {
				document.getElementById(SUPPLIER_ID + pLinea).value = "";
				document.getElementById(SUPPLIER_DIV + pLinea).innerHTML = "";
			}
		}		
	}
	
	function seleccionarSolicitante(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdSolicitante").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdSolicitante").value = "";
		}
	}
	
	function seleccionarAdministrador(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdUsuarioAdmin").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdUsuarioAdmin").value = "";
		}
	}
	
	function SeleccionarCalEmision(cal, date) {
		var str= new String(date);		
		document.getElementById("issuedateDiv").innerHTML = str;
	    document.getElementById("issuedate").value = str;
		if (cal) cal.hide();	
	}
	
	function SeleccionarCalLimite(cal, date) {
		var str= new String(date);		
		document.getElementById("closingdateDiv").innerHTML = str;
	    document.getElementById("closingdate").value = str;
		if (cal) cal.hide();	
	}
	
	function CerrarCal(cal) {
		cal.hide();
	}
	
	function MostrarCalendario(p_objID, funcSel) {
		var dte= new Date();		    	    
		var elem= document.getElementById(p_objID);
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	
	function enviarMail(idPedido, idProveedor) {		
		window.open("comprasEnvioPCTMail.asp?idPedido=" + idPedido + "&idProveedor=" + idProveedor, "_blank", "location=no,menubar=no,statusbar=no,height=240,width=500",false);		
	}
	
	function fillProveedor(pMS, linea, id, desc, hayCT, mail, crc, crcEncrypt) {
		var puedeModificarProv = false;
<%		if (esModificable) then	%>
			puedeModificarProv = true;
<%		else	%>
<%			if ((isFormSubmit()) and (accion <> ACCION_GRABAR)) then %>
				provExistentes = document.getElementById("provExistentes").value;
				if (linea >= provExistentes) puedeModificarProv = true;
<%			end if%>
<%		end if%>
		if (puedeModificarProv) {
			pMS.setValue(id + "-" + desc)
			seleccionarProveedor(linea, pMS);
		} else {
			document.getElementById(SUPPLIER_DIV + linea).innerHTML = id;
			document.getElementById(SUPPLIER_ID + linea).value = id;
			document.getElementById(SUPPLIER_DESC + linea).innerHTML = desc;			
<%			if ((pct_idEstado <= ESTADO_PCT_COTIZADO) and (pct_idEstado => ESTADO_PCT_AUTORIZADO)) then %>
				var cMail = document.getElementById(SUPPLIER_MAIL + linea);
				var aMail = document.createElement('a');
				aMail.href = "javascript:enviarMail(<% =idPedido %>, " + id + ")";
				cMail.appendChild(aMail);
				var imgMail = document.createElement('img');
				imgMail.src = "images/compras/PCT_publish-16x16.png";
				aMail.appendChild(imgMail);
<%			end if		%>
			var a = document.createElement("a");
			var img = document.createElement("img");
            if (hayCT.toLowerCase() == 'true') {			
				img.src = "images/compras/CTZ-16x16.png";			
				a.href = "javascript:verCotizaciones(<% =idPedido %>)";			
			} else {
				img.src = "images/compras/supplier_key-16x16.png";					
				a.id = "emailLink_"+id;
                a.name = "emailLink_"+id;
                a.href = "javascript:generarMailPCTInterno("+id+",'"+desc+"','"+mail+"','"+crcEncrypt+"','"+ crc +"')";
			}
			a.appendChild(img);
			document.getElementById(SUPPLIER_CT + linea).appendChild(a);
		}
	}
	function generarMailPCTInterno(pIdProveedor,pDsProveedor,pMailProveedor,pCRCEncrypt,pCRC){
         $('[name=emailLink_'+pIdProveedor+']').each(function() {
            var email = '<%= SENDER_COTIZACIONES %>';
            var subject = 'Solicitud de Cotizacion - REF: <%=pct_cdPedido %>';
            var emailBody = "---------------NO MODIFIQUE POR DEBAJO DE ESTA LINEA---------------%0D%0A%0D%0A";
                emailBody = emailBody + pCRC + "%0D%0A";                
                emailBody = emailBody + pCRCEncrypt + "%0D%0A%0D%0A";
                emailBody = emailBody + "------------------------------------------------------------------------------------%0D%0A%0D%0A%0D%0A%0D%0A";
            window.location='mailto:' + email + '?subject=' + subject + '&body=' +   emailBody;
          });
    }		
	function verCotizaciones(idPedido) {		
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + idPedido + "&tab=2", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	
	function submitInfo(acc) {		
		document.getElementById("etFile").value = up1.getFileName();
		document.getElementById("rgFile").value = up2.getFileName();
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function canSubmit(acc, btn) {		
		if ((up1.isReady) && (up2.isReady)) {
			submitInfo(acc);
		}
		else {						
			var b = idBtnGuardar;
			if (btn == 1) b = idBtnControl;
			tb.changeLook(b,"loading_small_green.gif");	
			setTimeout("canSubmit('" + acc + "', " + btn + ")", 1000);
		}
	}
	
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}
	
	function volver() {	
		<% if (pct_idPedido = 0) then %>
			location.href = "comprasAdministrarPedidos.asp";
		<% else %>
			window.close();
		<% end if %>
	}
	function cargarAreaDetalle(idObra,idArea,idDetalle,modificable)
	{
		if (idObra != 0){
			//cargo el combo del area y detalle
			ch.bind("comprasObtenerBudget.asp?idobra="+idObra+"&idarea="+idArea+"&idDeta="+idDetalle+"&modificable="+modificable,"callBackBudget()");
			ch.send();
			$("#areadetalle").show();
		}else{
			//borro el combo del area y detalle
			//deshabilito el combo para que no se envie en el submit
			$("#cmbAreaDetalle").attr('disabled', 'disabled'); 
			$("#areadetalle").hide();
		}
	}
	function callBackBudget()
	{
		//$("#cmbAreaDetalle").HTML(ch.response());
		document.getElementById("areadetalle").innerHTML = ch.response();
		$("#cmbAreaDetalle").attr('disabled', ''); 
	}
	function getAreaDetalle(me)
	{
		aux = $(me).val().split("-")
		$("#cmbIdArea").val(aux[0]);
		$("#cmbIdDeta").val(aux[1]);
	}
	
	function bodyOnLoad() {
		var myMS;
		
		cargarAreaDetalle('<%=pct_idObra%>','<%=pct_idArea%>','<%=pct_idDetalle%>','<%=esModificable%>');
		
		tb = new Toolbar('toolbar', 6, 'images/compras/');		
<%	if (esModificable or aceptaProveedor or cambiaPlazo) then %>
		idBtnGuardar = tb.addButtonSAVE("Guardar", "canSubmit('<% =ACCION_GRABAR %>',0)");
		idBtnControl = tb.addButton("control-16x16.png","Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");								
<%	end if
	if(flagDebeConfirmar) then %>
		tb.addButton("PCT_confirm-16x16.png", "Confirmar", "aceptarPedido('<% = pct_idPedido %>')");		
<%	end if
	if (esCancelable) then %>		
		tb.addButton("cancel-16x16.png", "Cancelar", "abrirCuadroCacenlar('<% = pct_idPedido %>')");
<%	end if	%>
		tb.addButton("previous-16x16.png", "Volver", "volver()");
		tb.draw();
<%	if (esModificable) then %>		
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;
		msSolicitante.setValue('<% =pct_dsSolicitante %>');		
<%	end if %>
	<%	if (esModificable or cambiaPlazo) then %>
		up1.draw();
		up2.draw();		
	<% end if %>
	SeleccionarCalEmision(undefined, '<% = pct_FechaInicio %>');
	SeleccionarCalLimite(undefined, '<% = pct_FechaCierre %>');
<%
	index = 0
	if (initProveedores()) then
		while (readNextProveedor())
%>
			myMS = agregarLineaProveedor(false);
            var auxCrc = "";
            var auxCrcEncrypt = "";
        <%  if (not pct_hayCotizacion) then %>
                auxCrc = 'SV:<%=SISTEMA_COMPRAS %>|PR:<%=pct_idProveedor%>|P:<%=pct_idPedido%>|F:V';
                auxCrcEncrypt = '<%= MD5(generarCRCByPCT(pct_idProveedor,pct_idPedido,"V"))%>';
        <%  end if %>
            fillProveedor(myMS, <% =index %>, <% =pct_idProveedor %>, '<% =pct_dsProveedor %>', '<% =pct_hayCotizacion %>','<% =pct_emailProveedor %>',auxCrc, auxCrcEncrypt );
<%
			index=index+1
		wend
	end if
	'Se muestran las lineas requeridas por auditoria
	minCantPro = getValorNorma("MINPRCP")
	while ((index < cint(minCantPro)) and esModificable)
%>
		agregarLineaProveedor(false);
<%		
		index=index+1
	wend	
	'Si se grabo, se muestra el numero asignado
	if (accion = ACCION_GRABAR) then
		if (flagGuardar) then
		%>	
			pp = new winPopUp('pp', 'comprasConfirmacionNumero.asp?cdPedido=<% =pct_cdPedido %>', '440', '100', 'Pedido Guardado', 'irPedidos()');
<%		end if	%>
<%	end if %>
<%  if (accion = ACCION_CONFIRMAR) then %>
        irPedidos();
<%  end if %>
	<%if (not isFormSubmit()) then %>
		provExistentes = lastProveedores;
		document.getElementById("provExistentes").value = provExistentes;
	<%end if%>
	pngfix();
	}
	window.onload = bodyOnLoad;
</script>
</head>
<body>
	<% Call GF_TITULO2("kogge64.gif","Pedido de Cotización") %>
	<div id="toolbar"></div><br>
	<form id="frmSel" name="frmSel" action="comprasPedidoCotizacion.asp" method="post">
	<%	if (pct_idEstado = ESTADO_PCT_CANCELADO) then	%>
		<table class="reg_header" align="center" width="80%" border="0" >
			<tr class="TDERROR">
				<td><% =GF_TRADUCIR("ESTE PEDIDO HA SIDO CANCELADO") %></td>
			</tr>
		</table>
	<%	end if	%>
	<table class="reg_header" align="center" width="80%" border="0">
		<tr><td colspan="6"><% call showErrors() %></td></tr>
	<%	if flagDebeConfirmar then %>
		<tr>
			<td colspan="6" style="background-color:#ffff99">
				<img src="images/compras/warning-16x16.png" align="absMiddle">&nbsp;<b><% =GF_TRADUCIR("Este pedido de cotizacion necesita ser confirmado por el solicitante.") %></b>
			</td>
		</tr>
		<%
		end if	
				
		if (idPedido > 0) then %>
			<tr>								
				<td align="right" class="numberStyle" colspan="6"><% =GF_TRADUCIR("Nº Pedido") %>&nbsp;<% =pct_cdPedido %></td>				
			</tr>
	<% 	end if %>
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Datos del Pedido") %></td>				
			</tr>
			<%
				
			%>
			<tr>
				<td class="reg_header_navdos"><% =GF_TRADUCIR("Ptda. Presup.") %></td>
				<td colspan="2">
					<% if (esModificable) then 
						Set rsObras = obtenerListaObras("", "", "","", OBRA_ACTIVA)
					%>						
						<select id="idObra" name="idObra" onchange="cargarAreaDetalle($(this).val(),0,0,'<%=esModificable%>')">
							<option value="0">- <% =GF_TRADUCIR("Sin Ptda. Presupuestaria") %> -
					<%	while (not rsObras.eof)	%>							
							<option value="<% =rsObras("IDOBRA") %>" <% if (rsObras("IDOBRA") = pct_idObra) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %>
					<%		rsObras.MoveNext()
						wend 	%>		
						</select>						
					<% else						
						if (pct_idObra <> 0) then
							Set rsObra = obtenerListaObras(pct_idObra, "", "","", "") 
							if (not rsObra.eof) then 
								response.write rsObra("CDOBRA") & " - " & rsObra("DSOBRA")
							end if						
						else
							response.write GF_TRADUCIR("Sin Ptda. Presupuestaria")
						end if
					%>
					<input type="hidden" name="idObra" id="idObra" value="<% =pct_idObra %>">
					<% end if %>		
					<td colspan="3">
						<div id="areadetalle"></div>
					</td>
				</td>				
			</tr>			
			<tr>
				<td class="reg_header_navdos"><% =GF_TRADUCIR("Solicitante") %></td>
				<td colspan="2">
					<% if (esModificable) then %>
						<div id="divSolicitante"></div>																		
					<% else
						response.write pct_dsSolicitante
					end if %>															
					<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =pct_cdSolicitante %>"/>
				</td>
				<td class="reg_header_navdos"><% =GF_TRADUCIR("Division") %></td>
				<td colspan="2">
				<% 	
				
				if ((esModificable) and (pct_idPedido = 0)) then 				
						strSQL="Select * from TBLDIVISIONES"
						Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
						%>
							<select id="idDivision" name="idDivision">
								<option value="<% =SIN_DIVISION %>" selected="true">- <% =GF_TRADUCIR("Seleccione") %> -
								<%		
								while (not rsDivision.eof) 	
									if (checkPointAcceso(rsDivision("IDDIVISION"))) then
										if not isAuditor(rsDivision("IDDIVISION")) then %>
											<option value="<% =rsDivision("IDDIVISION") %>" <% if (pct_idDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
								<%		end if
									end if
									rsDivision.MoveNext()
								wend	
								%>								
							</select>
				<% 	else
						response.write pct_dsDivision	%>										
						<input type="hidden" id="idDivision" name="idDivision" value="<% =pct_idDivision %>" />
				<%	end if %>															
			</tr>
			<tr>				
				<td class="reg_header_navdos" width="15%"><% =GF_TRADUCIR("Fecha Emisión") %></td>
				<td width="20%" align="center">																		
					<div id="issuedateDiv" class="labelStyle"><% =pct_FechaInicio %></div>															
					<input type="hidden" id="issuedate" name="issuedate" value="<% =pct_FechaInicio %>"/>										
				</td>
				<td width="10%">
				<% if (esModificable) then%>
					<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCalEmision)"><img id="imgEmision" src="images/compras/calendar-16x16.png"></a>
				<% end if %>
				</td>				
				<td width="15%" class="reg_header_navdos"><% =GF_TRADUCIR("Fecha Limite") %></td>
				<td width="20%" align="center">					
					<div id="closingdateDiv" class="labelStyle"><% =pct_FechaCierre %></div>	
					<input type="hidden" id="closingdate" name="closingdate" value="<% =pct_FechaCierre %>" />					
				</td>
				<td width="10%">
				<% 	if ((esModificable) or (cambiaPlazo)) then 	%>
						<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/compras/calendar-16x16.png"></a>
				<%  end if 	%>
				</td>				
			</tr>
			
			<tr>
			
			</tr>	
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Archivos") %></td>
			</tr>
			<tr>
				<td class="reg_header_navdos" colspan="2">
					<% =GF_TRADUCIR("Especificación técnica") %>
				</td>				
				<td colspan=3>
					<table><tr><td>
						<%response.write showUploadFiles(hayET,(esModificable or cambiaPlazo),"fileEspec","ET")%>
					</td></tr></table>
				</td>
			</tr>
			<tr>
				<td class="reg_header_navdos" colspan="2">
					<% =GF_TRADUCIR("Condiciones Particulares") %>
				</td>				
				<td colspan=3>
					<table><tr><td>
						<%response.write showUploadFiles(hayCP,(esModificable or cambiaPlazo),"fileReglamento","CP")%>
					</td></tr></table>
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Titulo") %></td>
			</tr>
			<tr>
				<td colspan="6">
				<% if (esModificable) then %>					
					<input type="text" name="titulo" maxLength="100" size="100" id="titulo" value="<% =pct_tituloPedido %>">
				<% else %>
					<% =pct_tituloPedido %>
					<input type="hidden" name="titulo" id="titulo" value="<% =pct_tituloPedido %>">
				<% end if %>
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Descripción (Max 4000 caracteres)") %></td>
			</tr>
			<tr>
				<td colspan="6">
				<% if (esModificable) then %>					
					<textarea type="text" name="description" maxLength="4000" cols="100" rows="5" id="description"><% =pct_dsPedido %></textarea>
				<% else %>
					<% =pct_dsPedido %>
					<input type="hidden" name="description" id="description" value="<% =pct_dsPedido %>">
				<% end if %>
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Proveedores") %></td>
			</tr>
			<% if (esModificable) then %>
			<tr>
				<td colspan="6" style="background-color:#ffff99" valign=middle>
					<img src="images/compras/warning-16x16.png" align="absMiddle">&nbsp;<b><% =GF_TRADUCIR("No olvide que debe proponer al menos " & minCantPro & " proveedores, si son menos justifiquelo en las observaciones") %>.</b>
				</td>
			</tr>
			<% end if  %>			
			<tr><td colspan="6">
				<table class="reg_header" width="100%" id="tblProveedores">
					<tr class="reg_header_nav">
						<td width="10%"><% =GF_TRADUCIR("Codigo") %></td>
						<td width="85%"><% =GF_TRADUCIR("Descripcion") %></td>						
						<td width="5%" class="reg_header" colspan="2" align="center">
						<%	if (aceptaProveedor) then %>
							<img src="images/compras/add-16x16.png" onClick="agregarLineaProveedor(true);" style="cursor:pointer">
						<%	end if %>
						</td>
					</tr>
				</table>
			</td></tr>			
			<tr>
				<td class="reg_header_nav" colspan="6"><% =GF_TRADUCIR("Observaciones") %></td>
			</tr>
			<tr>
				<td colspan="6">
				<% if (esModificable) then %>
					<textarea name="observaciones" cols="100" rows="5" maxlength="1000"><% =pct_observaciones %></textarea>
				<% else 
					if (pct_observaciones <> "") then 
						Response.write pct_observaciones 
					else
						Response.write GF_TRADUCIR("No hay observaciones.")
					end if
				%>
					<input type="hidden" name="observaciones" id="observaciones" value="<% =pct_observaciones %>">
				<% end if %>
				</td>
			</tr>			
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Administra")%>:</td>
				<td colspan="5"><% 	=LICITACIONES_ARGENTINA %>				
				</td>
			</tr>
		</table>		
		<table width="80%" align="center">
			<tr>
				<td align="right">
					<font class="smaller">
					<%	if (pct_idPedido > 0) then %>
							<% =GF_TRADUCIR("Cargó ") & pct_usuarioCarga & " el dia " & GF_FN2DTE(pct_momentoCarga)%>
						<% if (CDbl(pct_momento) <> CDbl(pct_momentoCarga)) then %>
							<% =" | " & GF_TRADUCIR("Modificó ") & pct_usuario & " el dia " & GF_FN2DTE(pct_momento)%>
					<%		end if 
						end if	%>
					</font>
				</td>
			</tr>
		</table>	
		
		<input type="hidden" id="accion" name="accion" value="">
		<input type="hidden" id="idPedido" name="idPedido" value="<% =idPedido %>">
		<input type="hidden" id="cdPedido" name="cdPedido" value="<% =pct_cdPedido %>">
		<input type="hidden" id="idEstado" name="idEstado" value="<% =pct_idEstado %>">		
		<input type="hidden" id="etFile" name="etFile" value="<% =especifTecnica %>">
		<input type="hidden" id="rgFile" name="rgFile" value="<% =condParticulares %>">
		<input type="hidden" id="cantProveedores" name="cantProveedores"  value="0">
		<input type="hidden" id="provExistentes" name="provExistentes"  value="<% =cantprovExistentes %>">		
		<input type="hidden" id="momentoCarga" name="momentoCarga"  value="<% =pct_momentoCarga%>">
		<input type="hidden" id="cdUsuarioCarga" name="cdUsuarioCarga"  value="<% =pct_usuarioCarga%>">		
        <input type="hidden" id="fechaInicioOld" name="fechaInicioOld" value="<%=fechaInicioOld %>"/>
        <input type="hidden" id="fechaCierreOld" name="fechaCierreOld" value="<%= fechaCierreOld%>"/>
        <input type="hidden" id="cdSolicitanteOld" name="cdSolicitanteOld" value="<%= cdSolicitanteOld%>"/>
        <input type="hidden" id="tituloPedidoOld" name="tituloPedidoOld" value="<%= tituloPedidoOld%>"/>
        <input type="hidden" id="idObraOld" name="idObraOld" value="<%= idObraOld%>"/>
        <input type="hidden" id="dsPedidoOld" name="dsPedidoOld" value="<%= dsPedidoOld%>"/>
        <input type="hidden" id="observacionesOld" name="observacionesOld" value="<%= observacionesOld%>"/>
        <input type="hidden" id="idAreaOld" name="idAreaOld" value="<%= idAreaOld%>"/>
        <input type="hidden" id="idDetalleOld" name="idDetalleOld" value="<%= idDetalleOld%>"/>
        <input type="hidden" id="idDivisionOld" name="idDivisionOld" value="<%= idDivisionOld%>"/>
        <input type="hidden" id="extension" name="extension" value="<%= pct_Extension%>"/>
	</form>
</body>
</html>