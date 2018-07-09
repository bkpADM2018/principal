<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/md5.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Function controlAcceso(pIdPedido, pIdProveedor, pasaporte, ByRef pPuedeNegarse)
	Dim rs,saltarFecha, payLoad
	controlAcceso = false
	puedeNegarse = true
	'Se controla si tiene acceos a la información
	controlAcceso = validarPasaporteCompras(pIdPedido, pIdProveedor, pasaporte)
	if (controlAcceso) then						
			controlAcceso = false
			'El pedido ya cerro, si tiene permiso, ignora esta fecha de cierre.
			if (retrievePayload(pasaporte, payLoad)) then			
					'ARREGLO TEMPORAL - Esta clave es nueva y los pasaporte ya emitidos no la tienen, en un tiempo breve puede eliminarse esta pregunta.
					if (payLoad.Exists("ACTI")) then										
						esActiUser = payLoad("ACTI")
						if (payLoad("ACTI")) then							
							'Es un usuario que accede con unpasaporte emitido para un usuario de TOEPFER
							if (payLoad("SKIP")) then	
								controlAcceso = true
								esAuditorCoordinador = true
							end if
						else
							'Es un usuario que accede con unpasaporte emitido para un usuario externo							
							if (clng(GF_DTE2FN(pct_FechaCierre)) >= clng(left(session("mmtoSistema"),8))) then controlAcceso = true
						end if
					else					
						'Es un usuario que accede con unpasaporte emitido para un usuario externo
						esActiUser = false
						if (clng(GF_DTE2FN(pct_FechaCierre)) >= clng(left(session("mmtoSistema"),8))) then controlAcceso = true
					end if
			end if		
		if (controlAcceso) then			
			'Se controla si ya se nego a cotizar.
			Set rs = getCotizaciones(pIdPedido, pIdProveedor)
			if (not rs.eof) then			
				if (rs("PATHCOTIZACION") = ACCION_PCT_RETIRARSE )then					
					controlAcceso = false
				else
					'Ya hay cotizaciones, no puede negarse a cotizar.
					puedeNegarse = false
				end if
			end if
		end if		
	end if		
End Function
'---------------------------------------------------------------------------------------------
'Permite controlar si los datos a los que se desea acceder son correctos
' El control sobre el estado es opcional, si no se desea controlar, pasarlo vacio.
Function validarAccesoPCT(pIdPedido, pIdProveedor, pFromIdEstado, pToIdEstado)
	Dim idPedido, idProveedor, permiso, validaDesde, validaHasta, sistema
	Dim result, rs, strSQL
	
	'1.- Se valida que el pedido exista, que el estado sea el correcto y que no haya cerrado
	strSQL = "Select * from TOEPFERDB.TBLPCTCABECERA where IDPEDIDO = " & pIdPedido
	if (pFromIdEstado <> "") then strSQL = strSQL & " and ESTADO >= " & pFromIdEstado
	if (pToIdEstado <> "") then strSQL = strSQL & " and ESTADO <= " & pToIdEstado
	'response.write strSQL
	Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)	
	if (not rs.eof) then
	'response.write "VAPCT-2<br>"
		'2.- Se valida que el proveedor se encuentre asociado al pedido.
		if (pIdProveedor > 0) then
			strSQL = "Select * from TOEPFERDB.TBLPCTPROVEEDORES where IDPEDIDO = " & pIdPedido & " and IDPROVEEDOR = " & pIdProveedor		
			Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)	
			if (not rs.eof) then 			
				result = true
			end if
		else
			result = true
		end if
	end if
	validarAccesoPCT = result
End Function
'-----------------------------------------------------------------------------------
Function enviarMail(pIdProveedor, asunto, strMsg, mailFrom, mailTo)
					
	if (mailTo <> "") then		
		Call GP_ENVIAR_MAIL(asunto, strMsg, mailFrom, mailTo)
	end if	
End Function
'-----------------------------------------------------------------------------------
Function grabarDatos(pIdProveedor, pFile)
		Dim strSQL, rsCotizacion, connCotizacion
		
		strSQL = "Insert into TBLPCTCOTIZACIONES(IDPEDIDO, IDPROVEEDOR, PATHCOTIZACION, FECHAPRESENTACION) values (" & pct_idPedido & ", " & pIdProveedor & ", '" & pFile  & "', " & session("MmtoSistema") & ")"		
		
		if ((pct_idEstado >= ESTADO_PCT_ABIERTO) and pct_tipoCompra = TIPO_PCT_CONCURSO ) then
				strSQL = "Insert into TBLPCTCOTIZACIONES(IDPEDIDO, IDPROVEEDOR, PATHCOTIZACION, FECHAPRESENTACION, CDUSRAPERTURA, FECHAAPERTURA) values (" & pct_idPedido & ", " & pIdProveedor & ", '" & pFile  & "', " & session("MmtoSistema") & ",'" & FIRMA_NO_USER & "'," & session("MmtoSistema") & ")"		
		end if
	
		Call executeQueryDb(DBSITE_SQL_INTRA, rsCotizacion, "EXEC", strSQL)
		if (pFile <> ACCION_PCT_RETIRARSE) then
			strSQL = "Select MAX(IDCOTIZACION) as IDCTZ from TBLPCTCOTIZACIONES"
			Call executeQueryDb(DBSITE_SQL_INTRA, rsCotizacion, "OPEN", strSQL)
			if isnull(rsCotizacion("IDCTZ")) then 
				grabarDatos = 1
			else	
				grabarDatos = rsCotizacion("IDCTZ")
			end if
		else
			grabarDatos = 0
		end if

End Function
'-----------------------------------------------------------------------------------
Function accionCotizar(pFolder, pIdProveedor, pDirectory)
	Dim thePath, theFil, strMsg, emailToepfer, mailProveedor, fileno,myFile,i
	accionCotizar = false
	set FSO = server.createObject("Scripting.FileSystemObject") 
	
	if (pFolder <> "" and FSO.FolderExists(pFolder)) then
		Set carpeta = FSO.GetFolder(pFolder) 
		Set archivos = carpeta.Files 
		for each archivo in archivos
			'Se registra la cotizacion
			fileno = grabarDatos(pIdProveedor, archivo.Name)
			'Se digitaliza la info.		
			thePath = archivo
			
			Call pctGrabaArchivo(pct_idPedido, thePath, fileno)
			
			'Se notifica por mail
			'Si la compra en por concurso al administrdor, sino se notifica al usuario que cargo el pedido.
			'JAS -if (pct_tipoCompra = TIPO_PCT_COMPARATIVA) then emailToepfer = getUserMail(pct_cdUsuarioAdmin)
			if (emailToepfer = "") then emailToepfer = obtenerMail(CD_TOEPFER)
			strMsg = "El proveedor " & pIdProveedor & "-" & Trim(getDescripcionProveedor(pIdProveedor)) & " ha publicado una cotizacion para el pedido " & pct_cdPedido & "-" & Trim(pct_tituloPedido)	
			Call enviarMail(pIdProveedor, GF_TRADUCIR("Sistema de Compras Web - Alerta Cotizacion Publicada"), strMsg, emailToepfer, emailToepfer)
			
			'Se notifica al proveedor
			mailProveedor = obtenerMail(pIdProveedor)
			strMsg = "Su cotizacion para el pedido '" & pct_cdPedido & "-" & Trim(pct_tituloPedido) & "' ha sido recibida correctamente." & vbCrLf & vbCrLf
			strMsg = strMsg & "Muchas Gracias"			
			Call enviarMail(pIdProveedor, GF_TRADUCIR("Sistema de Compras Web - Cotizacion Recibida"), strMsg, emailToepfer, mailProveedor)
			accionCotizar = true
		next
	end if
End Function
'-----------------------------------------------------------------------------------
Function deleteFolder(pDirectory)
	dim fs, dltFolder
	Set fs = CreateObject("Scripting.FileSystemObject")
	if ((fs.FolderExists(pDirectory)) and (pDirectory <> "")) then
		fs.deleteFolder(pDirectory)
	end if
End Function
'-----------------------------------------------------------------------------------
'*********************************************************
'***   COMIENZO DE LA PAGINA
'*********************************************************
Dim idPedido, accion, path, msg, clsMsg, puedeNegarse, colspan
Dim folder, emailToepfer, myFechaCierre, path2, pathWeb2, path3, pathWeb3, provPath, directory
Dim esAuditorCoordinador, payload, esActiUser

call GP_ConfigurarMomentos()
accion = GF_PARAMETROS7("accion", "", 6)
folder = GF_PARAMETROS7("ctFolder", "", 6)
idPedido = GF_PARAMETROS7("idPedido", 0, 6)
idProveedor = GF_PARAMETROS7("idProveedor", 0, 6)

provPath = PATH_COMPRAS_TEMP & "\\" & idProveedor
directory = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & idProveedor
files_fileCotizacion = GF_PARAMETROS7("files_fileCotizacion","",6)
esAuditorCoordinador = false

Call initHeader(idPedido)
Select case accion
	case ACCION_PCT_COTIZAR:
		clsMsg = "TDERROR"
		if (accionCotizar(directory, idProveedor, directory)) then
			msg = GF_TRADUCIR("Su cotización fue presentada exitosamente!")
			clsMsg = "msgOK"
			puedeNegarse = false
		end if
	case ACCION_PCT_RETIRARSE:
		Call grabarDatos( idProveedor, ACCION_PCT_RETIRARSE)
		strMsg = "El proveedor " & idProveedor & "-" & getDescripcionProveedor(idProveedor) & " ha decidido no participar del proceso de cotizacion para el pedido " & pct_cdPedido
		emailToepfer = obtenerMail(CD_TOEPFER)
		Call enviarMail(idProveedor, GF_TRADUCIR("Sistema de Compras Web - Proveedor no participa"), strMsg, emailToepfer, emailToepfer)
	case else
		'Se elimina la carpeta del proveedor en el directorio si existe
		Call deleteFolder(directory)
End Select
if (puedeNegarse) then colspan = "colspan='3'"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css"	 type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.msgOK {
	font-weight: bold;
	font-size: 14px;
	color: #44CC66;
}
</style>
		<script type="text/javascript" src="scripts/channel.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="scripts/JQueryUpload2.js"></script>		
		<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">
	
	function createUploader(){            
		uploader = new qq.FileUploader({
			element: document.getElementById('fileCotizacion'),
			multiple:false,
			action: 'uploadSubmitFile2.asp?accion=upload&folder=<%=provPath %>',
			allowedExtensions:["doc","pdf","rtf","xls","rar","zip", "docx", "xlsx", "msg"]			
		}); 
	}
				
	function submitForm(acc) {			
		document.getElementById("ctFolder").value = "<%=provPath %>";
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function canSubmit() {
		document.getElementById("tbLoading").style.visibility = 'visible';
		document.getElementById("btnLoading").style.visibility = 'hidden';
		submitForm('<% =ACCION_PCT_COTIZAR %>');
	}
	
	function bodyOnLoad() {
		<% if (accion <> ACCION_PCT_RETIRARSE) then %>
			createUploader();
		<% end if %>
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
		<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center">
			<tr>				
				<td <% =colspan %> class="titu_header" align="center"><b><% =GF_TRADUCIR("Cotizacion") %></b></td>
			</tr>				
			<% if (accion <> ACCION_PCT_RETIRARSE) then		%>
				<tr>				
					<td <% =colspan %> align="center">&nbsp;</td>
				</tr>
				<tr>
					<td align="center" width="50%">
						<table width="100%">
							<%	if (msg <>"") then	%>
							<tr>							
								<td  align="center" class="<% =clsMsg %>"><% =msg %></td>							
							</tr>
							<%	end if	%>	
							<tr>							
								<td width="50%"><% =GF_TRADUCIR("Adjunte el archivo de su cotización.") %>:</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td><div id="fileCotizacion"></div></td>
							</tr>
							<tr><td>(<% =GF_TRADUCIR("Solo se aceptan adjuntos de formato .doc, .pdf, .rtf, .xls, .rar y .zip.") %>)</td></tr>
							<tr><td align="center">&nbsp;</td></tr>
							<tr>
								<td align="center">
									<input type="button" id="btnLoading" onClick="javascript:canSubmit()" value="<% =GF_TRADUCIR("Presentar Cotización") %>"><br>									
									<table id="tbLoading" border="0" class="reg_header round_border_all" style="visibility:hidden">
										<tr>
											<td align="center"><strong><div>Subiendo archivo.</div><div>Aguarde por favor...</div></strong></td>
										</tr>
									</table>									
								</td>
							</tr>
						</table>
					</td>
					<% if (puedeNegarse) then %>
					<td align="center"></td>
					<td align="center" width="50%">					
						<table width="100%">
							<tr>
								<td align="center" width="50%"><% =GF_TRADUCIR("Por medio de la presente declaro nula la cotización solicitada") %>,</td>						
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td><% =GF_TRADUCIR("quedando de esta manera fuera del concurso de precios.") %></td>							
							</tr>
							<tr><td align="center">&nbsp;</td></tr>
							<tr><td align="center">&nbsp;</td></tr>
							<tr>
								<td align="center"><input type="button" onClick="javascript:submitForm('<% =ACCION_PCT_RETIRARSE %>')" value="<% =GF_TRADUCIR("Declino  Participar") %>"></td>							
							</tr>	
						</table>					
					</td>							
				<% end if %>					
				</tr>
		<%	else %>
				<tr>					
					<td <% =colspan %> align="center">
							<br><div><h5><% =GF_TRADUCIR("Ud ha sido excluido del concurso de precios.") %></h5></div>
							<br><div><h5><% =GF_TRADUCIR("Gracias por dedicar unos minutos y responder a nuestra solicitud.") %></h5></div>
					</td>
				</tr>
		<%	end if	%>						

		</table>
		<br>
		<%
		payload = ""
		Call addPayloadData(payload, "IDPEDIDO", idPedido)
		Call addPayloadData(payload, "SKIP", esAuditorCoordinador) 
		Call addPayloadData(payload, "ACTI", esActiUser)  		
		%>
		<form name="frmSel" id="frmSel" method="post" action="comprasCargarCotizacion.asp">
			<input type="hidden" id="ctFolder" name="ctFolder" value="">
			<input type="hidden" id="accion" name="accion" value="">
			<input type="hidden" id="idPedido" name="idPedido" value="<% =idPedido %>">
			<input type="hidden" id="idProveedor" name="idProveedor" value="<% =idProveedor %>">			
		</form>
</body>
</html>