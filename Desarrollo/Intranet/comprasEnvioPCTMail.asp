<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientoscompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/md5.asp"-->
<% 

Call comprasControlAccesoCM(RES_CC)


Function mostrarProveedor(pIdProveedor, pDsProveedor, pEmailProveedor, pAccion, pIndex)
%>
		<tr class="reg_Header_navdos">
				<td>
					<% =pIdProveedor %>-<% =pDsProveedor %>
					<input type="hidden" id="idProveedor_<% =pIndex %>" value="<% =pIdProveedor %>">
				</td>
				<td>
					<div id="txtmail_<% =pIndex %>"><% =pEmailProveedor %></div>
					<input type="hidden" id="actualMail_<% =pIndex %>" value="<% =pEmailProveedor %>">
					<input type="text" id="inputMail_<% =pIndex %>" style="display:none;" size="40" value="<% =pEmailProveedor %>">
				</td>
				<td>
				<div id="imgMail_<% =pIndex %>" style="cursor:pointer;"><img src="images/compras/edit-16x16.png" onClick="editMail(<% =pIndex %>)"></div></td>
				<td align="center">
				<% 	mostrarProveedor = false
				    Select case pAccion
				        case ACCION_EMAIL							
						    if (enviarMailProveedor(pIdProveedor, pEmailProveedor)) then	
							    mostrarProveedor = true							
				    %>			<img src="images/compras/action_ok-16x16.png"> 
				    <%		else	
				    %>			<img src="images/compras/action_error-16x16.png">
				    <%		end if
			            case ACCION_ACTIVAR
			                mostrarProveedor = true
				    %>			<img src="images/compras/mail_sent-16x16.png"> 
				<%	End Select  %>
				</td>
			</tr>
<%
End Function
'-------------------------------------------------------------------------------------------
Function generateAttachmentPCT(pRutaEspTecnica,pRutaParticulares,pRutaGenerales,pRutaContratista)
    Dim auxAtt
    if (Trim(pRutaEspTecnica) <> "") then auxAtt = auxAtt & pRutaEspTecnica & ";"
    if (Trim(pRutaParticulares) <> "") then auxAtt = auxAtt & pRutaParticulares & ";"
    if (Trim(pRutaGenerales) <> "") then auxAtt = auxAtt & pRutaGenerales & ";"
    if (Trim(pRutaContratista) <> "") then auxAtt = auxAtt & pRutaContratista & ";"
    auxAtt = left(auxAtt,len(auxAtt)-1)
    generateAttachmentPCT = auxAtt
End Function
'-------------------------------------------------------------------------------------------
Function enviarMailProveedor(pIdProveedor, emailProveedor)
	Dim strMsg, idUsuario, ds, flagFile,rutaEspTecnica,rutaParticulares,rutaGenerales,rutaContratista,auxPermiso,emailToepfer, letraDivision
	strMsg = ""
	enviarMailProveedor = false	
	emailToepfer = SENDER_COTIZACIONES
	if ((emailProveedor <> "") and (emailToepfer <> "")) then
		'Primero obtengo todos los archivos que estan adjuntos al Pedido para aclarar en el cuerpo del mail en que consiste cada uno
        rutaEspTecnica = ""
        if (hayEspecifTecnica(pct_idPedido)) then rutaEspTecnica = pctBinary2File(pct_idPedido, PCT_BINARY_SPECIFICATION, server.MapPath(PATH_COMPRAS_FINAL))
		rutaParticulares = ""
        if (hayCondParticulares(pct_idPedido)) then rutaParticulares = pctBinary2File(pct_idPedido, PCT_BINARY_CONDITIONS, server.MapPath(PATH_COMPRAS_FINAL))
        call setPaths(pct_idPedido, "Condiciones_Generales_Obra.doc", rutaGenerales, rutaGeneralesWeb)                
        letraDivision = getDivisionAbreviada(pct_idDivision)
        if (letraDivision <> CODIGO_EXPORTACION) then Call setPaths(pct_idPedido, "Manual_Contratista_" & letraDivision & ".zip", rutaContratista, rutaContratistaWeb)        
        strMsg = strMsg & "Se encuentra abierta la licitación para una importante obra." & vbCrLf & vbCrLf
        strMsg = strMsg & "REF: " & pct_cdPedido & " - " & pct_tituloPedido & vbCrLf
		strMsg = strMsg & "Division: " & pct_dsDivision & vbCrLf 
		strMsg = strMsg & "Fecha de Publicación: " & pct_FechaInicio & vbCrLf
		strMsg = strMsg & "Se recibirán cotizaciones hasta el: " & pct_FechaCierre & vbCrLf & vbCrLf
        if (Trim(pct_dsPedido) <> "") then
            strMsg = strMsg & "Detalle:" & vbCrLf
		    strMsg = strMsg & Replace(pct_dsPedido,ENTER_SYMBOL,vbCrLf) & vbCrLf
        end if
        strMsg = strMsg & "Para mayor detalle revisar los archivos adjuntos:" & vbCrLf
        strMsg = strMsg & " - " & Right(rutaGenerales, Len(rutaGenerales) - InStrRev(rutaGenerales, "\")) & GF_TRADUCIR(" (Condiciones generales)") & vbCrLf
        if (Trim(rutaContratista) <> "") then strMsg = strMsg & " - " & Right(rutaContratista, Len(rutaContratista) - InStrRev(rutaContratista, "\")) &GF_TRADUCIR(" (Considerar solo si realiza trabajos en nuestra terminal)") & vbCrLf
        if (Trim(rutaEspTecnica) <> "") then strMsg = strMsg & " - " & Right(rutaEspTecnica, Len(rutaEspTecnica) - InStrRev(rutaEspTecnica, "\")) & GF_TRADUCIR(" (Especificaciones técnicas)") & vbCrLf
        if (Trim(rutaParticulares) <> "") then strMsg = strMsg & " - " & Right(rutaParticulares, Len(rutaParticulares) - InStrRev(rutaParticulares, "\")) & GF_TRADUCIR(" (Condiciones particulares)") & vbCrLf
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "MUY IMPORTANTE" & vbCrLf
        strMsg = strMsg & "-----------------------" & vbCrLf
        strMsg = strMsg & "Para participar de la licitación, por favor responder ÚNICAMENTE ESTE MAIL adjuntando su cotización antes de la fecha de cierre." & vbCrLf        
        strMsg = strMsg & "Por normas de Auditoría Interna NO SE ACEPTARÁN COTIZACIONES QUE NO FUERAN ENVIADAS POR EL MEDIO ARRIBA DETALLADO." & vbCrLf        
        strMsg = strMsg & "ES IMPORTANTE QUE EL MAIL CON LA COTIZACIÓN CONTENGA LOS CÓDIGOS DEBAJO DETALLADOS YA QUE LOS MISMOS IDENTIFICAN AL PROVEEDOR Y AL PEDIDO. LA FALTA DE LOS MISMOS PROVOCA QUE EL SISTEMA RECHACE LA COTIZACION. " & vbCrLf
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "Para realizar consultas o comentarios por favor escriba a ArgentinaCompras@adm.com." & vbCrLf & vbCrLf
        strMsg = strMsg & "Lo saluda atte." & vbCrLf & vbCrLf
        strMsg = strMsg & "Dto. de  Compras" & vbCrLf
        strMsg = strMsg & "ADM AGRO S.R.L." & vbCrLf & vbCrLf
        strMsg = strMsg & generarSeccionServicioCompras(pIdProveedor, pct_idPedido, "F")        
		strMsg = strMsg & GF_TRADUCIR("AVISO LEGAL:") & vbcrlf & vbcrlf & GF_TRADUCIR("Se notifica a quienes sean invitados a participar y a quienes efectivamente participen en un concurso de precios, que toda la información suministrada por ADM AGRO S.R.L., como así también la información suministrada por las empresas participantes y por quien resulte la empresa proveedora, es de carácter estrictamente confidencial, motivo por el cual la comunicación a terceros, reproducción por cualquier medio, o cualquier otro uso o diseminación de dicha información está prohibida y será considerada ilegal. La invitación a participar en el concurso de precios no representa compromiso alguno por parte de Alfred C. Toepfer International Argentina SRL de adquirir el producto o servicio de que se trate, ni ningún otro. Las invitaciones son de carácter exclusivo para la/s empresa/s que las reciban, y no pueden ser trasladas a terceros. En el caso de que el proveedor tercerice el servicio, Alfred C. Toepfer International Argentina SRL se reserva el derecho de descartar la cotización presentada, o en su caso cancelar la adjudicación efectuada. Alfred C. Toepfer International Argentina SRL  en ningún caso será responsable por problemas técnicos que impidan la presentación de la documentación requerida a través del sitio en los plazos establecidos por el concurso.")		
        Call GP_ENVIAR_MAIL_ATTACHMENT(GF_TRADUCIR("Solicitud de Cotizacion") & " - REF: " & pct_cdPedido, strMsg, emailToepfer, emailProveedor, generateAttachmentPCT(rutaEspTecnica,rutaParticulares,rutaGenerales,rutaContratista))				
        'Se elimina el archivo creado temporalmente para adjuntar al mail.
		Set fso = CreateObject("Scripting.FileSystemObject")
		if (Trim(rutaParticulares) <> "") then fso.DeleteFile(rutaParticulares)
        if (Trim(rutaEspTecnica) <> "") then fso.DeleteFile(rutaEspTecnica)
		set fs=nothing
		enviarMailProveedor = true
	end if	
End Function
'------------------------------------------------------------------------------
Function enviarMailToepfer()
	Dim strMsg, idUsuario, ds, emailToepfer
	
	emailToepfer = SENDER_COTIZACIONES
	enviarMailToepfer = false
	if (emailToepfer <> "") then		
		strMsg = "Se ha publicado el pedido de cotización " & pct_cdPedido & vbCrLf
		strMsg = strMsg & vbCrLf & vbCrLf
		strMsg = strMsg & "Datos del Pedido" & vbCrLf
		strMsg = strMsg & "----------------" & vbCrLf
		strMsg = strMsg & "Codigo asignado.....: " & pct_cdPedido & vbCrLf		
		strMsg = strMsg & "Titulo..............: " & pct_tituloPedido & vbCrLf
		strMsg = strMsg & "Solicitante.........: " & pct_cdSolicitante & "-" & pct_dsSolicitante & vbCrLf		
		strMsg = strMsg & "Tipo de Pedido......: Pedido de Precios" & vbCrLf		
		strMsg = strMsg & "Fecha de Limite.....: " & pct_FechaCierre & vbCrLf
		
        strMsg = strMsg & "Descripcion: " & vbCrLf & Replace(pct_dsPedido,ENTER_SYMBOL,vbCrLf) & vbCrLf        
		strMsg = strMsg & "Proveedores Involucrados: " & vbCrLf
		if (initProveedores()) then
			while (readNextProveedor())
				strMsg = strMsg & pct_idProveedor & " - " & pct_dsProveedor & vbCrLf
				strMsg = strMsg & "Mail: " &pct_emailProveedor & vbCrLf & vbCrLf
			wend
		end if
		Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web - Alerta Pedido Publicado") & ": " & pct_cdPedido, strMsg, emailToepfer, obtenerMail(CD_TOEPFER))
		enviarMailToepfer = true
	end if	
End Function
'-------------------------------------------------------------------------------------------
Function marcarEnvio(idPedido)
	Dim strSQL, conn, rs, flagPublicado
	flagPublicado=false
    strSQL="Select * from TBLPCTCABECERA where IDPEDIDO=" & idPedido & " and ESTADO<" & ESTADO_PCT_PUBLICADO    
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
	    'Se actualiza el estado del pedido para marcar que ya fue publicado.
	    strSQL="Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_PUBLICADO & " where IDPEDIDO=" & idPedido
	    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	    flagPublicado = true
    end if	    
    marcarEnvio = flagPublicado
End Function
'---------------------------------------------------------------------------------------------
'Permite controlar si los datos a los que se desea acceder son correctos
' El control sobre el estado es opcional, si no se desea controlar, pasarlo vacio.
Function validarAccesoPCT(pIdPedido, pIdProveedor, pFromIdEstado, pToIdEstado)
	Dim idPedido, idProveedor, permiso, validaDesde, validaHasta, sistema
	Dim result, rs, strSQL
	
	result = false
	'1.- Se valida que el pedido exista, que el estado sea el correcto y que no haya cerrado
	strSQL = "Select * from TBLPCTCABECERA where IDPEDIDO = " & pIdPedido
	if (pFromIdEstado <> "") then strSQL = strSQL & " and ESTADO >= " & pFromIdEstado
	if (pToIdEstado <> "") then strSQL = strSQL & " and ESTADO <= " & pToIdEstado
	'response.write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
	'response.write "VAPCT-2<br>"
		'2.- Se valida que el proveedor se encuentre asociado al pedido.
		if (pIdProveedor > 0) then
			strSQL = "Select * from TBLPCTPROVEEDORES where IDPEDIDO = " & pIdPedido & " and IDPROVEEDOR = " & pIdProveedor		
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (not rs.eof) then 			
				result = true
			end if
		else
			result = true
		end if
	end if
	validarAccesoPCT = result
End Function
'*********************************************************
'******	COMIENZO DE LA PAGINA
'*********************************************************
Dim idProveedor, dsProveedor, emailProveedor, accion, idPedido, auxIndex

idPedido = GF_PARAMETROS7("idPedido", 0, 6)
idProveedor = GF_PARAMETROS7("idProveedor", 0, 6)
accion = GF_PARAMETROS7("accion", "", 6)

Call initHeader(idPedido)
if (not checkControlPCT()) then
	'No puede acceder, se lo envia a la pagina de error.
	response.redirect "comprasAccesoDenegado.asp"
end if

auxIndex=0

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript">
	var ch = new channel();
	function saveMail(pIndex) {
		var actualMail = document.getElementById("actualMail_" + pIndex).value;
		var newMail = document.getElementById("inputMail_" + pIndex).value;
		var idProv = document.getElementById("idProveedor_" + pIndex).value;
		if ((actualMail != newMail) && (newMail != '')) {
			document.getElementById("imgMail_" + pIndex).innerHTML = '<img src="images/loading_small_green.gif">';
			ch.bind("comprasActualizarMail.asp?idEmpresa=" + idProv + "&mail=" + newMail, "saveMailCallBack("+pIndex+")");
			ch.send();
		} else {
			alert('No se produjo ningun cambio en el mail');
			updateChange(pIndex, actualMail);
		}
	}
	function saveMailCallBack(pIndex) {
		var actualMail = document.getElementById("actualMail_" + pIndex).value;
		var newMail = ch.response();
		if (newMail != '') {
			updateChange(pIndex, newMail);
		} else {
			alert('El mail ingresado es invalido');
			updateChange(pIndex, actualMail);
		}
	}
	function updateChange(pIndex, mail) {
		document.getElementById("inputMail_" + pIndex).value = mail;
		document.getElementById("inputMail_" + pIndex).style.display = 'none';
		document.getElementById("txtmail_" + pIndex).innerHTML = mail;
		document.getElementById("txtmail_" + pIndex).style.display = 'block';
		document.getElementById("actualMail_" + pIndex).value = mail;
		document.getElementById("imgMail_" + pIndex).innerHTML = '<img src="images/compras/edit-16x16.png" onClick="editMail('+pIndex+')">';
	}
	function editMail(pIndex) {
		document.getElementById("txtmail_" + pIndex).style.display = 'none';
		document.getElementById("inputMail_" + pIndex).style.display = 'block';
		document.getElementById("imgMail_" + pIndex).innerHTML = '<img src="images/compras/save-16x16.png" onClick="saveMail('+pIndex+')">';
    }
    function skip() {        
        document.getElementById("accion").value = "<% =ACCION_ACTIVAR %>";        
        document.getElementById("frmSel").submit();
    }
    function enviar() {
	document.getElementById("btnEnviar").value="Enviando...";
	document.getElementById("frmSel").submit();
    }
</script>
<body>
<form name="frmSel" id="frmSel" action="comprasEnvioPCTMail.asp" method="POST">
	<table class="reg_Header" width="100%" align="center">		
		<tr class="reg_Header_nav">
			<td><% =GF_TRADUCIR("Proveedor") %></td>
			<td colspan="2" width="2%"><% =GF_TRADUCIR("Email") %></td>
			<td><% =GF_TRADUCIR("Status") %></td>			
		</tr>
<%
	flagPublicacion = false
	if (idProveedor = 0) then
		'Se envian los mail a todos los proveedores del pedido.
		if (initProveedoresDB()) then				
			while (readNextProveedorDB())			
				flagPublicacion = (flagPublicacion or mostrarProveedor(pct_idProveedor, pct_dsProveedor, pct_emailProveedor, accion, auxIndex))
				auxIndex = auxIndex + 1
			wend
			
		end if	
	else
		'Se envia solo a los proveedores solicitados.
		dsProveedor = getDescripcionProveedor(idProveedor)
		emailProveedor = obtenerMail(idProveedor)
		flagPublicacion = mostrarProveedor(idProveedor, dsProveedor, emailProveedor, accion, auxIndex)
	end if
	if (flagPublicacion) then		 
		 'Solo se notifica del envío si realmente fueron enviados.
		 if (marcarEnvio(pct_idPedido)) then Call enviarMailToepfer()
	end if
 if (pct_idEstado <= ESTADO_PCT_PUBLICADO) then
	
    	if (accion <> ACCION_EMAIL) then
%>		
		<tr><td align="center" colspan="4"><input type="button" id="btnEnviar" onclick="enviar()" value="<% =GF_TRADUCIR("Enviar") %>"></td></tr>
		<tr><td align="right" colspan="4"><a href="javascript:skip()"><% =GF_TRADUCIR("Los mails ya fueron enviados.") %></a></td></tr>
<%      
        end if %>
	    <tr>
	        <td colspan="4">
	            <img width="10px" height="10px" src="images/compras/action_ok-16x16.png"> <% =GF_TRADUCIR("El mail fue enviado con exito") %> | <img width="10px" height="10px" src="images/compras/action_error-16x16.png"> <% =GF_TRADUCIR("El mail no puedo ser enviado") %>
	            <input type="hidden" name="accion" id="accion" value="<% =ACCION_EMAIL %>">
	            <input type="hidden" name="idPedido" value="<% =idPedido %>">
	            <input type="hidden" name="idProveedor" value="<% =idProveedor %>">	
	        </td>
	    </tr>			
<% else	%>
	<table class="reg_Header" width="100%" align="center">
		<tr><td align="center"><% =GF_TRADUCIR("El pedido ya se encuentra cerrado y no pueden enviarse los mails.") %></td></tr>
<% end if %>
    </table>
</form>
</body>
</html>