<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
Const STRING_2_REPLACE = "???"
Const STRING_2_REPLACE_AD = "?#?"
'-----------------------------------------------
'ESTA PAGINA FUNCIONA COMO TABLERO DE CONTRATOS-
'DEVUELVE LA INFORMACI{ON DEL MISMO Y SUS PAGOS-
'TAMBIEN DESDE AQUI SE REALIZAN NUEVOS PAGOS   -
'-----------------------------------------------
'Function isCTCExpired : Comprueba si el contrato est� vencido o no
'			True  = Vencido
'			False = Vigente
Function isCTCExpired()
	isCTCExpired = false
	'Primero verifico si tiene una fecha de vencimiento asignada
	if (Cdbl(fechaVto) > 0) then
		'En caso de que tenga una fecha de vencimiento, compruebo que no halla expirado
		if (GF_DTEDIFF(Left(Session("MmtoDato"),8), Cdbl(fechaVto), "D") < 0) then	isCTCExpired = true
	end if
End function
'-------------------------------------------------------------------------------------------------------------
Function puedeExtenderCTC(idContrato, cdResponsable)
    
    Dim rol
        
    puedeExtenderCTC = true
    if (idContrato > 0) then
        'Hay un contrato se valida si puede extenderse.
        puedeExtenderCTC =false
        'El contrato no debe estar finalizado
        if ((CTC_estado >= ESTADO_CTC_AUTORIZADO) and (CTC_estado < ESTADO_CTC_FINALIZADO)) then 
            'El estado es valido, verifico si el usuario tiene permiso para alterar el contrato.
            rol = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
            if ((session("Usuario") = cdResponsable) or (rol = FIRMA_ROL_GTE_COMPRAS)) then
                if (not isCTCExpired()) then puedeExtenderCTC = true
            end if
        end if
    end if

End Function
'-----------------------------------------------
Function getImporteMovimiento(cdMoneda, rsPagos)
	dim factor
	factor=1
	
	if (cdMoneda = MONEDA_PESO) then
		getImporteMovimiento = factor * cDbl(rsPagos("IMPORTEPESOS"))		
	else
		getImporteMovimiento = factor * cDbl(rsPagos("IMPORTEDOLARES"))
	end if
End Function
'---------------------------------------------------
Function agregarFactura(ByRef stringPIC, idPIC, cdObra, idArea, idDetalle, observaciones, simboloMoneda, idFAC, tipoFac, porcentajePago, importeObra, importeAnticipo, importeFReparo, linea, lineaAD)
					
		Dim classNCR
			
		classNCR = ""
		if (tipoFac = PREFIX_NCR) then classNCR = "reg_header_rejected"
		
		stringPIC = stringPIC & "<tr class='TDNORMAL " & classNCR & "'>"
		
		'PIC, Observaciones y Partidas Presupuestarias
		if (linea = 1) then
			'Los textos STRING_2_REPLACE se reemplazaran, al momento de imprimir el PIC en pantalla.
			stringPIC = stringPIC & "<td align='center' rowSpan='"& STRING_2_REPLACE & "'>" & idPIC & "</td>"		
			CTC_observaciones = observaciones
			if (len(CTC_observaciones) > 50) then CTC_observaciones = left(CTC_observaciones, 50) & "..."
			'Los textos STRING_2_REPLACE se reemplazaran, al momento de imprimir el PIC en pantalla.
			stringPIC = stringPIC & "<td rowSpan='"& STRING_2_REPLACE & "'>" & CTC_observaciones & "</td>"
		end if
		
		if (lineaAD = 1) then
			'Detalle de Partida Presupuestaria.
			'Los textos STRING_2_REPLACE se reemplazaran, al momento de imprimir el PIC en pantalla.
			stringPIC = stringPIC & "<td rowSpan='"& STRING_2_REPLACE_AD & "' align='center'>" & cdObra & "-" & idArea & "-" & idDetalle & "</td>"
					
			CTC_Total_Porc = CTC_Total_Porc + porcentajePago
			if (idFAC <> 0) then CTC_Total_FAC_Porc = CTC_Total_FAC_Porc + porcentajePago			
		end if
		
		'Minuta	
		CTC_ImportePago = importeObra + importeAnticipo + importeFReparo
		
		stringPIC = stringPIC & "<td align='center'>" 
		if (idFAC <> 0) then 
		    stringPIC = stringPIC & tipoFac & " " & idFAC 
		    
		    CTC_Total_FAC_ImporteObra = CTC_Total_FAC_ImporteObra + importeObra 
		    CTC_Total_FAC_ImporteAnticipo = CTC_Total_FAC_ImporteAnticipo + importeAnticipo
		    CTC_Total_FAC_ImporteFReparo = CTC_Total_FAC_ImporteFReparo + importeFReparo
		    CTC_Total_FAC_ImportePago = CTC_Total_FAC_ImportePago + CTC_ImportePago				    
		    calculoImporteSaldoFAC = calculoImporteSaldoFAC - CTC_ImportePago
		end if
		stringPIC = stringPIC & "</td>"				
		
		CTC_Total_ImporteObra = CTC_Total_ImporteObra + importeObra 
		CTC_Total_ImporteAnticipo = CTC_Total_ImporteAnticipo + importeAnticipo
		CTC_Total_ImporteFReparo = CTC_Total_ImporteFReparo + importeFReparo				
						
		CTC_Total_ImportePago = CTC_Total_ImportePago + CTC_ImportePago		
		calculoImporteSaldoPIC = calculoImporteSaldoPIC - CTC_ImportePago
		
		'Importes
		stringPIC = stringPIC & "<td align='right'>"
		if (importeObra <> 0) then stringPIC = stringPIC & simboloMoneda & " " & GF_EDIT_DECIMALS(importeObra,2)
		stringPIC = stringPIC &	"</td>"
		
		stringPIC = stringPIC & "<td align='right'>"
		if (importeAnticipo <> 0) then stringPIC = stringPIC & simboloMoneda & " " & GF_EDIT_DECIMALS(importeAnticipo,2)
		stringPIC = stringPIC &	"</td>"			
					
		stringPIC = stringPIC & "<td align='right'>"
		if (importeFReparo <> 0) then stringPIC = stringPIC & simboloMoneda & " " & GF_EDIT_DECIMALS(importeFReparo,2)
		stringPIC = stringPIC &	"</td>"
		
		stringPIC = stringPIC & "<td align='right'>"
		if (CTC_ImportePago <> 0) then stringPIC = stringPIC & simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_ImportePago,2)
		stringPIC = stringPIC &	"</td>"
		
		if (lineaAD = 1) then
			stringPIC = stringPIC & "<td align='right' rowSpan='"& STRING_2_REPLACE_AD & "'>"
			stringPIC = stringPIC & GF_EDIT_DECIMALS(porcentajePago*100, 2) & " %"
			stringPIC = stringPIC &	"</td>"
		end if
		
		stringPIC = stringPIC & "<td align='right'>" & simboloMoneda & " " & GF_EDIT_DECIMALS(calculoImporteSaldoPIC,2) & "</td>"
				
		if (linea = 1) then
			stringPIC = stringPIC &	"<td align='center' rowSpan='"& STRING_2_REPLACE & "'>"
			stringPIC = stringPIC &	"<img style='cursor:pointer' title='" & GF_TRADUCIR("Ver Pedido Interno") & "' id='ID_" & idPIC & "' src='images\compras\CTZ-16X16.png' onclick='abrirCTZ(" & idPIC & ")'>"
			stringPIC = stringPIC &	"</td>"
		
			'Si no esta pagada, permite modificar y anular, si estan anuladas ya se filtraron anteriormente.
			if (idFAC = 0) then
				stringPIC = stringPIC &	"<td align='center' rowSpan='"& STRING_2_REPLACE & "'>"
				stringPIC = stringPIC &	"<img style='cursor:pointer' title='" & GF_TRADUCIR("Editar Pago") & "' id='ID_" & idPIC & "' src='images\compras\edit-16x16.png' onclick='editarPIC(" & CTC_idContrato & ", " & idPIC & ")'>"
				stringPIC = stringPIC &	"</td>"
				stringPIC = stringPIC & "<td align='center' rowSpan='"& STRING_2_REPLACE & "'>"
				stringPIC = stringPIC & "<img style='cursor:pointer' title='" & GF_TRADUCIR("Anular Pago") & "' id='ID_" & idPIC & "' src='images\compras\CTZ_cancel-16x16.png' onclick='anularCTZ(" & idPIC & ", " & CTC_idPedido & ", this)'>"
				stringPIC = stringPIC &	"</td>"
			else	
				stringPIC = stringPIC &	"<td rowSpan='"& STRING_2_REPLACE & "'>&nbsp;</td>"
				stringPIC = stringPIC &	"<td rowSpan='"& STRING_2_REPLACE & "'>&nbsp;</td>"
			end if	
		end if
		stringPIC = stringPIC & "</tr>"				
		
End Function		

'---------------------------------------------------
Function editAreaDetalle(stringPIC, linea)
	stringPIC = Replace(stringPIC, STRING_2_REPLACE_AD, linea)
	editAreaDetalle = stringPIC	
End Function		
'---------------------------------------------------
Function imprimirPIC(stringPIC, linea)
	stringPIC = Replace(stringPIC, STRING_2_REPLACE, linea)
	Response.Write stringPIC	
End Function		
'------------------------------------------------------------------------------------------------------------
'Dibuja dentro de la cabecera del contrato una linea con los datos de la obra
Function drawRowObra()
	Dim auxDsAreaObra, auxDsDetalleObra
 %>
	<tr>
		<td class="reg_Header_navdos" colspan="4" style="text-align: center;"><a href="Javascript:abrirPartidaCTC(<% =CTC_idContrato %>)"><% =GF_TRADUCIR("Administrar Partidas del Contrato") %></a></td>		                
	</tr><%
End function
'------------------------------------------------------------------------------------------------------------
'Envia mail indicando que se modifico el responsable del Contrato
' Parametros:  p_idContrato[int] -->  ID del Contrato
'              p_CdResponsableNew[string] -->  Nuevo responsable del contrato
'              p_CdResponsableOld[string] -->  Anterior responsable del contrato
'              p_CdContrato[string] -->  Codigo del contrato
Function enviarMailResponsableCTC(p_idContrato, p_CdResponsableNew, p_CdResponsableOld, p_CdContrato)
    Dim destino,mailMsj
    'Primero cargo los destinatarios del mail (Responsable nuevo,Responsable antiguo, Departamentos de Compras y Legales)    
    destino = getUserMail(p_CdResponsableOld) &" ; "& getUserMail(p_CdResponsableNew) &" ; "& MAILTO_COMPRAS &" ; "& SENDER_LEGALES
    mailMsj = "Se ha modificado el Responsable del contrato "& p_CdContrato & vbCrLf & vbCrLf
    mailMsj = mailMsj & "El anterior responsable era "& p_CdResponsableOld &" - "& getUserDescription(p_CdResponsableOld) &", "
	mailMsj = mailMsj & "el nuevo responsable es "& p_CdResponsableNew &" - "& getUserDescription(p_CdResponsableNew) & vbCrLf& vbCrLf
	mailMsj = mailMsj & "Usuario que modific�: " & session("Usuario") & " - " & getUserDescription(session("Usuario"))
    Call GP_ENVIAR_MAIL("Sistema de Compras Web - Modificaci�n del contrato: " & p_CdContrato, mailMsj, obtenerMail(CD_TOEPFER), destino)
End Function
'------------------------------------------------------------------------------------------------------------
Function procesarCambioResponsableCTC(p_CdResponsableNew, p_CdResponsableOld)
    Call actualizarResponsableCTC(CTC_idContrato, p_CdResponsableNew)
    Call enviarMailResponsableCTC(CTC_idContrato, p_CdResponsableNew, p_CdResponsableOld, CTC_cdContrato)
    CTC_cdResponsable = p_CdResponsableNew
    CTC_dsResponsable = getUserDescription(p_CdResponsableNew)
End function
'***********************************************
'*************  COMIENZO DE PAGINA  ************
'***********************************************
Dim cdMoneda, rsCTC, idFAC, idPIC, tipoFAC
Dim mailTo, mailMsj, calculoImporteSaldoPIC, calculoImporteSaldoFAC
Dim reg, rsPagos, params, stringPIC, idArea, idDetalle, facsAD, idArea_old, idDetalle_old
Dim idObra,accion,isPartidaModificada, myRol, importeSaldoPagar
'isPartidaModificada = false

CTC_idContrato = GF_PARAMETROS7("idContrato",0,6)

cdMoneda = GF_PARAMETROS7("cdMoneda","",6)

if (not loadContrato(CTC_idContrato, cdMoneda)) then Response.Redirect "comprasAccesoDenegado.asp"
calculoImporteSaldoPIC = CTC_TotalImporte
calculoImporteSaldoFAC = CTC_TotalImporte


fechaVto   = GF_PARAMETROS7("issuedate","",6)	
if(fechaVto = "")then fechaVto = CTC_fechaVto

accion   = GF_PARAMETROS7("accion","",6)
'SE EJECUTA EL SIGUIENTE LLAMADO POR AJAX PARA RECALCULAR EL SALDO
if (accion = ACCION_CALCULAR) then
    call ajusteSaldoPendiente(CTC_idContrato)
    response.end
end if
'Si es accion confirmar se modifica el responsable del Contrato por Ajax, ademas de informar por mail
if (accion = ACCION_CONFIRMAR) then
    cdResponsableNew = GF_PARAMETROS7("cdResponsableNew","",6)
    cdResponsableOld = GF_PARAMETROS7("cdResponsableOld","",6)
    Call procesarCambioResponsableCTC(cdResponsableNew, cdResponsableOld )
    Response.End
end if
simboloMoneda = getSimboloMoneda(CTC_cdMoneda)

myRol = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
'Se arma el link para volver de la edicion.
session("Origen") = "comprasCTC.asp?idContrato=" & CTC_idContrato & "&cdMoneda=" & CTC_cdMoneda

Set rsPagos = readCTCPagos(CTC_idContrato)
if(accion = ACCION_GRABAR)then	    
    'Con esta validacion en caso de que sea un CTC viejo(no tenga fecha Vto) se le puede agragar la fecha de vto mayor a la fecha actual
    if((Cdbl(Left(session("MmtoDato"), 8)) <= Cdbl(fechaVto))or(fechaVto = 0))then
	    Call actualizarFechaVto(CTC_idContrato ,fechaVto)
	    'Envio los mails de aviso de cambio.	    
	    mailTo = SENDER_LICITACIONES & getUserMail(CTC_cdResponsable) & SENDER_LEGALES
	    mailMsj = "Se ha modificado la fecha de vencimiento del contrato/servicio: " & CTC_cdContrato & " - " & CTC_Titulo & vbCrLf
		mailMsj = mailMsj & "Division: " & getDivisionDS(CTC_idDivision) & vbCrLf
	    mailMsj = mailMsj & "La nueva fecha de vencimiento es: " & GF_FN2DTE(fechaVto) & vbCrLf
	    mailMsj = mailMsj & vbCrLf & vbCrLf
	    mailMsj = mailMsj & "Responsable del cambio: " & session("Usuario") & " - " & getUserDescription(session("Usuario"))
	    Call GP_ENVIAR_MAIL("Sistema de Compras Web - Modificaci�n del contrato: " & CTC_cdContrato, mailMsj, obtenerMail(CD_TOEPFER), mailTo)	    
    else
	    setError(PERIODO_ERRONEO)
    end if	
end if

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>Contrato <% =CTC_cdContrato %> - Resumen de Movimientos </title>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
	<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
	<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
	
	<script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/paginar.js"></script>
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
    <script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>	
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	
	<script type="text/javascript">
		var ch = new channel();
	    var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");		
		function lightOn(tr) {
			tr.className = "reg_Header_navdosHL";
		}

		function lightOff(tr) {
			tr.className = "reg_Header_navdos";
		}
	
		function irHome() {
			location.href = "comprasIndex.asp";
		}

		function recargar() {			
		    var cdmoneda = document.getElementById("cdMoneda").options[document.getElementById("cdMoneda").selectedIndex].value;
			document.location.href = "comprasCTC.asp?idContrato=<% =CTC_idContrato %>&cdMoneda=" + cdmoneda;	
		}

		function cambiarMoneda() {
			document.getElementById("frmCTC").submit();
		}

		function mostrarAjustes() {
			document.getElementById("ajustes").innerHTML="<table align='center'><tr><td><img src='images/compras/loading_big.gif'></td></tr></table>";
			ch.bind('comprasCTCAjustesAjax.asp?id=<% =CTC_idContrato %>&cdMoneda=<% =CTC_cdMoneda %>','mostrarAjustes_callback()');
			ch.send();
		}	
		
		function mostrarAjustes_callback() {
			var resp = ch.response();
			document.getElementById("ajustes").innerHTML = resp;
		}
		
		function deleteAjuste(img, idAjuste){
			if (confirm("Esta seguro que desea eliminar este ajuste?")){
				img.src='images/loading_small_green.gif';
				ch.bind('comprasAnularAJUCTCAjax.asp?idAjuste=' + idAjuste + '&idContrato=<% =CTC_idContrato %>' ,'deleteAjuste_callback()');
				ch.send();
			}
		}
 
		function deleteAjuste_callback() {		
			mostrarAjustes();
		}
		
		function bodyOnLoad() {
			var tb = new Toolbar('toolbar', 6, 'images/compras/');
			tb.addButton("Home-16x16.png", "Home", "irHome()");
			tb.addButtonREFRESH("Recargar", "recargar()");
			tb.addButton("invoice-16.png", "Recalcular Saldo", "recalcularSaldo()");
		    <%if (CTC_estado >= ESTADO_CTC_AUTORIZADO) then %>		        
		        tb.addButton("CTC_Payment_new-16x16.png", "Agregar Pago", "nuevoPago(<% =CTC_idContrato %>)");		        
			<%  if (CTC_estado < ESTADO_CTC_FINALIZADO) then 
				    if ((canConfirmCTC(session("Usuario"), CTC_idContrato)) or (puedeExtenderCTC(CTC_idContrato, CTC_cdResponsable))) then %>
					tb.addButtonSAVE("Guardar", "Guardar()");
				<% end if
				end if %>				
			<% end if %>			
			tb.draw();			
			mostrarAjustes();
		}
		function recalcularSaldo(){
		    ch.bind("comprasCTC.asp?idContrato=<%=CTC_idContrato%>&accion=<%=ACCION_CALCULAR%>", "recalcularSaldoCallback()");
		    ch.send();
		}
		function recalcularSaldoCallback()
		{		    
		    alert ("Saldo actualizado");		 
		}
		function Guardar(){			
			document.getElementById("accion").value = '<%= ACCION_GRABAR%>';			
			document.getElementById("frmCTC").submit();
		}
		
		function irAjusteCTC() {
			var myPage, w, h;
			w=770;
			h=600;
			myPage = "comprasAjusteCTC.asp?idContrato=<% =CTC_idContrato %>";
			var puw = new winPopUp('popupAjuCTC',myPage, w, h,'Ajuste de Contrato - Total Contrato', "recargar()");
		}
		
		function irAjusteVUCTC() {
			var myPage, w, h;
			w=770;
			h=400;
			myPage = "comprasAjusteVUCTC.asp?idContrato=<% =CTC_idContrato %>";
			var puw = new winPopUp('popupAjuCTC',myPage, w, h,'Ajuste de Contrato - Valor Unitario', "recargar()");
		}
		
		function abrirAFEPrint(id){
			window.open("comprasAFEPrint.asp?idAFE=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
		}

		function nuevoPago(id) {
    <%      Set rsBgt = leerBudgetActivosCTC(CTC_idContrato)
		    if (not rsBgt.eof) then     %>
				var puw = new winPopUp('popUpPago', 'comprasCTCPago.asp?idContrato='+id, 530, 590, 'Cargar Pago', 'recargar()');				
    <%      else %>
                alert("No hay definida una partida presupuestaria activa para el contrato.");
    <%      end if %>                    
		}

		function editarPIC(idCTC, idPIC) {			
			var puw = new winPopUp('popUpPago', 'comprasCTCPago.asp?idContrato='+idCTC+'&idPIC='+idPIC, 530, 590, 'Cargar Pago', 'recargar()');
		}

		function abrirCTZ(idCTZ) {
			window.open ("comprasPICPrint.asp?idCotizacionElegida=" + idCTZ, "_blank", "resizable=yes,location=no,menubar=no,statusbar=no",false);
		}
		
		function abrirPedido(id) {
			window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=0", "_blank", "resizable=yes,location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
		}
		
		function anularCTZ(idCotizacion, idPedido, img){
			if (confirm("Esta seguro que desea anular este Pedido Interno?")) {
				img.src = "images/loading_small_green.gif"
				ch.bind("comprasAnularCTZAjax.asp?idCotizacion=" + idCotizacion + "&idPedido=" + idPedido, "anularCTZCallback('" + img.id + "')");
				ch.send();
			}
		}

		function anularCTZCallback(pId){
			var myTable = document.getElementById("TBL_CTC_PAGOS");
			var myImg = document.getElementById(pId);
			myTable.deleteRow(myImg.parentNode.parentNode.rowIndex);
		}

		function abrirObra(id) {
			window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=600,width=900",false);
		}

		function abrirAdjunto(id) {
			window.open("comprasOpenArchivo.asp?idContrato=" + id, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=200,width=300",false);
		}
		
		function MostrarCalendario(p_objID, funcSel){
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
		
		function CerrarCal(cal){
			cal.hide();
		}
		
		function SeleccionarCalEmision(cal, date){
			var str= new String(date);
			document.getElementById("issuedateDiv").innerHTML = str;
		    document.getElementById("issuedate").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			if (cal) cal.hide();
		}
		
		function abrirPartidaCTC(idContrato){
			var puw = new winPopUp("popUpModificarObra","comprasCTCPartidaPopUp.asp?idContrato="+idContrato,"800","400","Modifcar Partida","recargar();");
		}
		function abrirAFEs() {
			window.open('comprasPopUpAFE.asp?idObra=<%=CTC_idObra %>&idPedido=<% =CTC_idPedido %>', '_blank','resizable=yes,location=no,menubar=no,statusbar=no,height=400,width=500,scrollbars=yes',false);
		}	
		function editarResponsable(){
		    document.getElementById("dsResponsable").style.display = "block";
		    document.getElementById("divResponsable").style.display = "none";
		    autocompleteResponsable();
		    var elementImg = document.getElementById("imgResponsable");
		    elementImg.src = "images/save-16.png"
		    elementImg.title = "Guardar responsable";
		    if (isFirefox) {
		        document.getElementById("imgResponsable").setAttribute("onclick", "javascript:guardarResponsable();")
		    } else {
		        elementImg['onclick'] = new Function("javascript:guardarResponsable()");
		    }		    
		}
		function guardarResponsable(){
		    var cdResponsableNew = document.getElementById("cdResponsable").value;
		    if ((cdResponsableNew != "")&&(cdResponsableNew != "<%= CTC_cdResponsable %>")){
		        document.getElementById("imgResponsable").src = "images/loading_small_green.gif";
		        ch.bind("comprasCTC.asp?idContrato=<%=CTC_idContrato%>&cdResponsableNew="+cdResponsableNew+"&cdResponsableOld=<%= CTC_cdResponsable %>&accion=<%=ACCION_CONFIRMAR%>", "guardarResponsable_Callback()");
		        ch.send();
		    }
		    else {
		        document.getElementById("dsResponsable").style.display = "none";
		        document.getElementById("divResponsable").style.display = "block";
		        document.getElementById("divResponsable").innerHTML = "<%= CTC_dsResponsable %>";
		        var elementImg = document.getElementById("imgResponsable");
		        elementImg.src = "images/edit-16.png"
		        elementImg.title = "Editar responsable";
		        if (isFirefox) {
		            elementImg.setAttribute("onclick", "javascript:editarResponsable();")
		        } else {
		            elementImg['onclick'] = new Function("javascript:editarResponsable()");
		        }
		    }
		}
		function guardarResponsable_Callback(){
		    //Se submite la pagina para que la variabla global CdResponsable tome el nuevo valor,debido a que es 
		    //utilizada para validar el permiso de extender fecha de vencimiento del CTC
		    document.getElementById("frmCTC").submit();
		}
		function autocompleteResponsable(){
		    $( "#dsResponsable" ).autocomplete({
		        minLength: 2,
		        source: "comprasStreamElementos.asp?tipo=JQPersonas",
		        focus: function( event, ui ) {
		            $( "#dsResponsable" ).val(ui.item.nombre);
		            return false;
		        },
		        select: function( event, ui ) {
		            $( "#dsResponsable"    ).val (ui.item.nombre);
		            $( "#cdResponsable"  ).val (ui.item.cdusuario );					
		            return false;
		        },
		        change: function( event, ui ) {
		            if (!ui.item){
		                $( "#dsResponsable").val ("");
		                $( "#cdResponsable"  ).val ("");						
		            }
		        }				
		    })
			.data( "autocomplete" )._renderItem = function( ul, item ) {
			    return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
					.appendTo( ul );
			};
		}
	</script>
</head>
<body onload="bodyOnLoad()">	
	<div id="toolbar"></div>
	<br>
	<form id="frmCTC" name="frmCTC">
		<table width="90%" align="center" border="0">
			<tr>
				<td width="70%">&nbsp;</td>
				<td>
					<table width="100%" align="right" cellpadding="2" cellspacing="1" class="reg_Header" border="0">
						<tr>
							<td><% =GF_TRADUCIR("Seleccione Moneda") %>:</td>
							<td>
								<select id="cdMoneda" name="cdMoneda" onChange="cambiarMoneda();">
									<option value="<%=MONEDA_PESO%>" <%if CTC_cdMoneda = MONEDA_PESO then response.write "selected"%> ><% =GF_TRADUCIR("Peso argentino") %></option>
									<option value="<%=MONEDA_DOLAR%>" <%if CTC_cdMoneda = MONEDA_DOLAR then response.write "selected"%> ><% =GF_TRADUCIR("Dolar estadounidense") %></option>
								</select>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<br>
		<table width="90%" align="center" cellpadding="2" cellspacing="1" class="reg_Header" border="0">	
			<tr><td colspan="5"><% call showErrors() %></td></tr>
			<tr>
				<td width="8%" class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Contrato") %></td>
				<td width="20%" class="reg_Header_navdos" colspan="3">&nbsp;<b><% =CTC_cdContrato %></b></td>				
				<td width="8%" class="reg_Header_nav round_border_top_left"><% =GF_TRADUCIR("Total Contrato") %></td>
				<td width="20%" class="reg_Header_navdos">&nbsp;<b><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_TotalImporte, 2) %></b></td>
				<td class="reg_Header_navdos">
				    <%  if (myRol = FIRMA_ROL_GTE_COMPRAS) then %>
				        <a href="javascript:irAjusteCTC()"><img src="images/compras/edit-16x16.png" /></a>
				    <% end if %>
				</td> 
				<td width="44%" rowSpan="10"><span id="ajustes"></span></td>
			</tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Titulo") %></td>
				<td class="reg_Header_navdos" colspan="3">&nbsp;<b><% =CTC_Titulo %>&nbsp;</b></td>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Tipo") %></td>
				<td class="reg_Header_navdos" colspan="2">&nbsp;<b><% =getDsTipoCTC(CTC_tipo) %>&nbsp;</b></td>
			</tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Saldo a Pagar") %></td>
				<td class="reg_Header_navdos" colspan="3"><% =GF_EDIT_DECIMALS(CTC_Total_ImporteSaldo, 2) %></td>				
                <td class="reg_Header_nav"><% =GF_TRADUCIR("Fecha Vto") %></td>
				<td class="reg_Header_navdos">
					<div id="issuedateDiv" class="labelStyle">
					<%  if(CDbl(fechaVto) = 0)then
							Response.Write "Sin fecha definida"
						else
							Response.Write GF_FN2DTE(fechaVto)							
						end if
					%></div>
					<input type="hidden" id="issuedate" name="issuedate" value="<% =fechaVto %>" />
				</td>
				<%  if (puedeExtenderCTC(CTC_idContrato, CTC_cdResponsable)) then %>
					<td class="reg_Header_navdos" align="center"><a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalEmision)"><img id="imgLimite" src="images/DATE.gif"></a></td>
				<% else %>
					<td class="reg_Header_navdos" align="center"></td>
				<% end if %>
			</tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Fondo de Reparo") %></td>
				<td class="reg_Header_navdos" colspan="3">&nbsp;<b><% =CTC_FReparo %>&nbsp;%</b></td>
                <td class="reg_Header_nav"><% =GF_TRADUCIR("Proveedor") %></td>
				<td class="reg_Header_navdos" colspan="2">&nbsp;<b><% =CTC_idProveedor &"-"& CTC_dsProveedor %></b></td>
			</tr>
			<tr>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Responsable") %></td>
				<td class="reg_Header_navdos" colspan="2">&nbsp;
                    <div id="divResponsable"> <b><% =CTC_dsResponsable %></b> </div>
                    <input type="text" id="dsResponsable" name="dsResponsable" style="width:100%;display:none;" value="<%= CTC_dsResponsable %>" /> 
                    <input type="hidden" id="cdResponsable" name="cdResponsable" value="<%= CTC_cdResponsable %>" />
				</td>
                <td class="reg_Header_navdos">
                    <%  if ((myRol = FIRMA_ROL_GTE_COMPRAS)or(myRol = FIRMA_ROL_LEGALES)) then %>
                    <img src="images/edit-16.png" id="imgResponsable" style="cursor:pointer;" title="Editar responsable" onclick="javascript:editarResponsable()"/>
                    <% end if %>
                </td>
                <td class="reg_Header_nav"><% =GF_TRADUCIR("AFEs") %></td>
				<td class="reg_Header_navdos">&nbsp;<b><% =GF_TRADUCIR("Ver y trabajar con los AFE") %></b></td>
				<td class="reg_Header_navdos"><img onclick="javascript:abrirAFEs()" src="images/compras/afe-16x16.png" title="Ver AFE" style="cursor:pointer"></td>
			</tr>
			<tr>			
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Pedido de Precio") %></td>
				<% if (pct_idPedido > 0) then %>
				<td class="reg_Header_navdos" colspan="2" >&nbsp;<b><% =pct_cdPedido & " --> " & Left(pct_tituloPedido, 30) & "..."%></b></td>
				<td class="reg_Header_navdos" align="center"><img onclick="abrirPedido(<% =pct_idPedido %>)" style="cursor:pointer" src="images/compras/PCT-16x16.png" title="Ver Pedido de Cotizacion"></td>				
				<%  else     %>
                <td class="reg_Header_navdos" colspan="2" >&nbsp;<b>-</b></td>
				<td class="reg_Header_navdos" align="center"></td>								
				<% end if 
				if (CTC_tipo = CTC_TIPO_UNITARIO) then
				%>
				<td class="reg_Header_nav"><% =GF_TRADUCIR("Valor Unitario") %></td>
				<td class="reg_Header_navdos"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_valorUnitario, 2) %></td>		
				<td class="reg_Header_navdos">
				    <%  if (myRol = FIRMA_ROL_GTE_COMPRAS) then %>
				        <a href="javascript:irAjusteVUCTC()"><img src="images/compras/edit-16x16.png" /></a>
				    <% end if %>
				</td> 
				<% end if %>		
			</tr>
			
			<% Call drawRowObra() %>
		</table>
		<br>
		<table id="TBL_CTC_PAGOS" class="reg_Header" align="center" width="90%">
			<tr><td colspan="11"><div id="paginacion"></div></td></tr>
			<tr class="reg_Header_nav">
				<td width="3%" align="center" class="round_border_top_left"><% =GF_TRADUCIR("PIC") %></td>				
				<td width="25%"><% =GF_TRADUCIR("Descripci�n") %></td>
				<td width="10%" align="center"><% =GF_TRADUCIR("Partida") %></td>
				<td width="7%" align="center"><% =GF_TRADUCIR("Ord.Pago") %></td>				
				<td width="10%"><% =GF_TRADUCIR("Cumplido") %></td>
				<td width="10%"><% =GF_TRADUCIR("Anticipo") %></td>				
				<td width="10%"><% =GF_TRADUCIR("F. de Reparo") %></td>
				<td width="10%"><% =GF_TRADUCIR("Pago Efectivo") %></td>
				<td width="5%" align="center">%</td>
				<td width="10%"><% =GF_TRADUCIR("Saldo") %></td>
				<td width="2%" align="center">.</td>
				<td width="2%" align="center">.</td>
				<td width="3%" class="round_border_top_right" align="center">.</td>
			</tr>
			<tr class="reg_Header_navdos">
				<td colspan="3"><% =GF_TRADUCIR("Contrato") %>:&nbsp;<%=CTC_cdContrato%></td>
				<td colspan="6">&nbsp;</td>
				<td align="right"><%= simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_TotalImporte, 2)%></td>
				<td colspan="3">&nbsp;</td>
			</tr>
<%			CTC_Total_ImporteObra		= 0
			CTC_Total_ImporteAnticipo	= 0
			CTC_Total_ImporteFReparo	= 0
			CTC_Total_ImportePago		= 0
			CTC_Total_Porc				= 0
			
			CTC_Total_FAC_ImporteObra		= 0
			CTC_Total_FAC_ImporteAnticipo	= 0
			CTC_Total_FAC_ImporteFReparo	= 0
			CTC_Total_FAC_ImportePago		= 0
			CTC_Total_FAC_Porc				= 0
			
			reg = 0			
			While (not rsPagos.eof)
				'Inicializo el flag de PIC en el primer pic de la lista para que se comience sumarizando los datos del PIC antes de imprimirse.				
				facs = 0				
				stringPIC = ""				
				idPIC = CLng(rsPagos("IDPIC"))
				cdObra = rsPagos("CDOBRA")
				idPIC_old = idPIC								
				observaciones = rsPagos("Observaciones")
				'Analiza todos los registros del PIC antes de imprimir para saber si tiene m�s de una minuta.					
				While ((not rsPagos.eof) and (idPIC_old = idPIC))
					facsAD = 0
					CTC_PIC_ImporteObra		= 0
					CTC_PIC_ImporteAnticipo = 0
					CTC_PIC_ImporteFReparo	= 0
					idArea = CLng(rsPagos("IDAREA"))
					idArea_old = idArea
					idDetalle = CLng(rsPagos("IDDetalle"))
					idDetalle_old = idDetalle
					While ((not rsPagos.eof) and (idPIC_old = idPIC) and (idArea = idArea_old) and (idDetalle = idDetalle_old))	
						'Inicializo los valores												
						CTC_ImporteObra		= 0
						CTC_ImporteAnticipo = 0
						CTC_ImporteFReparo	= 0						
						idFAC = CLng(rsPagos("IDFAC"))
						idFAC_old = idFAC
						tipoFAC = rsPagos("TIPOFAC")
						CTC_PjePago			= (CDbl(rsPagos("IMPORTEDOLARESPIC")) / CTC_TotalImporte)*100						
						if (CTC_cdMoneda = MONEDA_PESO) then CTC_PjePago		= (CDbl(rsPagos("IMPORTEPESOSPIC")) / CTC_TotalImporte)*100
						'Recorro todas las facruras asociadas al PIC, si no hay facturas, habr� una sola linea con el ID de Factura NULL y se tomaran los importes del PIC.		
						While ((not rsPagos.eof) and (idPIC_old = idPIC) and (idArea = idArea_old) and (idDetalle = idDetalle_old) and (idFAC = idFAC_old))
							Select case rsPagos("IDARTICULO")																				
								Case ITEM_ANTICIPO_OBRAS_EN_CURSO
									CTC_ImporteAnticipo = CTC_ImporteAnticipo + getImporteMovimiento(CTC_cdMoneda, rsPagos)
								Case ITEM_FONDO_REPARO_ARS
									CTC_ImporteFReparo = CTC_ImporteFReparo + getImporteMovimiento(CTC_cdMoneda, rsPagos)								
								Case ITEM_FONDO_REPARO_USD
									CTC_ImporteFReparo = CTC_ImporteFReparo + getImporteMovimiento(CTC_cdMoneda, rsPagos)					
								Case ITEM_FONDO_REPARO_ARS_IVA	'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano en el PIC.
									CTC_ImporteFReparo = CTC_ImporteFReparo + getImporteMovimiento(CTC_cdMoneda, rsPagos)								
								Case ITEM_FONDO_REPARO_USD_IVA	'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano en el PIC.
									CTC_ImporteFReparo = CTC_ImporteFReparo + getImporteMovimiento(CTC_cdMoneda, rsPagos)
								Case else 'Cualquiera de los tipos de pago de obra
									CTC_ImporteObra = CTC_ImporteObra + getImporteMovimiento(CTC_cdMoneda, rsPagos)										
							End Select
							rsPagos.MoveNext()
							if (not rsPagos.eof) then 
								idPIC = CLng(rsPagos("IDPIC"))
								idArea = CLng(rsPagos("IDAREA"))
								idDetalle = CLng(rsPagos("IDDetalle"))
								idFAC = CLng(rsPagos("IDFAC"))
							end if
						wend						
						if (idFAC_old > 0) then 
							facs = facs + 1		
							facsAD = facsAD + 1
							'Resto los datos de la factura al saldo del PIC.
							CTC_PIC_ImporteObra		= CTC_PIC_ImporteObra - CTC_ImporteObra
							CTC_PIC_ImporteAnticipo = CTC_PIC_ImporteAnticipo - CTC_ImporteAnticipo
							CTC_PIC_ImporteFReparo	= CTC_PIC_ImporteFReparo - CTC_ImporteFReparo						
							Call agregarFactura(stringPIC, idPIC_old, cdObra, idArea_old, idDetalle_old, observaciones, simboloMoneda, idFAC_old, tipoFAC, CTC_PjePago, CTC_ImporteObra, CTC_ImporteAnticipo, CTC_ImporteFReparo, facs, facsAD)
						else																	
							'Los datos leidos corresponden al PIC, tomo los valores para calcular saldos.
							CTC_PIC_ImporteObra		= CTC_ImporteObra
							CTC_PIC_ImporteAnticipo = CTC_ImporteAnticipo
							CTC_PIC_ImporteFReparo	= CTC_ImporteFReparo						
						end if												
					wend
					'Imprimo la ultima linea del AREA-DETALLE (El saldo sin facturar del PIC.									
					'if ((CTC_PIC_ImporteObra <> 0) or (CTC_PIC_ImporteAnticipo <> 0) or (CTC_PIC_ImporteFReparo <> 0)) then
					if (CTC_PIC_ImporteObra+CTC_PIC_ImporteAnticipo+CTC_PIC_ImporteFReparo <> 0) then
						facs = facs + 1
						facsAD = facsAD + 1
						Call agregarFactura(stringPIC, idPIC_old, cdObra, idArea_old, idDetalle_old, observaciones, simboloMoneda, 0, "", CTC_PjePago, CTC_PIC_ImporteObra, CTC_PIC_ImporteAnticipo, CTC_PIC_ImporteFReparo, facs, facsAD)						
					end if					
					'Se editan los codigos de campo que determinan la cantidad de filas a unificar.
					stringPIC = editAreaDetalle(stringPIC, facsAD)
				wend											
				'Imprimo en pantalla las lineas del PIC
				Call ImprimirPIC(stringPIC, facs)
			wend								
%>
			<tr class="reg_Header_nav">
				<td class="round_border_bottom_left" align="right" colspan="4"><% =GF_TRADUCIR("TOTAL CECs") %></td>				
				<td align="right">
					<% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_ImporteObra,2) %>
					<input type="hidden" id="hidSaldo" name="hidSaldo" value="&nbsp;<%=simboloMoneda & " " & GF_EDIT_DECIMALS(Cdbl(CTC_TotalImporte) - Cdbl(CTC_Total_ImporteObra) - Cdbl(CTC_Total_ImporteAnticipo),2) %>">
				</td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_ImporteAnticipo,2) %></td>				
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_ImporteFReparo,2) %></td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_ImportePago,2) %></td>
				<td align="right"><% =GF_EDIT_DECIMALS(CTC_Total_Porc*100, 2) & " %" %></td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(calculoImporteSaldoPIC,2) %></td>
				<td class="round_border_bottom_right" colspan="3">&nbsp;</td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr class="reg_Header_nav">
				<td class="round_border_bottom_left" align="right" colspan="4"><% =GF_TRADUCIR("TOTAL FACTURAS") %></td>				
				<td align="right">
					<% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_FAC_ImporteObra,2) %>
					<input type="hidden" id="Hidden1" name="hidSaldoFAC" value="<%=simboloMoneda & " " & GF_EDIT_DECIMALS(Cdbl(CTC_TotalImporte) - Cdbl(CTC_Total_FAC_ImporteObra) - Cdbl(CTC_Total_ImporteAnticipo),2) %>">
				</td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_FAC_ImporteAnticipo,2) %></td>				
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_FAC_ImporteFReparo,2) %></td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CTC_Total_FAC_ImportePago,2) %></td>
				<td align="right"><% =GF_EDIT_DECIMALS(CTC_Total_FAC_Porc*100, 2) & " %" %></td>
				<td align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(calculoImporteSaldoFAC,2) %></td>
				<td class="round_border_bottom_right" colspan="3">&nbsp;</td>
			</tr>
		</table>
		<input id="idContrato" name="idContrato" type="hidden" value="<%=CTC_idContrato%>">
		<input id="accion" name="accion" type="hidden" value="<%=ACCION_SUBMITIR%>">
	</form>
</body>
</html>