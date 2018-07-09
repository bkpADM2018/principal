<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosREM.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<%
'------------------------------------------------------------------------
'Funcion que permite determinar si el Remito puede modificarse
Function REMNuevo(idRemito) 
	REMNuevo = false		
	if (idRemito = 0) then	
		'Es un Remito nuevo
		REMNuevo = true
	end if		
End Function
'------------------------------------------------------------------------
Function hayPIC()
	hayPIC = false		
	if (REM_idPIC > 0) then	
		'Hay un PIC como parametro.
		hayPIC = true
	end if		
End Function
'------------------------------------------------------------------------
Function obtenerPics(id)
	dim strSQL, conn, rs
	strSQL = "Select * from (Select IDPIC, IDARTICULO, SUM(CANTIDAD) CANTIDAD from TBLREMPIC Where IDREMITO = " & id & " group by IDPIC, IDARTICULO) R inner join TBLCTZCABECERA C "
	strSQL = strSQL & "on R.IDPIC = C.IDCOTIZACION "
	'Response.Write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerPics = rs
End Function
'------------------------------------------------------------------------
Function getObservacionesPic(pIdPic)
	Dim strSQL, rtrn
	rtrn = ""
	strSQL = "Select observaciones from tblctzcabecera where idcotizacion = " & pIdPic
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.Eof Then rtrn = rs("observaciones")
	getObservacionesPic = rtrn	
End Function
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idRemito, index, controlOK, submitPage, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myRemitoComment
dim rsAlmacenes, idPic, descArticulo, pAbrev, rsREMAsociados
dim esREMNuevo, articulos, auxPic, cdInterno, rsArt
call GP_ConfigurarMomentos()

idRemito = GF_PARAMETROS7("idRemito", 0, 6)
cdRemito = GF_PARAMETROS7("cdREM", "", 6)
accion = GF_PARAMETROS7("accion","",6)


resp = GF_PARAMETROS7("resp","",6)

controlOK = false
if (isFormSubmit()) then
	'Se controlan los datos.
	if ((accion = ACCION_GRABAR) or (accion = ACCION_CONTROLAR)) then
		controlOK = controlarRemito()
		if ((accion = ACCION_GRABAR) and (controlOK)) then			
			idRemito = grabarFormulario()			
			Response.Redirect "almacenAdministrarRem.asp"
		end if
	end if
end if

'Se cargan los datos del Remito para mostrar en pantalla
Call initHeaderREM(idRemito)

Set rsAlmacenes = obtenerListaAlmacenesUA()
if ((not rsAlmacenes.eof) and (REM_idAlmacen = 0)) then REM_idAlmacen = rsAlmacenes("IDALMACEN")

esREMNuevo = REMNuevo(idRemito)
%>
<html>
<head>
<title>Remito de Cotizacion</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
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
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/uploadManager.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	//Constantes - Nombre de Campo	
	var SUPPLIER_ID = "supplier";	
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";
	var ITEM_DESC = "articuloItem";
	var ITEM_DIV = "itemDiv";
	var ITEM_AMOUNT = "amount";
	var ITEM_AMOUNT_UNIT = "abreviatura";	
	var ITEM_AMOUNT_TEXT = "amount_text";	
	var ITEM_CD_INTERNO = "cdInterno";
	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var lastArticulos = 0;		
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ms = new Array();
	var myPopUp;
	
	function updateLinkArticulo() {		
		var cmb = document.getElementById("idAlmacen");
		for (k in ms) {
			var link = ms[k].MSURL.substring(0,ms[k].MSURL.lastIndexOf("=")+1);			
			link += cmb.options[cmb.selectedIndex].value;
			ms[k].setNewURL(link);
			//Se blanquean los campos de la fila			
			document.getElementById(ITEM_AMOUNT + k).value = "";
			document.getElementById(ITEM_ID + k).value = "";
			document.getElementById(ITEM_DIV + k).innerHTML = "";
			document.getElementById(ITEM_AMOUNT_UNIT + k).innerHTML = "";			
			document.getElementById(ITEM_CD_INTERNO + k).innerHTML = "";
		}		
	}
	
	function aceptarRemito(id) {
		document.getElementById("resp").value = "OK";
		document.getElementById("frmSel").submit();
	}
	
	function abrirConsultaArticulos(){
		myPopUp = new PopUpWindow('Iframe', 'almacenArticulosParaREM.asp?idAlmacen=' + document.getElementById("idAlmacen").value, '640', '500', 'Articulos Requeridos')
	}
		
	function agregarLineaArticulo() {		
		var obj = undefined;
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCdInterno = rArticulo.insertCell(2);
		var cCantidad = rArticulo.insertCell(3);		
		var cUnidad = rArticulo.insertCell(4);		
		var iCodigo = document.createElement('input');
		
		iCodigo.type = "hidden";
		iCodigo.id = ITEM_ID + lastArticulos;
		iCodigo.name = ITEM_ID + lastArticulos;
		iCodigo.size= 7;
		iCodigo.maxLength = 5;				
		cCodigo.appendChild(iCodigo);			
		var dCodigo = document.createElement('div');
		dCodigo.className = "labelStyle";
		dCodigo.id = ITEM_DIV + lastArticulos;		
		cCodigo.appendChild(dCodigo);
		
		var iDescripcion = document.createElement('div');		
		iDescripcion.id = ITEM_DESC + lastArticulos;				
		cDescripcion.appendChild(iDescripcion);	
		
		//CODIGO INTERNO
		cCdInterno.align = 'center';
		var iCdInterno = document.createElement('div');		
		iCdInterno.id = ITEM_CD_INTERNO + lastArticulos;				
		cCdInterno.appendChild(iCdInterno);
			
		
		cCantidad.align = 'right';	
		<%	if (esREMNuevo) then	%>
		var iCantidad = document.createElement('input');												
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		iCantidad.size= 5;
		iCantidad.maxLength= 9;
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");						
		} else {
			iCantidad['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}			
		iCantidad.id = ITEM_AMOUNT + lastArticulos;				
		iCantidad.style.textAlign = "right";
		cCantidad.appendChild(iCantidad);
		<%	else	%>		
		var dCantidadText = document.createElement('span');
		dCantidadText.id = ITEM_AMOUNT_TEXT + lastArticulos;
		cCantidad.appendChild(dCantidadText);				
		<%	end if	%>
					
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		cUnidad.width = '5%';
		cUnidad.appendChild(dCantidadUnidad);
		
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;		
	}

	<% if (not hayPIC()) then	%>
	function seleccionarArticulo(linea, vss) {
		var desc = "";
		if (vss){
			if (typeof(vss) != "boolean") desc = vss.getSelectedItem();
		}			
		if (desc.indexOf('|') != -1) {					
			var arr2 = desc.split('[');	
			var arr = arr2[0].split('|');
			document.getElementById(ITEM_ID + linea).value = arr[0];
			document.getElementById(ITEM_DIV + linea).innerHTML = arr[0];						
			document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = arr2[1].replace(/]/,"");
			document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = arr[2];
			vss.setValue(arr[1]);			
		} else {
			if (desc == "") {
				document.getElementById(ITEM_ID + linea).value = "";
				document.getElementById(ITEM_DIV + linea).innerHTML = "";
				document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = "";				
				document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = "";
			}
		}
	}

	function seleccionarProveedor(ms) {				
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idProveedor").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idProveedor").value = 0;	
			ms.setValue("");
		}		
	}
	<%	end if %>
	
	function fillArticulo(linea, id, desc, cantidad, unit, cdInterno) {
		document.getElementById(ITEM_DIV + linea).innerHTML = id;
		document.getElementById(ITEM_ID + linea).value = id;
		<%	if (esREMNuevo) then	%>
		document.getElementById(ITEM_AMOUNT + linea).value = cantidad;											
		<%	else	%>
		document.getElementById(ITEM_AMOUNT_TEXT + linea).innerHTML = cantidad;				
		<%	end if	%>
		document.getElementById(ITEM_DESC + linea).innerHTML = desc;
		document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = unit;
		document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = cdInterno;
	}

	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function canSubmit(acc, btn) {		
			submitInfo(acc);		
	}
	
	function irRemitos() {
		location.href = "almacenAdministrarRemitos.asp";
	}
	
	function volver() {	
		location.href = "almacenAdministrarREM.asp";
	}
	function cerrar() {	
		var refPopUpArt;
		refPopUpArt = startIWin('popupREM');
		refPopUpArt.hide(); 
	}
	function bodyOnLoad() {	
		var myMS;
		var tb = new Toolbar('toolbar', 6, 'images/almacenes/');									
		<% if (esREMNuevo) or isFormSubmit() then %> 
			tb.addButton("save-16x16.png", "Guardar", "canSubmit('<% =ACCION_GRABAR %>',0)");
			tb.addButton("accept-16x16.png", "Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");			
			tb.addButton("previous-16x16.png", "Volver", "volver()");
		<% else %> 
			tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		<% end if %> 		
		
		tb.draw();
	<%
	index = 0	
	if (initArticulos()) then				
		while (readNextArticulo())%>			
			agregarLineaArticulo();			
			fillArticulo(<% =index %>, '<% =REM_idArticulo %>', '<% =REM_dsArticulo %>', <% =REM_cantidad %>, '<% =REM_abreviaturaUnidad %>', '<% =REM_cdInterno%>');
			<%
			index=index+1
		wend
	end if
	%>	
	
		pngfix();
		resaltarArticulosConErrores();
	}	
	function resaltarArticulosConErrores(){
		//resalta con otro color articulos con errores, que consigue del array arrArticulosConErrores
		var tblArticulos = document.getElementById("tblArticulos");
		<%
		dim iArticulos
		For iArticulos = 0 to ubound(arrArticulosConErrores)%>
	      for (i=0; i< <%=index%>;i++){
			if (document.getElementById('item' + i).value == '<%=arrArticulosConErrores(iArticulos)%>') tblArticulos.rows[i+1].className = 'reg_Header_Error';
		  }
	    <%next%>
	}
	function closeWin() {
		parent.location.reload();
	}
	function printPic(id) {
		window.open("comprasPICPrint.asp?idCotizacionElegida=" + id, "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function abrirPic(id) {
		location.href="comprasPIC.asp?verRemitos=true&idCotizacionElegida=" + id		
	}
	function abrirREM(id) {
		location.href="almacenREM.asp?idRemito=" + id		
	}
	function verDetalle(img, pIndex) {
		if (document.getElementById('detalleArt_' + pIndex).style.display == "none"){
			img.src = "images/almacenes/Menos.gif";
			document.getElementById('detalleArt_' + pIndex).style.display = '';
		}
		else{
			img.src = "images/almacenes/Mas.gif";
			document.getElementById('detalleArt_' + pIndex).style.display = 'none';
		}
	}
	function lightOn(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdosHL reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdosHL";
		}
	}
	function lightOff(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdos reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdos";
		}
	}
</script>
</head>
<body onLoad="bodyOnLoad()">	
	<div id="toolbar"></div><br>
	<form id="frmSel" name="frmSel" action="almacenREMTitulo.asp" method="POST">
	<table class="reg_Header" align="center" width="90%" border="0" >
		<% if (REM_estado = ESTADO_BAJA) then %>
			<tr>
				<td align="center" class="labelStyle reg_header_rejected" colspan="5"><% =GF_TRADUCIR("El siguiente Remito ha sido dado de baja")  %></td>
			</tr>
		<% end if %>			
		<tr><td colspan="5"><% call showErrors() %></td></tr>
		<%	
		if (idRemito > 0) then %>
			<tr>								
				<td align="right" class="numberStyle" colspan="5">
					<% 
					if REM_cdRemito = CODIGO_REM_ANULACION then 
						Response.Write GF_TRADUCIR("Id Anulación Remito:") 
					else
						Response.Write GF_TRADUCIR("Id Remito:") 
					end if	
					%>
					&nbsp;<% =idRemito %>
				</td>
			</tr>
	<% 	end if %>
			<tr>
				<td class="reg_Header_nav" colspan="6">
					<% 
					if REM_cdRemito = CODIGO_REM_ANULACION then 
						Response.Write GF_TRADUCIR("Datos de la Anulación Remito:") 
					else
						Response.Write GF_TRADUCIR("Datos del Remito:") 
					end if	
					%>
				</td>
			</tr>
			<tr>
			<%if (esREMNuevo) then %>
				<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Nro Remito") %></td>
				<td align="left">																			
					<input type="text" id="nroRemito" name="nroRemito" value="<% =REM_nroRemito %>" onKeyPress="return controlDatos(this, event, 'N')"/>										
				</td>
			<%else%>
				<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Nro Remito")%></td>
				<td align="left">																		
					<div id="cdValeDiv"><% =REM_nroRemito %></div>															
					<input type="hidden" id="nroRemito" name="nroRemito" value="<% =REM_nroRemito %>"/>										
				</td>
			<%end if%>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Proveedor") %></td>
				<td colspan="2">
					<% =REM_idProveedor & "-" & REM_dsProveedor %>
					<input type="hidden" id="idProveedor" name="idProveedor" value="<% =REM_idProveedor %>"/>					
				</td>		
			</tr>
			<tr>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
				<td >
					<% if (esREMNuevo) then %>						
						<select id="idAlmacen" name="idAlmacen" onChange="updateLinkArticulo()">
						<%	while (not rsAlmacenes.eof)	%>
							<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = REM_idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
					<%		rsAlmacenes.MoveNext()
						wend 	%>		
						</select>						
					<% else						 
						 Set rsAlmacenes = obtenerListaAlmacenes(REM_idAlmacen)
						 if (not rsAlmacenes.eof) then 
							 response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
						 end if%>
						<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =REM_idAlmacen %>">
					<%end if%>
				</td>	
				<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Fecha") %></td>
				<td>																		
					<div id="issuedateDiv" class="labelStyle"><% =REM_Fecha %></div>
				</td>
				<td></td>				
			</tr>			
			<tr><td class="reg_Header_nav" colspan="5"><% =GF_TRADUCIR("Detalle") %></td></tr>
			<tr><td colspan="6">
				<table class="reg_Header" width="100%" id="tblArticulos">
					<tr class="reg_Header_nav">
						<td align="center"><% =GF_TRADUCIR("Codigo") %></td>
						<td><% =GF_TRADUCIR("Descripcion") %></td>
						<td><% =GF_TRADUCIR("Cd. Interno") %></td>
						<td colspan="2" align="center"><% =GF_TRADUCIR("Cantidad") %></td>
					</tr>
				</table>
			</td></tr>			
			<% if (hayPIC()) then %>			
				<tr><td class="reg_Header_nav" colspan="5"><% =GF_TRADUCIR("Observaciones") %></td></tr>
				<tr><td colspan="5"><%= getObservacionesPic(REM_idPIC) %><td><tr>
			<% end if %>
		</table>
		<table align="center" width="90%" cellpadding="0" cellspacing="0" border="0">
			<tr><td>Cargo: <% =REM_usuario & " - " & getUserDescription(REM_usuario) & " | " & GF_FN2DTE(REM_momento) %></td></tr>
		</table>
		<br>
<% if (not esREMNuevo) then %>
		<table align="center" width="80%" class="reg_Header">
			<tr>
				<td align="center" colspan="5"><b><% =GF_TRADUCIR("PIC asociados al Remito") %></b></td>
			</tr>	
			<tr class="reg_Header_nav">
				<td align="center"><% =GF_TRADUCIR(".") %></td>
				<td align="center"><% =GF_TRADUCIR("Nro. Pic") %></td>
				<td align="center"><% =GF_TRADUCIR("Proveedor") %></td>
				<td align="center"><% =GF_TRADUCIR("Fecha de Entrega") %></td>
				<td align="center"><% =GF_TRADUCIR(".") %></td>
				<td align="center"><% =GF_TRADUCIR(".") %></td>
			</tr>	

<%	'obtengo PICs asociados
	Set rsPic = obtenerPics(idRemito)
	reg=0	
	if (not rsPic.eof) then			
			while (not rsPic.eof)
				reg=reg+1
				idPic = cDbl(rsPic("IDCOTIZACION"))
				auxPic = idPic
			%>							
				<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this, '')" onMouseOut="javascript:lightOff(this, '')">
					<td align="center"><img style="cursor:pointer;" onclick="verDetalle(this,<%=reg%>)" title="<%=GF_TRADUCIR("Detalle Articulos")%>" src="images/Mas.gif"></td>
					<td align="center"><%=idPic%></td>
					<td align="left"><% =rsPic("IDPROVEEDOR") & "-" & getDescripcionProveedor(rsPic("IDPROVEEDOR")) %></td>
					<td align="center"><% =GF_FN2DTE(rsPic("FECHAENTREGA")) %></td>
					<td align="center"><img style="cursor:pointer;" onclick="abrirPic(<%=rsPic("IDCOTIZACION")%>)" title="<%=GF_TRADUCIR("Ver Pic")%>" src="images/almacenes/PIC-16x16.png"></td>
					<td align="center"><img style="cursor:pointer;" onclick="printPic(<%=rsPic("IDCOTIZACION")%>)" title="<%=GF_TRADUCIR("Imprimir Pic")%>" src="images/almacenes/printer-16x16.png"></td>
				</tr>
				<tr>
					<td colspan="2"></td>
					<td colspan="2" align="left">
						<div id="detalleArt_<%=reg%>" style="display:none;">
							<table width="100%">
								<tr class="reg_Header_nav">
									<td align="center"><% =GF_TRADUCIR("ID Articulo") %></td>
									<td align="center"><% =GF_TRADUCIR("Descripción") %></td>
									<td align="center"><% =GF_TRADUCIR("Cd. Interno") %></td>
									<td align="center"><% =GF_TRADUCIR("Cantidad") %></td>
								</tr>
							<%	
							While ((not rsPic.eof) and (idPic = auxPic))%>
								<tr class="reg_Header_navdos">
									<td align="center"><b><% =rsPic("IDARTICULO") %></b></td>
									<% Call getArticuloFull(rsPic("IDARTICULO"), descArticulo, pAbrev) %>
									<td align="left"><% =descArticulo %></td>
									<%
										call executeQueryDb(DBSITE_SQL_INTRA, rsArt, "OPEN", "Select * from TBLARTICULOSDATOS where IDALMACEN=" & REM_idAlmacen & " and  idArticulo=" & rsPic("IDARTICULO"))
										if (not rsArt.eof) then cdInterno = rsArt("CDINTERNO")
									%>
									<td align="left"><% =cdInterno %></td>
									<td align="right"><% =rsPic("CANTIDAD") & " " & pAbrev & " " %></td>
								</tr>
								<%		
									rsPic.MoveNext 
									if (not rsPic.eof) then 
										auxPic = cDbl(rsPic("IDCOTIZACION"))
									else
										auxPic = 0
									end if
							Wend 
							%>
							</table>
						</div>
					</td>
					<td></td>
				</tr>
	<%		wend
		end if
		if (reg = 0) then
	%>				
			<tr><td class="TDNOHAY" colspan="6"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<%	end if %>
	</table><br>
	<%
	strSQL = "Select * from TBLREMCABECERA where NROREMITO = " & REM_nroRemito
	strSQL = strSQL & " and IDPROVEEDOR = " & REM_idProveedor & " and IDALMACEN = " & REM_idAlmacen
	strSQL = strSQL & " and IDREMITO <> " & idRemito
	call executeQueryDb(DBSITE_SQL_INTRA, rsREMAsociados, "OPEN", strSQL)
	%>
	<% if (not rsREMAsociados.eof) then %>
		<table align="center" width="80%" class="reg_Header">
			<tr>
				<td align="center" colspan="6"><b><% =GF_TRADUCIR("Remitos Relacionados") %></b></td>
			</tr>	
			<tr class="reg_Header_nav">
				<td align="center" colspan="2"><% =GF_TRADUCIR("Id") %></td>
				<td align="center"><% =GF_TRADUCIR("Codigo") %></td>
				<td align="center"><% =GF_TRADUCIR("Numero") %></td>
				<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
				<td align="center"><% =GF_TRADUCIR("Relación") %></td>
			</tr>
			<% While (not rsREMAsociados.eof) %>
				<tr class="reg_Header_navdos reg_Header_navdos <% if (rsREMAsociados("ESTADO") = ESTADO_BAJA) then Response.Write "reg_header_rejected" %>" style="cursor:pointer;"  onMouseOver="javascript:lightOn(this, <%=rsREMAsociados("ESTADO")%>)" onMouseOut="javascript:lightOff(this, <%=rsREMAsociados("ESTADO")%>)"  onclick="abrirREM(<% =rsREMAsociados("IDREMITO") %>)">
					<td align="center" width="3%"><img src="images/almacenes/REM-16x16.png"  onclick="abrirREM(<% =rsREMAsociados("IDREMITO") %>)"></td>
					<td align="center"><% =rsREMAsociados("IDREMITO") %></td>
					<td align="center"><% =rsREMAsociados("CDREMITO") %></td>
					<td align="center"><% =rsREMAsociados("NROREMITO") %></td>
					<td align="center"><% =GF_FN2DTE(rsREMAsociados("FECHA")) %></td>
					<td align="left">
						<%
							if (rsREMAsociados("CDREMITO") = CODIGO_REM_ANULACION) then
								Response.Write "Anulación"
							else
								Response.Write "Original"
							end if
						%>
					</td>
				</tr>
				<% rsREMAsociados.MoveNext %>
			<% Wend %>
		</table><br>
	<% end if %>
<% end if %>
		<input type="hidden" id="accion" name="accion" value="">
		<input type="hidden" id="idRemito" name="idRemito" value="<% =idRemito %>">
		<input type="hidden" id="cdREM" name="cdREM" value="<% =cdRemito %>">
		<input type="hidden" id="ref" name="ref" value="<% =REM_idPIC %>">
		<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">
		<input type="hidden" name="resp" id="resp" value="MAYBE">		
	</form>
</body>
</html>