<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'-----------------------------------------------------------------------------------------------
Function obtenerListaVales(idPedido) 
	Dim strSQL, rs, myWhere, conn
	'Ajusto Paginacion
	strSQL = "select * from tblvalescabecera where partidapendiente = " & idPedido	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaVales = rs
End Function
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim idPedido
Dim vales, rsSector
Dim cdSolicitante, dsSolicitante

idPedido = GF_PARAMETROS7("idPedido",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
cdUsuario = ""
totalCorriente = 0

GP_ConfigurarMomentos

Set vales = obtenerListaVales(idPedido)
lineasTotales = vales.RecordCount
initHeaderPMDB(idPedido)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
.numberStyle {
	font-weight: bold;
	font-size: 12px;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/diagram.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var lastArticulos = 0;	
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";
	var ITEM_DESC = "itemDesc";
	var ITEM_DIV = "itemDiv";
	var ITEM_AMOUNT = "amount";
	var ITEM_AMOUNT_UNIT = "abreviatura";	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");			
	var ch = new channel();
	function abrirVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function anularVale(idAlmacen, idVale, cdVale, idPM, img){
		if (confirm("Esta seguro que desea anular el vale?")) {
			img.src = "images/loading_small_green.gif"
			ch.bind("almacenAnularValeAjax.asp?idAlmacen=" + idAlmacen + "&idVale=" + idVale + "&cdVale=" + cdVale + "&idPM=" + idPM, "anularValeCallback('" + img.id + "')");
			ch.send();			
		}		
	}
	function anularValeCallback(pId){
		document.getElementById(pId).src = "images/almacenes/accept-16x16.png";
		document.getElementById(pId).onclick = "";
		document.getElementById(pId).title = "";
		submitInfo();
	}

	function submitInfo(acc) {		
		document.getElementById("frmSel").submit();
	}
	
	function printPM() {
		window.open("almacenValePedidoPrint.asp?idPedido=<% =idPedido %>", "_new", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	function cerrar() {	
		window.close(); 
	}
	function recargar() {		
		window.location.reload();
	}
	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}		
	function ajustarPM() {
		window.open("almacenValesTitulo.asp?TC=0&pmReferencia=<% =idPedido %>&cdVAle=<%=CODIGO_VS_AJUSTE_PEDIDO%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	function bodyOnLoad() {	
		var tb = new Toolbar("toolbar", 6, "images/almacenes/");
		tb.addButton("printer-16x16.png", "Imprimir", "printPM()");
		tb.addButton("AJP-16x16.png", "Ajustar", "ajustarPM()");
		tb.draw();
	<%	if (not vales.eof) then		%>								
			var pgn = new Paginacion("paginacion");							
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "almacenTableroPM.asp?idPedido=<% =idPedido %>");
	<%	end if 
		index = 0
		if (initArticulosDB(idPedido)) then				
			while (readNextArticuloDB())%>			
			myMS = agregarLineaArticulo();
			fillArticulo(myMS, <% =index %>, '<% =pm_idArticulo %>', '<% =pm_dsArticulo %>', <% =pm_cantidad %>, '<% =pm_abreviaturaUnidad %>');					
	<%			index=index+1
			wend
		end if
	%>			
		pngfix();
	}
	
	function agregarLineaArticulo() {		
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCantidad = rArticulo.insertCell(2);		
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
		var iCantidad = document.createElement('input');		
		iCantidad.type = "hidden";
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		iCantidad.size= 5;
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlDatos(this, event, 'N')");
			iCantidad.setAttribute('onblur', "return controlCampo(this, 'N')");				
		} else {
			iCantidad['onkeypress'] = new Function("return controlDatos(this, event, 'N')");
			iCantidad['onblur'] = new Function("return controlCampo(this, 'N')");
		}			
		iCantidad.id = ITEM_AMOUNT + lastArticulos;		
		cCantidad.align = 'right';
		cCantidad.appendChild(iCantidad);
		var ms;		
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		cCantidad.appendChild(dCantidadUnidad);
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;
		return ms;
	}	
	function fillArticulo(pMS, linea, id, desc, cantidad, unit) {
		document.getElementById(ITEM_DIV + linea).innerHTML = id;
		document.getElementById(ITEM_ID + linea).value = id;
		document.getElementById(ITEM_DESC + linea).innerHTML = desc;		
		document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = cantidad + " " + unit;		
		document.getElementById(ITEM_AMOUNT + linea).value = cantidad;		
	}	
</script>
</head>
<body onLoad="bodyOnLoad()">
	<form name="frmSel" id="frmSel">			
	<div id="toolbar"></div><br>
	<table class="reg_Header" align="center" width="90%" border="0" >							
			<tr>								
				<td align="center" class="numberStyle" colspan="4"><% =GF_TRADUCIR("Pedido de Materiales") %></td>				
			</tr>
			<tr>								
				<td align="right" class="numberStyle" colspan="4"><% =GF_TRADUCIR("Nro. Pedido") %>&nbsp;<% =idPedido %></td>				
			</tr>
			<tr>
				<td class="reg_Header_nav" colspan="4"><% =GF_TRADUCIR("Datos del Pedido") %></td>				
			</tr>
			<tr>
				<td class="reg_Header_navdos" ><% =GF_TRADUCIR("Part. Presup") %></td>
				<td colspan="3" >
					<%
					call loadDatosObra(PM_idObra, PM_cdObra, PM_dsObra, 0, "", 0, "", 0, "", "", "", "", "")
					response.write PM_cdObra & " - " & PM_dsObra
					%>
				</td>
			</tr>
			<tr>
				<td class="reg_Header_navdos" ><% =GF_TRADUCIR("Sector") %></td>
				<td colspan="3" >
					<%
					Set rsSector = obtenerSectores(PM_idSector)
					if (not rsSector.eof) then response.write rsSector("IDSECTOR") & " - " & rsSector("DSSECTOR")
					%>
				</td>
			</tr>
			<tr>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Solicitante") %></td>
				<td>
					<% 
					if (PM_idAlmacenDest <> 0) then
						Set rsAlmacenes = obtenerListaAlmacenes(PM_idAlmacenDest) 
						if (not rsAlmacenes.eof) then 
						    response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
						end if
					else
						response.write pm_dsSolicitante
					end if
					%>
				</td>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
				<td>
					<% 	 Set rsAlmacenes = obtenerListaAlmacenes(pm_idAlmacen) 
						 if (not rsAlmacenes.eof) then 
							 response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
						 end if%>
				</td>						
			</tr>
			<tr>				
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha Solicitud") %></td>
				<td align="center">																		
					<div id="issuedateDiv" class="labelStyle"><% =pm_FechaSolicitud %></div>
				</td>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha Requerido") %></td>
				<td align="center">					
					<div id="closingdateDiv" class="labelStyle"><% =pm_FechaRequerido %></div>						
				</td>				
			</tr>
			<tr>
				<td class="reg_Header_nav" colspan="4"><% =GF_TRADUCIR("Comentarios") %></td>
			</tr>			
			<tr>
				<td colspan="4"><% =GF_TRADUCIR(PM_comentario) %></td>
			</tr>			
			<tr>
				<td class="reg_Header_nav" colspan="4"><% =GF_TRADUCIR("Detalle") %></td>
			</tr>
			<tr><td colspan="4">
				<table class="reg_Header" width="100%" id="tblArticulos">
					<tr class="reg_Header_nav">
						<td width="10%"><% =GF_TRADUCIR("Código") %></td>
						<td width="80%"><% =GF_TRADUCIR("Descripción") %></td>
						<td width="10%"><% =GF_TRADUCIR("Cantidad") %></td>
					</tr>					
				</table>
			</td></tr>	
		</table>
	<br>
	
	<table align="center" width="90%" class="reg_Header">
			<tr>
				<td align="center" colspan="5"><b><% =GF_TRADUCIR("Vales asociados al Pedido") %></b></td>				
			</tr>	
			<% 	if (not vales.eof) then %>
			<tr><td colspan="10"><div id="paginacion"></div></td></tr>
		<%	end if 	%>
			<tr class="reg_Header_nav">
				<td align="center" colspan="2"><% =GF_TRADUCIR("Nro.") %></td>				
				<td align="center"><% =GF_TRADUCIR("Tipo") %></td>
				<td align="center"><% =GF_TRADUCIR("Fecha") %></td>
				<td align="center"><% =GF_TRADUCIR("Solicitante") %></td>
				<td align="center">.</td>
				<td align="center">.</td>
			</tr>	
			
<%	reg=0	
	if (not vales.eof) then			
			while ((not vales.eof) and (reg < mostrar))
				reg=reg+1
			%>							
			<tr class="reg_Header_navdos <% if (vales("ESTADO") = ESTADO_BAJA) then Response.Write "reg_header_rejected" %>">
				<td align="center"><%=vales("NRVALE")%></td>
				<td align="center"><img src="images/almacenes/<%=vales("CDVALE")%>-16x16.png"></td>
				<td align="left"><% =getLeyendaCdVale(vales("CDVALE")) & " (" & vales("CDVALE") & ")"%></td>
				<%
					cdSolicitante = vales("cdSolicitante")
					dsSolicitante = getUserDescription(cdSolicitante)
				%>
				<td align="center"><% =GF_FN2DTE(vales("FECHA")) %></td>
				<td align="center"><% =dsSolicitante %></td>	
				<td align="center"><img style="cursor:pointer;" onclick="abrirVale(<%=vales("IDVALE")%>)" title="<%=GF_TRADUCIR("Imprimir Vale")%>" src="images/almacenes/printer-16x16.png"></td>
				<% if (ConfirmaValeAnular(vales("IDVALE"))) then %>
					<td align="center"><img id="IMG_<%=vales("IDVALE")%>" style="cursor:pointer;" onclick="anularVale(<%=vales("IDALMACEN")%>, <%=vales("IDVALE")%>,'<%=vales("CDVALE")%>', <% =idPedido %>, this)" title="<%=GF_TRADUCIR("Anular Vale")%>" src="images/almacenes/vale_reget-16x16.png"></td>
				<%else%>
					<td align="center">.</td>
				<% end if %>
			</tr>
	<%			vales.MoveNext()				
			wend %>
	<%end if
		if (reg = 0) then
	%>				
			<tr><td class="TDNOHAY" colspan="7"><% =GF_TRADUCIR("No se encontraron datos para mostrar") %></td></tr>
	<%	end if %>
		</table>	
		<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">
		<input type="hidden" id="idPedido" name="idPedido"  value="<%=idPedido%>">
</form>
</body>
</html>