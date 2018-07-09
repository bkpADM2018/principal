<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosREM.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<%

'------------------------------------------------------------------------
Function hayPIC()
	hayPIC = false		
	if (REM_idPIC > 0) then	
		'Hay un PIC como parametro.
		hayPIC = true
	end if		
End Function
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idRemito, index, controlOK, submitPage, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim flagGuardar, rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myRemitoComment
dim rsAlmacenes, closePopUp
dim esREMNuevo, articulos, idPIC
call GP_ConfigurarMomentos()


idRemito = GF_PARAMETROS7("idRemito", 0, 6)
'cdRemito = GF_PARAMETROS7("cdREM", "", 6)
accion = GF_PARAMETROS7("accion","",6)
'REM_cdRemito = PREFIX_REM_X
closePopUp = ""
resp = GF_PARAMETROS7("resp","",6)

flagGuardar = false
controlOK = false
if (isFormSubmit()) then
	'Se controlan los datos.
	if ((accion = ACCION_GRABAR) or (accion = ACCION_CONTROLAR)) then
		controlOK = controlarRemitoAnulacion()
		'Response.Write "r" & controlOK
		'Response.End 
		if ((accion = ACCION_GRABAR) and (controlOK)) then
			'if (idRemito = 0) then flagGuardar = true
			idRemito = grabarFormularioAnulacion(idRemito)
			'Response.End 
			'Se notifica a los responsables internos de la carga de un nuevo Remito.
			closePopUp = "cerrar();"
			'Response.Redirect "almacenAdministrarRem.asp"
		end if
	end if
end if

'Se cargan los datos del Remito para mostrar en pantalla
Call initHeaderREMDB(idRemito)
REM_cdRemito = PREFIX_REM_X

Set rsAlmacenes = obtenerListaAlmacenesUA()
if ((not rsAlmacenes.eof) and (REM_idAlmacen = 0)) then REM_idAlmacen = rsAlmacenes("IDALMACEN")

esREMNuevo = false
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
	var ITEM_AMOUNT_O = "amountS";
	var ITEM_AMOUNT_O_UNIT = "abreviaturaS";	
	var ITEM_AMOUNT_O_TEXT = "amount_textS";
	var ITEM_SALDO = "original";
		

	var ITEM_CD_INTERNO = "cdInterno";
	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var lastArticulos = 0;		
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ms = new Array();
	var myPopUp;

		
	function agregarLineaArticulo() {		
		var obj = undefined;
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCdInterno = rArticulo.insertCell(2);
		var cCantidadO = rArticulo.insertCell(3);	
		var cUnidadO = rArticulo.insertCell(4);			
		var cCantidad = rArticulo.insertCell(5);
		var cUnidad = rArticulo.insertCell(6);		

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


		//CANTIDAD ORIGINAL
		cCantidadO.align = 'center';
		var dCantidadOrig = document.createElement('div');		
		dCantidadOrig.id = ITEM_AMOUNT_O + lastArticulos;				
		cCantidadO.appendChild(dCantidadOrig);
		var dCantidadUnidadOrig = document.createElement('span');
		dCantidadUnidadOrig.id = ITEM_AMOUNT_O_UNIT + lastArticulos;
		cUnidadO.width = '5%';
		cUnidadO.appendChild(dCantidadUnidadOrig);
		
		var iCantidadO = document.createElement('input');
		iCantidadO.type = "hidden";
		iCantidadO.id = ITEM_SALDO + lastArticulos;
		iCantidadO.name = ITEM_SALDO + lastArticulos;
		cCantidadO.appendChild(iCantidadO);	

		
		//CANTIDAD NUEVA
		cCantidad.align = 'center';
		var dCantidad = document.createElement('div');		
		dCantidad.id = ITEM_AMOUNT_TEXT + lastArticulos;				
		cCantidad.appendChild(dCantidad);
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		cUnidad.width = '5%';
		cUnidad.appendChild(dCantidadUnidad);
		
		var iCantidad = document.createElement('input');
		iCantidad.type = "hidden";
		iCantidad.id = ITEM_AMOUNT + lastArticulos;
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		cCantidad.appendChild(iCantidad);	
				
		/*
		var iCantidad = document.createElement('input');												
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		iCantidad.id = ITEM_AMOUNT + lastArticulos;				
		iCantidad.size= 5;
		iCantidad.maxLength= 9;
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");						
		} else {
			iCantidad['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}			
		iCantidad.style.textAlign = "right";
		cCantidad.appendChild(iCantidad);
					
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		cUnidad.width = '5%';
		cUnidad.appendChild(dCantidadUnidad);
		*/
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;		
	}

	function fillArticulo(linea, id, desc, cantidadOriginal, cantidad, unit, cdInterno) {
		document.getElementById(ITEM_DIV + linea).innerHTML = id;
		document.getElementById(ITEM_ID + linea).value = id;
		document.getElementById(ITEM_DESC + linea).innerHTML = desc;
		document.getElementById(ITEM_AMOUNT + linea).value = cantidadOriginal;					
		document.getElementById(ITEM_AMOUNT_TEXT + linea).innerHTML = cantidadOriginal;					
		document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = unit;
		
		
		document.getElementById(ITEM_AMOUNT_O + linea).innerHTML = cantidadOriginal;					
		document.getElementById(ITEM_AMOUNT_O_UNIT + linea).innerHTML = unit;
		document.getElementById(ITEM_SALDO + linea).value = cantidadOriginal;					
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
		//parent.location.reload();
		var refPopUpArt;
		refPopUpArt = startIWin('popupREMAnulacion');
		refPopUpArt.hide(); 
	}
	function bodyOnLoad() {	
		var myMS;
		var tb = new Toolbar('toolbar', 6, 'images/almacenes/');									
		tb.addButton("accept-16x16.png", "Confirmar", "canSubmit('<% =ACCION_GRABAR %>',0)");
		//tb.addButton("accept-16x16.png", "Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");			
		tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		
		tb.draw();
		<%	
		index = 0	
		if (initArticulos()) then				
			while (readNextArticulo())%>			
				agregarLineaArticulo();			
				fillArticulo(<% =index %>, '<% =REM_idArticulo %>', '<% =REM_dsArticulo %>', <% =REM_CantOriginal %>, <% =REM_cantidad %>, '<% =REM_abreviaturaUnidad %>', '<% =REM_cdInterno%>');
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
</script>
</head>
<body onLoad="bodyOnLoad();<%=closePopUp%>;">	
	<div id="toolbar"></div><br>		
	<form id="frmSel" name="frmSel" action="almacenREMAnulacion.asp" method="POST">	
	<table class="reg_Header" align="center" width="90%" border="0" >				
		<tr><td colspan="5"><% call showErrors() %></td></tr>
			<tr>								
				<td align="right" class="numberStyle" colspan="5"><% =GF_TRADUCIR("Id Remito:") %>&nbsp;<% =idRemito %></td>				
			</tr>
			<tr>
				<td class="reg_Header_nav" colspan="6"><% =GF_TRADUCIR("Datos del Remito") %></td>				
			</tr>
			<tr>
				<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Nro Remito")%></td>
				<td align="left">																		
					<div id="cdValeDiv"><% =REM_nroRemito %></div>															
					<input type="hidden" id="nroRemito" name="nroRemito" value="<% =REM_nroRemito %>"/>										
				</td>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Proveedor") %></td>
				<td colspan="2">
					<% =REM_idProveedor & "-" & REM_dsProveedor %>
					<input type="hidden" id="idProveedor" name="idProveedor" value="<% =REM_idProveedor %>"/>					
				</td>		
			</tr>
			<tr>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
				<td>
					<%
					Set rsAlmacenes = obtenerListaAlmacenes(REM_idAlmacen)
					if (not rsAlmacenes.eof) then 
					    response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
					end if
					%>
					<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =REM_idAlmacen %>">
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
						<td align="center"><%=GF_TRADUCIR("Codigo") %></td>
						<td align="center"><%=GF_TRADUCIR("Descripcion") %></td>
						<td align="center"><%=GF_TRADUCIR("Cd. Interno") %></td>
						<td colspan="2" align="center"><%=GF_TRADUCIR("Recibidos") %></td>
						<td colspan="2" align="center"><%=GF_TRADUCIR("A Devolver") %></td>
					</tr>					
				</table>
			</td></tr>
		</table>
		<input type="hidden" id="accion" name="accion" value="">
		<input type="hidden" id="idRemito" name="idRemito" value="<%=idRemito%>">
		<input type="hidden" id="cdREM" name="cdREM" value="<%=PREFIX_REM_X%>">
		<input type="hidden" id="ref" name="ref" value="<%=REM_idPIC%>">
		<input type="hidden" id="cantArticulos" name="cantArticulos" value="0">
		<input type="hidden" name="resp" id="resp" value="MAYBE">		
	</form>
</body>
</html>