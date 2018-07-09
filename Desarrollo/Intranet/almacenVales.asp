<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call controlAccesoAL("")
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idVale, index, esModificable, controlOK, submitPage, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim rsObras, rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myvaleSalidaComment
Dim minCantPro, nrmName, fromTC, aux
Dim rsAlmacenes, devueltos,rsBudget, puedeTransferir, textChecked
Dim title1, title2, title3, mainTitle, lastChar, auxAju, esTransferencia
Dim idPMReferencia, idPMReferenciaHDDN, myJSClose, flagGrabarVale, flagGrabarValeDetalle
Dim myIdAlmacen

flagGrabarValeDetalle = true
myIdAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
'Response.write "ACA" & VS_idAlmacen
Call GP_ConfigurarMomentos()
VS_secBudget = 0
VS_secBudgetArea = 0
myJSClose = ""
VS_comentario = GF_PARAMETROS7("comentario","",6)
idVale = GF_PARAMETROS7("idVale",0,6)
if idVale <> 0 then flagGrabarValeDetalle = false
idPMReferencia = GF_PARAMETROS7("pmReferencia",0,6)
VS_cdVale = GF_PARAMETROS7("cdVale","",6)
fromTC = GF_PARAMETROS7("TC",0,6)
idPMReferenciaHDDN = GF_PARAMETROS7("pmReferenciaHDDN",0,6)
accion = GF_PARAMETROS7("accion","",6)
esTransferencia=false
if ((GF_PARAMETROS7("esTransferencia","",6) <> "") or _
	(VS_cdVale = CODIGO_VS_TRANSFERENCIA) or _
	(VS_cdVale = CODIGO_VS_RECEPCION)) then esTransferencia=true	
if esTransferencia then textChecked = "CHECKED"

controlOK = false
myNewPM = 0
estaPMReferencia = hayPMReferencia(idPMReferencia)
esValeSalidaNuevo = false
if idVale = 0 then 
	esValeSalidaNuevo = true
	if VS_cdVale = CODIGO_PM then 	
		Set rsAlmacenes = obtenerListaAlmacenesSolicitud()
		if (myIdAlmacen = 0) then 
			if (not rsAlmacenes.eof) then myIdAlmacen = rsAlmacenes("IDALMACEN")
		end if
	else	
		Set rsAlmacenes = obtenerListaAlmacenesUA()
		if (myIdAlmacen = 0) then 
			if (not rsAlmacenes.eof) then myIdAlmacen = rsAlmacenes("IDALMACEN")
		end if
	end if
end if
if not (isFormSubmit()) then
	'Response.write "<br>NO SUBMITIO"
	'No submitio la pagina, primera vez que entra
	if estaPMReferencia then
		'Response.write "<br>TIENE PM"
		'Tiene PM, se carga toda la info desde el PM
		call initHeaderPMDB(idPMReferencia)
		call PM2VS
		call initArticulosDB(idPMReferencia)
		VS_CantArticulos = PM_CantArticulos
	else
		'Response.write "<br>NO TIENE PM"
		'No tiene PM, se carga o bien desde el Vale(si exite el IdVale o se carga en blanco)
		call initHeaderVale(idVale)
		if myIdAlmacen > 0 then	VS_idAlmacen = myIdAlmacen 
		call initArticulosVale(idVale)
	end if
else
	'Response.write accion & "<br>"
	'Submitio la pagina
	if estaPMReferencia then
		'Response.write "<br>TIENE PM"
		'Cargo un PM, leer info desde alli
		Call initHeaderPMDB(idPMReferencia)
		Call PM2VS
		VS_secBudget = GF_PARAMETROS7("secBudget", 0, 6)
		VS_FechaSolicitud = GF_PARAMETROS7("issuedate", "", 6)	
		if (VS_FechaSolicitud = "") then VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		
		Call initArticulosDB(idPMReferencia)		
		VS_CantArticulos = PM_CantArticulos
		VS_ArticuloActual = PM_ArticuloActual		
	else
		'Response.write "<br>NO TIENE PM"
		'No cargo un PM, leer desde vale o pagina
		Call initHeaderVale(idVale)	
		Call initArticulosVale(idVale)
	end if	
	'Si se esta cargando una transferencia, la partida y el sector no son datos que deban cargarse
	'ya que el almacen que envia la mercadería no sabe para que la esta enviando.	
	if (esTransferencia) then
		VS_secBudget = 0
		VS_idObra = 0
		VS_idBudgetArea = 0
		VS_idBudgetDetalle = 0
	end if
	'Controlar el Val
	controlOK = controlarVale(idVale)
	
	if ((accion = ACCION_GRABAR) and (controlOK)) then
		'Grabar pedido de materiales si no existe
		if ((not estaPMReferencia) and (VS_cdVale <> CODIGO_VS_ENTRADA)) then
			'El vale a generar debe tener un PM respaldatorio.
			call VS2PM
			if VS_cdVale = CODIGO_PM then
				'Estoy grabando un PM!
				if PM_idAlmacenDest <> 0 then 
					if PM_idAlmacen = PM_idAlmacenDest then 
						PM_idAlmacenDest = 0
					else	
						if (esTransferencia) then 
							idAlmacenAux = PM_idAlmacen
							PM_idAlmacen = PM_idAlmacenDest
							PM_idAlmacenDest = idAlmacenAux
						else	
							PM_idAlmacenDest = 0
						end if	
					end if	
				end if	
			end if	
			'Se graba la cabecera del PM.
			idPMReferencia = grabarHeaderPMInsert()
		end if
		'Grabar o actualizar saldo del detalle del PM		
		VS_ArticuloActual = 0		
		if (VS_cdVale <> CODIGO_VS_ENTRADA and idVale = 0) then
			while ((readNextArticuloVale(idVale)))
					if (estaPMReferencia and VS_saldo > 0) or (not estaPMReferencia and (VS_cantidad > 0 or VS_saldo > 0)) then
						if estaPMReferencia then
							if (VS_cdVale = CODIGO_PM) or (VS_cdVale = CODIGO_VS_SALIDA) or (VS_cdVale = CODIGO_VS_PRESTAMO) or (VS_cdVale = CODIGO_VS_TRANSFERENCIA) then 
								call actualizarPMDetalle(idPMReferencia, VS_idArticulo, VS_saldo)
							end if	
						else
							Call grabarPMDetalle(idPMReferencia, VS_idArticulo, VS_cantidad, VS_saldo)
						end if
						if (VS_saldo > 0) then flagGrabarVale = true  
					end if	
			wend		
			if (idVale <> 0) then flagGrabarVale = true 							
		else
			'Si hay articulos graba o si edita cabecera
			if (readNextArticuloVale(idVale) or (idVale <> 0)) then flagGrabarVale = true
		end if		
		
		'Grabar Vale cabecera y detalle.
		if (VS_cdVale = CODIGO_VS_RECEPCION) then VS_idAlmacen = VS_idAlmacenDest
		VS_ArticuloActual = 0
		if (flagGrabarVale) then
			call grabarHeaderVale(idVale, idPMReferencia)
			call grabarComentarioVale(idVale, VS_comentario)
			if (flagGrabarValeDetalle) then
				while (readNextArticuloVale(idVale))		
						if (VS_cdVale = CODIGO_VS_ENTRADA) then VS_saldo=VS_cantidad
						if (grabarValeDetalle(idVale, idPMReferencia)) then
							call actualizarStock()
						end if
				wend
				Call grabarPreciosVigentesPorArticulo(idVale) 	
				'call ActualizarPrecios(idVale, CODIGO_VS_RECEPCION)
				Call ActualizarPrecios(idVale, VS_cdVale)
			end if
		end if
		'Se setea la accion para ejecutar al cierre.
		if ((fromTC <> 1) and (fromTC <> 2)) then 			
			myJSClose = "location.href= 'almacenAdministrarPedidosMateriales.asp'"			
		elseif (fromTC = 2) then
			myJSClose = "location.href= 'almacenTablerodeControl.asp?idAlmacen=" & myIdAlmacen & "'"
		else
			myJSClose = "cerrar();"
		end if
	end if
end if
if idVale <> 0 then
	title3 = "Cantidad"
else
	select case (VS_cdVale)
		case CODIGO_VS_DEVOLUCION
			title1 = "Pedidos"
			title2 = "Devueltos/<br>Prestados"
			title3 = "Devuelve"
			mainTitle = "Devolucion"
		case CODIGO_VS_PRESTAMO
			title1 = "Pedidos"
			title2 = "Prestados"
			title3 = "Entrego"
			mainTitle = "Prestamo"
		case CODIGO_VS_SALIDA
			title1 = "Pedidos"	
			title2 = "Entregados"
			title3 = "Entrego"
			mainTitle = "Salida"
		case CODIGO_VS_ENTRADA
			title1 = "N/A"	
			title2 = "N/A"
			title3 = "Ingresan"
			mainTitle = "Entrada"
		case CODIGO_VS_TRANSFERENCIA
			title1 = "Pedidos"	
			title2 = "Transferidos"
			title3 = "Entrego"
			mainTitle = "Transferencia"
		case CODIGO_VS_RECEPCION
			title1 = "Pedidos"	
			title2 = "Recib./<br>Transf."
			title3 = "Recibo"		
			mainTitle = "Recepcion"
		case CODIGO_PM
			title1 = "-"	
			title2 = "-"
			title3 = "Pedidos"		
			mainTitle = "Pedido de Materiales"
		case else
			title1 = "Pedidos"	
			title2 = "Entregado"
			title3 = "Saldo"
			mainTitle = "No especificado"
	end select	
end if
'-----------------------------------------------------------------------------------	
Function hayPMReferencia(pIdPMReferencia) 
	dim strSQL, rs, km, kc, tmp, pos	
	hayPMReferencia = false		
	if (idPMReferencia <> 0) then
		strSQL="select * from TBLPMCABECERA where IDPEDIDO=" & pIdPMReferencia
		call executeQueryDb(DBSITE_SQL_INTRA, Rs, "OPEN", strSQL)
		if (not rs.eof) then
			hayPMReferencia = true
		else
			Call setError(PM_REFERENCIA_NO_EXISTE)
			hayPMReferencia = false
		end if
	end if
End Function
%>
<html>
<head>
<title><%=GF_TRADUCIR("Almacen - Vales")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
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
.ui-autocomplete-loading { background: white url('images/loading_small_green.gif') right center no-repeat; }

	.ui-autocomplete-category {
		font-weight: bold;
		padding: .2em .4em;
		margin: auto;
		text-align:center;
		line-height: 1.5;
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
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	
	<% if flagGrabarVale then %>
		if (confirm("Desea realizar la impresion del vale?")) {
			window.open('almacenValePedidoPrint.asp?idVale=<%=idVale%>','ImprimirVale','','');
		}		
	<% end if %>
	<%=myJSClose%>
	
	//Constantes - Nombre de Campo
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";
	var ITEM_DESC = "articuloItem";
	var ITEM_DIV = "itemDiv";
	var ITEM_AMOUNT = "amount";
	var ITEM_SALDO = "saldo";
	var ITEM_CUMPLIDO = "cumplido";
	var ITEM_AMOUNT_UNIT = "abreviatura";	
	var ITEM_CUMPLIDO_TEXT = "cumplidoText";
	var ITEM_SALDO_TEXT = "saldoText";
	var ITEM_CD_INTERNO = "cdInterno";
	var ITEM_STOCK = "existenciaText";
	var ITEM_STOCK_TEXT = "existencia";
	var ITEM_DEVUELTOS = "devueltos";
	
	var myAutoCompletesIndexs = {};
	var lastCategory = "";
	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var lastArticulos = 0;		
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ch = new channel();		
	var ms = new Array();
	
	var flagSaving = false;
		
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
	
	function agregarLineaArticulo() {		
		var obj = undefined;
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var index;		
		index = 2;
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCdInterno = rArticulo.insertCell(2);
		var cStock = rArticulo.insertCell(3);
		var cCantidad = rArticulo.insertCell(4);
		var cCumplido = rArticulo.insertCell(5);		
		var cSaldo = rArticulo.insertCell(6);	
		var iCodigo = document.createElement('input');
		
		//CODIGO
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
		
		//TEXTO
		var iText = document.createElement('input');		
		iText.type="text";
		iText.id = ITEM_DESC + lastArticulos + "_text" ;
		iText.name = ITEM_DESC + lastArticulos + "_text";
		iText.size = 50;
		
		//DESCRIPCION
		var iDescripcion = document.createElement('div');		
		iDescripcion.id = ITEM_DESC + lastArticulos;				
		iDescripcion.appendChild(iText);
		
		cDescripcion.appendChild(iDescripcion);
		
		
		//CODIGO INTERNO
		cCdInterno.align = 'center';
		var iCdInterno = document.createElement('div');		
		iCdInterno.id = ITEM_CD_INTERNO + lastArticulos;				
		cCdInterno.appendChild(iCdInterno);	

		//STOCK
		cStock.align = 'center';
		var iStock = document.createElement('span');		
		iStock.id = ITEM_STOCK + lastArticulos;				
		cStock.appendChild(iStock);	
		var iTxtStock = document.createElement('input');
		iTxtStock.type = "hidden";
		iTxtStock.name = ITEM_STOCK_TEXT + lastArticulos;
		iTxtStock.id = ITEM_STOCK_TEXT + lastArticulos;
		cStock.appendChild(iTxtStock);	
		
		//CANTIDAD	
		cCantidad.align = 'right';
		var iCantidad = document.createElement('input');	
		var dCantidadUnidad = document.createElement('span');
		<% if estaPMReferencia then %>	
			iCantidad.type = "hidden";
			dCantidadUnidad.style.display = 'none';
		<% end if %>	
		iCantidad.name = ITEM_AMOUNT + lastArticulos;
		iCantidad.size= 4;
		iCantidad.align = 'center';
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");			
		} else {
			iCantidad['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}
		iCantidad.id = ITEM_AMOUNT + lastArticulos;		
		cCantidad.appendChild(iCantidad);
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + lastArticulos;
		dCantidadUnidad.style.textAlign = "right";
		iCantidad.style.textAlign = "right";
		cCantidad.appendChild(dCantidadUnidad);		
		
		//CUMPLIDO
		cCumplido.align = 'right';
		var iCumplido = document.createElement('input');
		var dCumplido = document.createElement('span');
		iCumplido.type = "hidden";
		iCumplido.name = ITEM_CUMPLIDO + lastArticulos;
		iCumplido.size = 4;
		if (isFirefox) {
			iCumplido.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");			
		} else {
			iCumplido['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}
		iCumplido.id = ITEM_CUMPLIDO + lastArticulos;
		<% if VS_cdVale = CODIGO_VS_ENTRADA or VS_cdVale = CODIGO_PM then %>
			cCumplido.style.display = 'none';
			dCumplido.style.display = 'none';
		<% end if %>
		cCumplido.appendChild(iCumplido);	
		var iDevueltos = document.createElement('input');
		iDevueltos.type = "hidden";
		iDevueltos.name = ITEM_DEVUELTOS + lastArticulos;	
		iDevueltos.id = ITEM_DEVUELTOS + lastArticulos;		
		cCumplido.appendChild(iDevueltos);			
		//CUMPLIDO TEXT	
		dCumplido.id = ITEM_CUMPLIDO_TEXT + lastArticulos;	
		iCumplido.style.textAlign = "right";
		cCumplido.appendChild(dCumplido);
		
		//SALDO
		cSaldo.align = 'right';
		var iSaldo = document.createElement('input');	
		var dSaldo = document.createElement('span');	
		iSaldo.name = ITEM_SALDO + lastArticulos;
		iSaldo.size= 4;
		if (isFirefox) {
			iSaldo.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");					
		} else {
			iSaldo['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}		
		iSaldo.id = ITEM_SALDO + lastArticulos;		
		cSaldo.appendChild(iSaldo);
		<%if VS_cdVale = CODIGO_VS_ENTRADA or VS_cdVale = CODIGO_PM then %>
			iSaldo.type = "hidden";
			dSaldo.style.display = 'none';
		<% end if %>

		//SALDO TEXT
		dSaldo.id = ITEM_SALDO_TEXT + lastArticulos;
		iSaldo.style.textAlign = "right";
		cSaldo.appendChild(dSaldo);
		<% if idvale=0 then %>
			//Las funciones internas del autocomplete como focus, select o changes, no se crean en el momento de creacion
			//del autocomplete, sino que se ejecutan en esos momentos (al hacer foco, seleccionar item, cambiar valor).
			//Por lo tanto no podemos utilizar la variable 'lastArticulos' para identificar el indice del autocomplete
			//porque tomaria el valor de la variable al momento ejecutar la accion en lugar del valor al momento de la creacion
			//por tal motivo guardamos un objeto donde tenemos como key el id del autocomplete y como valor el indice del mismo
			//siendo el id del autocomplete accesible mediente this.id podemos identificar dentro de este objeto el indice que
			//necesitamos.
			myAutoCompletesIndexs[ITEM_DESC + lastArticulos + "_text"] = lastArticulos									
		    var link = "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen=" + document.getElementById("idAlmacen").value;		
			$( "#"+ITEM_DESC + lastArticulos + "_text" ).autocomplete({
				minLength: 2,
				//El source se setea al seleccionar un almacen
				source: link,
				focus: function( event, ui ) {
					$( "#"+ITEM_DESC + myAutoCompletesIndexs[this.id] + "_text" ).val(ui.item.dsarticulo);
					return false;
				},
				select: function( event, ui ) {
					var myIndex = myAutoCompletesIndexs[this.id];
					$( "#"+ITEM_DESC + myIndex + "_text").val (ui.item.dsarticulo);
					$( "#"+ITEM_ID + myIndex).val (ui.item.idarticulo);
					$( "#"+ITEM_DIV + myIndex).html (ui.item.idarticulo);
					$( "#"+ITEM_CD_INTERNO + myIndex).html(ui.item.cdinterno);
					$( "#"+ITEM_STOCK + myIndex).html(ui.item.stock);
					$( "#"+ITEM_STOCK_TEXT + myIndex).val(ui.item.stock);
					$( "#"+ITEM_AMOUNT_UNIT + myIndex).html("&nbsp;"+ui.item.abreviatura);
					$( "#"+ITEM_SALDO_TEXT + myIndex).html("&nbsp;"+ui.item.abreviatura);
					return false;
				},
				change: function( event, ui ) {
					if (!ui.item)
					{
						lastCategory = "";
						var myIndex = myAutoCompletesIndexs[this.id];
						$( "#"+ITEM_DESC + myIndex + "_text").val("");
						$( "#"+ITEM_ID + myIndex).val ("");
						$( "#"+ITEM_DIV + myIndex).html ("");
						$( "#"+ITEM_CD_INTERNO + myIndex).html("");
						$( "#"+ITEM_STOCK + myIndex).html("");
						$( "#"+ITEM_STOCK_TEXT + myIndex).val("");
						$( "#"+ITEM_AMOUNT_UNIT + myIndex).html("");
						$( "#"+ITEM_SALDO_TEXT + myIndex).html("");
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				if (item.stock == null) {
					item.stock = 0;
				}
				if (item.cdinterno == null) {
					item.cdinterno = 0;
				}
				
				li_Item = $( "<li></li>" )
							.data( "item.autocomplete", item )
							.append( "<a><font style='font-size:10;'>" + item.idarticulo + " - " + item.dsarticulo + " - "+item.stock+ " ["+item.abreviatura+"]</font></a>" )
							.appendTo( ul );
							
				if (lastCategory != item.idcategoria) {
					lastCategory = item.idcategoria;
					return $(ul)
						.append( "<li class='ui-autocomplete-category'>" + item.dscategoria + "</li>" ).append(
							li_Item
						);
				} else {
					return li_Item;
				}
			};
			ms.push($( "#"+ITEM_DESC + lastArticulos + "_text" ))
		<% end if %>
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;
		<% if idvale<>0 then %>
			iCantidad.type = "hidden";
			cStock.style.display = 'none';	
			cCumplido.style.display = 'none';
			cSaldo.style.display = 'none';
		<% end if %>
		return obj;
	}

	function updateLinkArticulo() {		
		var cmb = document.getElementById("idAlmacenCmb");
		document.getElementById("idAlmacen").value = cmb.options[cmb.selectedIndex].value;
		var link = "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen=" + cmb.options[cmb.selectedIndex].value;				
		for (k in ms) {
		    //Se cambia dinamicamente la propiedad source del autocomplete.
		    ms[k].autocomplete( "option", "source", link);		    		    
			//Se blanquean los campos de la fila			
			document.getElementById(ITEM_AMOUNT + k).value = "";
			document.getElementById(ITEM_ID + k).value = "";
			document.getElementById(ITEM_DESC + k + "_text").value = "";			
			document.getElementById(ITEM_DIV + k).innerHTML = "";
			document.getElementById(ITEM_AMOUNT_UNIT + k).innerHTML = "";
			document.getElementById(ITEM_SALDO_TEXT + k).innerHTML = "";
			document.getElementById(ITEM_CD_INTERNO + k).innerHTML = "";
			document.getElementById(ITEM_STOCK_TEXT + k).value = "";
			document.getElementById(ITEM_STOCK + k).innerHTML = "";
		}		
	}
	
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
			<% if VS_cdVale <> CODIGO_VS_ENTRADA then %>	
				document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = arr2[1].replace(/]/,"");
			<% end if								%>	
			document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = arr2[1].replace(/]/,"");
			document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = arr[2];
			document.getElementById(ITEM_STOCK_TEXT + linea).value = arr[3];
			document.getElementById(ITEM_STOCK + linea).innerHTML = arr[3];
			vss.setValue(arr[1]);			
		} else {
			if (desc == "") {
				document.getElementById(ITEM_ID + linea).value = "";
				document.getElementById(ITEM_DIV + linea).innerHTML = "";
				document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = "";
				document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = "";
				document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = "";
				document.getElementById(ITEM_STOCK_TEXT + linea).value = "";
				document.getElementById(ITEM_STOCK + linea).innerHTML = "";
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
	
	function fillArticulo(vss, linea, id, desc, cantidad, cumplido, saldo, unit, cdInterno, stock, devueltos) {
		<% if idvale <> 0 then %>
			document.getElementById(ITEM_DIV + linea).innerHTML = id;
			document.getElementById(ITEM_ID + linea).value = id;
			document.getElementById(ITEM_DESC + linea).innerHTML = desc;
			document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = cdInterno;
			document.getElementById(ITEM_AMOUNT + linea).value = cantidad;
			document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = cantidad + " " + unit;
		<% else %>
		
			<%		
			if (esModificable) then 				%>
						document.getElementById(ITEM_DESC + linea + "_text").value = desc;
						//vss.setValue(id + "-" + desc + "[" + unit + "]");
						//seleccionarArticulo(linea, vss);		
			<%		else				 				%>		
						document.getElementById(ITEM_DIV + linea).innerHTML = id;
						document.getElementById(ITEM_ID + linea).value = id;
						<%if not estaPMReferencia then%>
							document.getElementById(ITEM_DESC + linea + "_text").value = desc;		
						<% else %>				
							document.getElementById(ITEM_DESC + linea).innerHTML = desc;		
						<% end if %>	

						if (document.getElementById(ITEM_AMOUNT + linea).type != "text"){
							document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = cantidad + " " + unit;
						}
						else{
							document.getElementById(ITEM_AMOUNT_UNIT + linea).innerHTML = "&nbsp;" + unit;
						}						
						if (document.getElementById(ITEM_SALDO + linea).type != "text"){
							document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = saldo + " " + unit;
						}	
						else{							
							document.getElementById(ITEM_SALDO_TEXT + linea).innerHTML = "&nbsp;" + unit;
						}

						if (document.getElementById(ITEM_CUMPLIDO + linea).type != "text"){
							//Para la pantalla se arma el string
						<%	if ((VS_cdVale=CODIGO_VS_DEVOLUCION) or (VS_cdVale=CODIGO_VS_RECEPCION)) then	 %>							
							document.getElementById(ITEM_CUMPLIDO_TEXT + linea).innerHTML = devueltos + "/" + cumplido + " " + unit;							
						<%	else	%>
							document.getElementById(ITEM_CUMPLIDO_TEXT + linea).innerHTML = cumplido + " " + unit;
						<%	end if	%>
						}	
						else{
							document.getElementById(ITEM_CUMPLIDO_TEXT + linea).innerHTML = "&nbsp;" + unit;
						}
			<%		end if				 			%>
					document.getElementById(ITEM_AMOUNT + linea).value = cantidad;
					document.getElementById(ITEM_CUMPLIDO + linea).value = cumplido;					
					document.getElementById(ITEM_DEVUELTOS + linea).value = devueltos;
					document.getElementById(ITEM_SALDO + linea).value = saldo;		
					document.getElementById(ITEM_CD_INTERNO + linea).innerHTML = cdInterno;				
					document.getElementById(ITEM_STOCK_TEXT + linea).value = stock;				
					document.getElementById(ITEM_STOCK + linea).innerHTML = stock;
		<% end if %>
	}

	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function guardar(acc, btn) {
	    if (!flagSaving) {
	        flagSaving = true;
	        canSubmit(acc, btn);	        
	    }
	}
	
	function canSubmit(acc, btn) {		
			submitInfo(acc);		
	}
	
	function irvaleSalidas() {
		location.href = "almacenAdministrarvaleSalidas.asp";
	}
	function cerrar() {	
		var refPopUpArt;
		refPopUpArt = startIWin('popupArt');
		refPopUpArt.hide();
	}
	
	function volverAS() {	
	<%	if isFormSubmit() then%>	
		location.href = "almacenValeSalida.asp?cdVale=VMS";
		<%else%>
		window.close();
	<%end if%>
	}
	function volver() {	
		window.history.back(); 
	}

	function irHome() {
		location.href = "almacenIndex.asp";
	}	
	function irTDC() {
		location.href = "almacenTableroDeControl.asp";
	}	
	function irAPM() {
		location.href = "almacenAdministrarPedidosMateriales.asp";
	}	
	function bodyOnLoad() {			
		var myMS;
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		<% if fromTC <> 1 then %> 
			tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% end if %> 
		idBtnGuardar = tb.addButtonSAVE("Guardar", "guardar('<% =ACCION_GRABAR %>',0)");
		idBtnControl = tb.addButtonCONFIRM("Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");	
		<%if not estaPMReferencia and idVale=0 then%>
			var msSolicitante = new MagicSearch("", "divSolicitante", 30, 4, "comprasStreamElementos.asp?tipo=personas");
			msSolicitante.setToken(";");
			msSolicitante.onBlur = seleccionarSolicitante;
			msSolicitante.setValue('<% =vs_dsSolicitante %>');
		<%end if%>
		<%if fromTC <> 1 then%> 
			tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");	
			<% if VS_cdVale = CODIGO_PM  then %> 
				tb.addButton("PM_folder-16x16.png", "Adm. PM", "irAPM()");	
			<% end if %> 
		<% else %> 
			tb.addButton("close-16x16.png", "Cerrar", "cerrar()");
		<% end if %> 
		tb.draw();
		<% if (((not estaPMReferencia)) and (VS_cdVale <> CODIGO_VS_ENTRADA))then %>
			SeleccionarCalLimite(undefined, '<% = VS_FechaRequerido %>');
		<%end if
		index = 0
		if ((estaPMReferencia) and (idPMReferenciaHDDN<>idPMReferencia)) then
			if VS_cdVale = CODIGO_VS_RECEPCION then PM_idAlmacen = PM_idAlmacenDest
			while (readNextArticuloDB())
					if VS_cdVale = CODIGO_VS_RECEPCION then PM_idAlmacenDest = PM_idAlmacen
					PM2VS_DET
					
					auxAju = getTotalArticuloxVale(idPMReferencia, vs_idarticulo, CODIGO_VS_AJUSTE_PEDIDO)
					if CLng(auxAju) <> 0 then VS_cantidad = cdbl(VS_cantidad) - cdbl(auxAju)
					if (VS_cdVale=CODIGO_VS_DEVOLUCION) then
						devueltos = getCantidadDevuelta(idPMReferencia, VS_idArticulo)
						VS_cumplido = clng(VS_cantidad) - clng(VS_saldo)
						VS_saldo = VS_cumplido - clng(devueltos)						
					elseif (VS_cdVale=CODIGO_VS_RECEPCION) then
						devueltos = getCantidadRecibida(idPMReferencia, VS_idArticulo)
						VS_cumplido = clng(VS_cantidad) - clng(VS_saldo)
						VS_saldo = VS_cumplido - clng(devueltos)
					else
						VS_cumplido = clng(VS_cantidad) - clng(VS_saldo)						
					end if									
					%>
						myMS = agregarLineaArticulo();
						fillArticulo(myMS, <% =index %>, '<% =vs_idArticulo %>', '<% =vs_dsArticulo %>', <% =VS_cantidad %>, '<% =VS_cumplido %>', <% =VS_saldo %>, '<% =vs_abreviaturaUnidad %>', '<% =VS_cdInterno%>', '<% =VS_existencia%>', '<% =devueltos %>');
					<%
				index=index+1
			wend
		elseif (idPMReferenciaHDDN<>idPMReferencia) then
			while (readNextArticuloValeParams())%>
					myMS = agregarLineaArticulo();
					fillArticulo(myMS, <% =index %>, '', '', 0, 0, 0, '', '', 0);					
					<%
				index=index+1
			wend
		else
			linea=0
			while (readNextArticuloVale(idVale))						
					devueltos = GF_PARAMETROS7("devueltos" & linea, 0, 6) 
					linea = linea+1
					%>
					myMS = agregarLineaArticulo();					
					fillArticulo(myMS, <% =index %>, '<% =VS_idArticulo %>', '<% =VS_dsArticulo %>', <% =VS_cantidad %>, <% =VS_cumplido %>, <% =VS_saldo %>, '<% =vs_abreviaturaUnidad %>', '<% =VS_cdInterno%>', '<%=VS_existencia%>', '<% =devueltos %>');					
					<%
				index=index+1
			wend
		end if
		'SI NO ESTA INDICADO EL ID DE PEDIDO DE MATERIALES DE REFERENCIA, el Minimo se muestran 5 lineas de articulos a completar
		if not (estaPMReferencia) then
			while ((index < 5))%>
				agregarLineaArticulo();
			<%index=index+1
			wend
		end if
		if ((esTransferencia) or (VS_cdVale = CODIGO_VS_ENTRADA)) then %>
			ShowHide('almTransferenciaCHK','almTransferencia');
		<% else	%>			
			actualizarBudgets(<%=VS_idObra%>,<%=VS_idBudgetArea%>,<%=VS_idBudgetDetalle%>);			
		<%
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
function actualizarBudgetsCallback(idObra){
	document.getElementById("secBudgetDiv").innerHTML = ch.response(); 	
	var tr = document.getElementById("trSectores");
	if (idObra == 0) {
		//No hay obra elegida, muestro para que elija sector.		
		tr.style.visibility="visible";
		tr.style.position="relative";
	} else {
		//si eligio obra, no debe elegir sector
		tr.style.visibility="hidden";
		tr.style.position="absolute";
		document.getElementById("idSector").selectedIndex=0;
	}
}

function readBudgetArea() {
	document.getElementById('idBudgetArea').value=$("#idBudgetDetalle option:selected").attr("alt");
}

function actualizarBudgets(idObra, idBudgetArea, idBudgetDetalle){
	var myReadOnly = 1;
	<% if not estaPMReferencia then %>	
		myReadOnly = 0;
	<% end if %>
	ch.bind("almacenObtenerBudget.asp?idObra=" + idObra + "&idBudgetArea=" + idBudgetArea + "&idBudgetDetalle=" + idBudgetDetalle + "&readOnly=" + myReadOnly + "&accion=<%=ACCION_PROCESAR%>", "actualizarBudgetsCallback(" + idObra + ")");
	ch.send();	
}
function keyPressed(e) {
	key=(document.all) ? e.keyCode : e.which;
	if(key==13) return false;
}
function ShowHide(pChkId, pId){
	var chk = document.getElementById(pChkId);
	if (chk != null) {
		var div = document.getElementById(pId);
		if (chk.checked){
			div.style.display = "block";
			document.getElementById("esTransferencia").value = 1;
			document.getElementById("partidaRow").style.visibility = "hidden";
			document.getElementById("partidaRow").style.position  = "absolute";
			document.getElementById("trSectores").style.visibility = "hidden";
			document.getElementById("trSectores").style.position  = "absolute";
		} else {
			div.style.display = "none";
			document.getElementById("esTransferencia").value = "";
			document.getElementById("partidaRow").style.visibility = "visible";
			document.getElementById("partidaRow").style.position  = "relative";
			document.getElementById("trSectores").style.visibility = "visible";
			document.getElementById("trSectores").style.position  = "relative";
		}
	}	
}
</script>
</head>

<script>
</script>
</head>
<body onLoad="bodyOnLoad()" onkeypress="return keyPressed(event)">	
<div id="toolbar"></div>
<br>		
<%
	if fromTC = 1 then
		submitPage = "almacenVales.asp"
	else
		submitPage = "almacenValesTitulo.asp"	
	end if	
%>
<form id="frmSel" name="frmSel" action="<% =submitPage %>" method="POST">	
	<table class="reg_Header" align="center" width="95%" border="0">				
		<tr>
			<td colspan="8">
				<%call showErrors()%>
			</td>
		</tr>
		<% 
			if (VS_cdVale <> CODIGO_PM) then 
			'No es un PM => Se muestra titulo de Tipo de vale y PM de referencia. 
			%> 
				<tr>
					<td class="reg_Header_nav" colspan="1" align="center"><font class="big"><% =ucase(VS_cdVale) %></font></td>
					<td align="right" class="reg_Header_nav" colspan="7"><% =GF_TRADUCIR("Nº Pedido de Materiales de Referencia") %>&nbsp;
					<% if (estaPMReferencia) then
						 Response.write idPMReferencia %>
						<input type="hidden" id="pmReferencia" name="pmReferencia" value="<% =idPMReferencia %>" onchange="submitInfo('<%=ACCION_SUBMITIR%>');" size=5></td>
					<% else %>
						<input type="text" id="pmReferencia" name="pmReferencia" value="<% =idPMReferencia %>" onchange="submitInfo('<%=ACCION_SUBMITIR%>');" size=5></td>
					<% end if %>	
					
					<input id="pmReferenciaHDDN" type="hidden" name="pmReferenciaHDDN" value="<% =idPMReferencia %>" onchange="submitInfo('<%=ACCION_SUBMITIR%>');"></td>
				</tr>
			<%
			end if			
		if ((VS_cdVale <> CODIGO_VS_ENTRADA) and _
		    (VS_cdVale <> CODIGO_VS_TRANSFERENCIA) and _
		    (VS_cdVale <> CODIGO_VS_RECEPCION)) then 
		%>
		<tr id="partidaRow">
			<td class="reg_Header_navdos"><%= GF_TRADUCIR("Part. Pres.") %></td>
			<td colspan="6">
				<% 	if not estaPMReferencia then 
						Set rsObras = obtenerListaObras("", "", "","",OBRA_ACTIVA)
				%>
					<select id="idObra" name="idObra" onchange="actualizarBudgets(this.value,0,0)">
						<option value="0">- <% =GF_TRADUCIR("Seleccione") %>
					<%	while (not rsObras.eof)	%>
							<option value="<% =rsObras("IDOBRA") %>" <% if (rsObras("IDOBRA") = vs_idObra) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %></option>
					<%		rsObras.MoveNext()
						wend 	%>
					</select>
				<% else													
					Set rsObra = obtenerListaObras(vs_idObra, "", "","","") 
					if (not rsObra.eof) then 
						response.write rsObra("CDOBRA") & " - " & rsObra("DSOBRA")
					end if	
					%>
					<input type="hidden" name="idObra" id="idObra" value="<% =vs_idObra %>">					
				<%end if%>
				&nbsp;&nbsp;&nbsp;<span id="secBudgetDiv"></span>
			</td>
		</tr>		
		<tr id="trSectores" style="visibility:visible;position:relative">
			<td class="reg_Header_navdos"><%= GF_TRADUCIR("Sector") %></td>
			<td colspan="6">
				<% 	if not estaPMReferencia then 
						Set rsSectores = obtenerSectores("")
				%>
					<select id="idSector" name="idSector">
						<option value="0">- <% =GF_TRADUCIR("Seleccione") %>
					<%	while (not rsSectores.eof)	%>
							<option value="<% =rsSectores("IDSECTOR") %>" <% if (rsSectores("IDSECTOR") = VS_idSector) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsSectores("DSSECTOR")) %></option>
					<%		rsSectores.MoveNext()
						wend 	%>
					</select>
				<% else		
					Set rsSectores = obtenerSectores(VS_idSector)																
					if (not rsSectores.eof) then 
						response.write rsSectores("DSSECTOR")
					end if	
					%>
					<input type="hidden" name="idSector" id="idSector" value="<% =VS_idSector %>">					
				<%end if%>				
			</td>
		</tr>
		<% end if%>
		<tr>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Solicitante") %></td>
			<td colspan="1" width="25%">
				<% if not estaPMReferencia and idVale=0 then %>
					<div id="divSolicitante"></div>																		
				<% else
					response.write VS_dsSolicitante
				end if %>															
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =VS_cdSolicitante %>"/>
			</td>

			<td class="reg_Header_navdos" width="15%">
			<%
				'Entrego material de un pedido ya echo
				Response.write GF_TRADUCIR("Fecha del Vale") 				
			%>
			</td>
			<td align="center" width="10%">											
				<div id="issuedateDiv"><% =VS_FechaSolicitud %></div>															
				<input type="hidden" id="issuedate" name="issuedate" value="<% =VS_FechaSolicitud %>"/>														
			</td>						
			<% if (((not estaPMReferencia)) and (VS_cdVale <> CODIGO_VS_ENTRADA))then %>
				<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Fecha Requerido") %></td>
			<%	else	%>
				<td></td>
			<%  end if %>			
			<td width="5%">
			<% if (((not estaPMReferencia and idVale=0)) and (VS_cdVale <> CODIGO_VS_ENTRADA))then %>
				<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/DATE.gif"></a>
			<% end if %>
			</td>
			<td align="center" width="10%">					
				<% if (((not estaPMReferencia)) and (VS_cdVale <> CODIGO_VS_ENTRADA))then %>
					<div id="closingdateDiv"><% =VS_FechaRequerido %></div>						
				<%	end if %>
				<input type="hidden" id="closingdate" name="closingdate" value="<% =VS_FechaRequerido %>" />
			</td>			
		</tr>	
		<tr>
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
			<td>
				<% 
				if not estaPMReferencia and idVale=0 then
						cant1 = rsAlmacenes.recordCount
						if rsAlmacenes.recordCount = 1 then
							 response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")						        
						%>
							<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =rsAlmacenes("IDALMACEN") %>">
						<%
						else	
						%>
							<select id="idAlmacenCmb" name="idAlmacenCmb" onChange="updateLinkArticulo()">
								<%	
								while (not rsAlmacenes.eof)	
									%>
									<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = vs_idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
									<%	
									rsAlmacenes.MoveNext()
								wend 	
								%>
							</select>
							<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =vs_idAlmacen %>">
						<%								
						end if
						
						if VS_cdVale = CODIGO_PM then
							'Ver si puede transferir
							Set rsAlmacenes2 = obtenerListaAlmacenesSolicitud()
							cant2 = rsAlmacenes2.recordCount
							if (cant2>1) or (cant2<cant1 and cant3>0) then puedeTransferir = true
						end if
				else
						if VS_cdVale = CODIGO_VS_RECEPCION then
							 Set rsAlmacenes = obtenerListaAlmacenes(vs_idAlmacenDest) 
							 if (not rsAlmacenes.eof) then 
								 response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
							 end if%>
							<input type="hidden" name="idAlmacenDest" id="idAlmacenDest" value="<% =vs_idAlmacenDest %>">
						<%else
							 Set rsAlmacenes = obtenerListaAlmacenes(vs_idAlmacen) 
							 if (not rsAlmacenes.eof) then 
								 response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
							 end if%>
							<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =vs_idAlmacen %>">
						<%end if
				end if
				%>
			</td>
		
		
		
		<%
		select case VS_cdVale
			case CODIGO_VS_TRANSFERENCIA
				if not estaPMReferencia then
					%>
					<td class="reg_Header_navdos" colspan="2"><%=GF_TRADUCIR("Transferir a")%></td>
					<td colspan="3">
						<%
						Set rsAlmacenes = obtenerListaAlmacenesSolicitud()
							%>
							<select id="idAlmacenDest" name="idAlmacenDest">
							<%	
								while (not rsAlmacenes.eof)
									%>
									<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = vs_idAlmacenDest) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
									<%		
									rsAlmacenes.MoveNext()
								wend 	
							%>
							</select>
					</td>
					<%		
				else
					%>
					<td class="reg_Header_navdos"><% =GF_TRADUCIR("Transferir a") %></td>
					<td colspan="3">
						<%
							Set rsAlmacenes = obtenerListaAlmacenes(VS_idAlmacenDest) 
							if (not rsAlmacenes.eof) then 
								response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
							end if
						%>
						<input type="hidden" id="idAlmacenDest" name="idAlmacenDest" value="<% =VS_idAlmacenDest %>">
					</td>
					<%
				end if
			case CODIGO_VS_RECEPCION			
				%>
				<td class="reg_Header_navdos"><% =GF_TRADUCIR("Recibir de") %></td>
				<td colspan="3">
					<%
						Set rsAlmacenes = obtenerListaAlmacenes(VS_idAlmacen) 
						if (not rsAlmacenes.eof) then 
							response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
						end if
					%>
					<input type="hidden" id="idAlmacen" name="idAlmacen" value="<% =VS_idAlmacen %>">
				</td>
				<%
			case CODIGO_PM
				if puedeTransferir then 
					if not estaPMReferencia then 
						%>
						<td class="reg_Header_navdos" colspan="2">
							<%=GF_TRADUCIR("Transferir desde")%>
							<input id="almTransferenciaCHK" onclick="ShowHide('almTransferenciaCHK', 'almTransferencia');" type="checkbox" name="transferir" id="transferir" style="border-style:hidden;cursor:pointer;" <%=textChecked%>>
						</td>
						<td colspan="5" align="left">
							
							<div id="almTransferencia" style="float: left;display: none;">
								<table border=0 cellspacing=1 cellpadding=1 width=100%><tr>							
								<td>
									<%
									Set rsAlmacenes = obtenerListaAlmacenesSolicitud()
										%>
										<select id="idAlmacenDest" name="idAlmacenDest">
											<option value="0">- Seleccione -</option>
											<%	
											while (not rsAlmacenes.eof)
												%>
												<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = vs_idAlmacenDest) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
												<%		
												rsAlmacenes.MoveNext()
											wend 	
											%>
										</select>	
								</td>
								</tr>
								</table>
							</div>
						</td>
						<%
					else
						%>
						<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen Destino") %></td>
						<td colspan="6">
							<%
								Set rsAlmacenes = obtenerListaAlmacenes(VS_idAlmacenDest) 
								if (not rsAlmacenes.eof) then 
									response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
								end if
							%>
							<input type="hidden" id="idAlmacenDest" name="idAlmacenDest" value="<% =VS_idAlmacenDest %>">
						</td>
						<%
					end if
					%>
					</tr>
				<%end if			
		end select					
		if ((estaPMReferencia) and (VS_cdVale = CODIGO_VS_SALIDA)) then %>
			<tr>
				<td class="reg_Header_nav" colspan="8"><% =GF_TRADUCIR("Comentario del Pedido de Materiales") %></td>
			</tr>
			<tr>
				<td colspan="8" align=center>
					<table width="80%"><tr><td>
						<%	if (len(PM_comentario) > 0) then 
								Response.write PM_comentario
							else
								Response.Write GF_TRADUCIR("Sin Comentario")
							end if
						%>
					</td></tr></table>
				</td>
			</tr>
		<% end if %>
		<tr>
			<td class="reg_Header_nav" colspan="8"><% =GF_TRADUCIR("Comentario") %></td>
		</tr>
		<tr>
			<td colspan="8" align=center><textarea name="comentario" id="comentario" cols="100"><%=VS_comentario%></textarea>
			</td>
		</tr>

		<tr>
			<td class="reg_Header_nav" colspan="8"><% =GF_TRADUCIR("Detalle") %></td>
		</tr>
		<tr>
			<td colspan="8">
				<table class="reg_Header" width="100%" border=0 id="tblArticulos">
					<tr class="reg_Header_nav">
						<td width="10%" align="center"><% =GF_TRADUCIR("Codigo") %></td>
						<td width="50%"><% =GF_TRADUCIR("Descripcion") %></td>
						<td align="center"><% =GF_TRADUCIR("Cd. Interno") %></td>							
						<%
						if (idVale=0) then%>
							<td align="center"><% =GF_TRADUCIR("Stock") %></td>
						<%
						end if
						if ((VS_cdVale <> CODIGO_VS_ENTRADA) AND (VS_cdVale <> CODIGO_PM) AND idVale=0) then%>
							<td align="center"><% =GF_TRADUCIR(title1) %></td>
							<td align="center"><% =GF_TRADUCIR(title2) %></td>
						<%
						end if
						%>
						<td align="center"><% =GF_TRADUCIR(title3) %></td>
					</tr>
					<tr>
						<td colspan="7" align="right">
						<%	if not estaPMReferencia and idVale=0 then%>
							<img src="images/add.gif" onclick="agregarLineaArticulo();" style="cursor:pointer">
						<%	end if %>
						</td>					
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id="accion" name="accion" value="">
	<input type="hidden" id="idVale" name="idVale" value="<% =idVale %>">
	<input type="hidden" id="cdVale" name="cdVale" value="<% =VS_cdVale %>">
	<input type="hidden" id="TC" name="TC" value="<% =fromTC %>">
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">	
	<input type="hidden" id="esTransferencia" name="esTransferencia" value="<% if (esTransferencia) then Response.Write esTransferencia %>">
</form>
</body>
</html>
<%
'---------------------------------------------------------------------------------------------
sub VS2PM()
	'PM = VS
	PM_FechaSolicitud = VS_FechaSolicitud
	PM_FechaRequerido = VS_FechaRequerido
	PM_cdSolicitante = VS_cdSolicitante
	PM_dsSolicitante = VS_dsSolicitante
	PM_idPedido = VS_idPedido
	PM_idAlmacen = VS_idAlmacen
	PM_idAlmacenDest = VS_idAlmacenDest	
	PM_idObra = VS_idObra
	PM_idSector = VS_idSector
	PM_idBudgetArea = VS_idBudgetArea
	PM_idBudgetDetalle = VS_idBudgetDetalle
	PM_comentario = VS_comentario
	PM_usuario = VS_usuario
	PM_momento = VS_momento
	PM_hayCabecera = VS_hayCabecera
end sub
'---------------------------------------------------------------------------------------------
sub VS2PM_DET()
	PM_idArticulo = VS_idArticulo
	PM_dsArticulo = VS_dsArticulo
	PM_idUnidad = VS_idUnidad
	PM_abreviaturaUnidad = VS_abreviaturaUnidad
	PM_cantidad = VS_cantidad
	PM_saldo = VS_saldo
	PM_cdInterno = VS_cdInterno
	PM_articuloStock = VS_existencia
end sub
'---------------------------------------------------------------------------------------------
sub PM2VS()
	'VS = PM
	VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		
	VS_FechaRequerido = PM_FechaRequerido
	VS_cdSolicitante = PM_cdSolicitante
	VS_dsSolicitante = PM_dsSolicitante
	VS_idPedido = PM_idPedido
	VS_idAlmacen = PM_idAlmacen
	VS_idAlmacenDest = PM_idAlmacenDest
	VS_idObra = PM_idObra
	VS_idBudgetArea = PM_idBudgetArea
	VS_idBudgetDetalle = PM_idBudgetDetalle
	VS_idSector = PM_idSector
	VS_usuario = PM_usuario
	VS_momento = PM_momento
	VS_hayCabecera = PM_hayCabecera
end sub
'---------------------------------------------------------------------------------------------
sub PM2VS_DET()
	VS_idArticulo = PM_idArticulo
	VS_dsArticulo = PM_dsArticulo
	VS_idUnidad = PM_idUnidad
	VS_abreviaturaUnidad = PM_abreviaturaUnidad
	VS_cantidad = PM_cantidad
	VS_saldo = PM_saldo
	VS_cdInterno = PM_cdInterno
	VS_existencia = PM_articuloStock
end sub
%>