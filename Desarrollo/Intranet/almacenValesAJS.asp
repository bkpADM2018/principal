<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Call controlAccesoAL("")
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim idVale, index, controlOK, accion, cambiaPlazo, esCancelable, esPopUp, myOnUnload
Dim rsObras, rsComentarios, strSQL, conn, aceptaProveedor, flagDebeConfirmar, myvaleSalidaComment
dim minCantPro, nrmName, aux
dim rsAlmacenes, devueltos,rsBudget, VS_Stock
dim title1, title2, title3, mainTitle, lastChar
dim flagGrabarVale
dim myIdAlmacen

myIdAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
Set rsAlmacenes = obtenerListaAlmacenesUA()
if (myIdAlmacen = 0) then 
	if (not rsAlmacenes.eof) then myIdAlmacen = rsAlmacenes("IDALMACEN")
end if
'Response.write "ACA" & VS_idAlmacen
call GP_ConfigurarMomentos()
VS_secBudget = 0
VS_secBudgetArea = 0
idVale = GF_PARAMETROS7("idVale",0,6)
VS_cdVale = GF_PARAMETROS7("cdVale","",6)
accion = GF_PARAMETROS7("accion","",6)
resp = GF_PARAMETROS7("resp","",6)
VS_comentario = GF_PARAMETROS7("comentario","",6)
controlOK = false
myNewPM = 0

if not (isFormSubmit()) then
		call initHeaderVale(idVale)
		VS_cdSolicitante = session("Usuario")
		VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
		if (idVale = 0) then 
			if (myIdAlmacen > 0) then	VS_idAlmacen = myIdAlmacen 
		end if			
		call initArticulosVale(idVale)
else
	call initHeaderVale(idVale)
	VS_cdSolicitante = session("Usuario")
	VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
	call initArticulosVale(idVale)
	VS_FechaSolicitud = GF_PARAMETROS7("issuedate", "", 6)	
	if (VS_FechaSolicitud = "") then VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		

	'Controlar el Vale
	controlOK = controlarVale(idVale)
	if ((accion = ACCION_GRABAR) and (controlOK)) then
		VS_ArticuloActual = 0
			call grabarHeaderVale(idVale,0)
			call grabarComentarioVale(idVale, VS_comentario)
			flagGrabarVale = true
			while (readNextArticuloVale(idVale))
				if (VS_saldo >= 0) then					
					call grabarValeDetalle(idVale, 0)
					call actualizarStock()					
				end if	
			wend
			call grabarPreciosVigentesPorArticulo(idVale)
			Call grabarFirmasValeAJS(idVale)
			Response.Redirect "almacenAjustes.asp"
	end if
end if

%>
<html>
<head>
<title>Almacen - Vales</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
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
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	
	<% if flagGrabarVale then %>
		if (confirm("Desea realizar la impresion del vale?")) {
			window.open('almacenValePedidoPrint.asp?idVale=<%=idVale%>','Imprimir Vale');
		}
		
	<% end if %>
	
	//Constantes - Nombre de Campo
	var SUPPLIER_ID = "supplier";
	var SUPPLIER_DESC = "companyName";
	var SUPPLIER_DIV = "supplierDiv";
	var SUPPLIER_MAIL = "supplierMail";
	var SUPPLIER_CT = "cotizacion";
	var ITEM_ID = "item";
	var ITEM_DESC = "articuloItem";
	var ITEM_DIV = "itemDiv";
	var ITEM_STOCK_ACTUAL = "amount";
	var ITEM_STOCK_ACTUAL_TEXT = "amount_text";
	var ITEM_STOCK_ACTUAL_UNIT = "amount_unit";
	
	var ITEM_STOCK_NUEVO = "saldo";
	var ITEM_STOCK_NUEVO_UNIT = "saldo_unit";
	//var ITEM_SALDO = "saldo";
	
	var myAutoCompletesIndexs = {};

	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;
	var lastProveedores = 0;
	var lastArticulos = 0;		
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ch = new channel();	
	var ms = new Array();
	var lastCategory = "";
		
	function aceptarvaleSalida(id) {
		document.getElementById("resp").value = "OK";
		document.getElementById("frmSel").submit();
	}	
	
	function agregarLineaArticulo() {		
		var obj = undefined;
		var tblArticulos = document.getElementById("tblArticulos");
		var rArticulo = tblArticulos.insertRow(lastArticulos+1);
		var index;
		index = 2;
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cStockActual = rArticulo.insertCell(2);
		var cStockNuevo  = rArticulo.insertCell(3);		

		//CODIGO
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

		
		//Stock Actual
		var iStockActual = document.createElement('input');	
		iStockActual.name = ITEM_STOCK_ACTUAL + lastArticulos;
		iStockActual.id = ITEM_STOCK_ACTUAL + lastArticulos;	
		iStockActual.type = 'hidden';
		iStockActual.size= 4;
		var dStockActual = document.createElement('span');		
		dStockActual.id = ITEM_STOCK_ACTUAL_TEXT + lastArticulos;				
		cStockActual.appendChild(dStockActual);				
		var dStockActualUnidad = document.createElement('span');
		dStockActualUnidad.id = ITEM_STOCK_ACTUAL_UNIT + lastArticulos;	
		dStockActualUnidad.align = 'right';
		cStockActual.align = 'right';
		cStockActual.appendChild(iStockActual);
		cStockActual.appendChild(dStockActual);
		cStockActual.appendChild(dStockActualUnidad);
				
		//Nuevo stock
		var iStockNuevo = document.createElement('input');	
		iStockNuevo.name = ITEM_STOCK_NUEVO + lastArticulos;
		iStockNuevo.id = ITEM_STOCK_NUEVO + lastArticulos;		
		iStockNuevo.size= 4;
		if (isFirefox) {
			iStockNuevo.setAttribute('onkeypress', "return controlDatos(this, event, 'N')");
			//iStockNuevo.setAttribute('onblur', "return controlCampo(this, 'N')");				
		} else {
			iStockNuevo['onkeypress'] = new Function("return controlDatos(this, event, 'N')");
			//iStockNuevo['onblur'] = new Function("return controlCampo(this, 'N')");
		}
		var dStockNuevoUnidad = document.createElement('span');
		dStockNuevoUnidad.id = ITEM_STOCK_NUEVO_UNIT + lastArticulos;	
		dStockNuevoUnidad.align = 'right';
		cStockNuevo.align = 'right';
		iStockNuevo.style.textAlign = 'right';
		cStockNuevo.appendChild(iStockNuevo);
		cStockNuevo.appendChild(dStockNuevoUnidad);

		<% if idVale <> 0 then %>
			iStockNuevo.type = 'hidden';
			dStockNuevoUnidad.style.display = 'none';
		<% end if 		
		if (idVale = 0) then 				%>
		
			//Las funciones internas del autocomplete como focus, select o changes, no se crean en el momento de creacion
			//del autocomplete, sino que se ejecutan en esos momentos (al hacer foco, seleccionar item, cambiar valor).
			//Por lo tanto no podemos utilizar la variable 'lastArticulos' para identificar el indice del autocomplete
			//porque tomaria el valor de la variable al momento de ejecutar la accion en lugar del valor al momento de la creacion
			//por tal motivo guardamos un objeto donde tenemos como key el id del autocomplete y como valor el indice del mismo
			//siendo el id del autocomplete accesible mediente this.id podemos identificar dentro de este objeto el indice que
			//necesitamos.
			myAutoCompletesIndexs[ITEM_DESC + lastArticulos + "_text"] = lastArticulos
			
			$( "#"+ITEM_DESC + lastArticulos + "_text" ).autocomplete({
				minLength: 2,
				//El source se setea al seleccionar un almacen
				source: "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen=" + document.getElementById("idAlmacen").value,
				focus: function( event, ui ) {
					$( "#"+ITEM_DESC + myAutoCompletesIndexs[this.id] + "_text" ).val(ui.item.dsarticulo);
					return false;
				},
				select: function( event, ui ) {
					var myIndex = myAutoCompletesIndexs[this.id];
					$( "#"+ITEM_DESC + myIndex + "_text").val (ui.item.dsarticulo);
					$( "#"+ITEM_ID + myIndex).val (ui.item.idarticulo);
					$( "#"+ITEM_DIV + myIndex).html (ui.item.idarticulo);
					$( "#"+ITEM_STOCK_ACTUAL_TEXT + myIndex).html(ui.item.stock);
					$( "#"+ITEM_STOCK_ACTUAL + myIndex).val (ui.item.stock);
					$( "#"+ITEM_STOCK_ACTUAL_UNIT + myIndex).html("&nbsp;"+ui.item.abreviatura);
					$( "#"+ITEM_STOCK_NUEVO + myIndex).val(0);
					$( "#"+ITEM_STOCK_NUEVO_UNIT + myIndex).html("&nbsp;"+ui.item.abreviatura);
					
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
						$( "#"+ITEM_STOCK_ACTUAL_TEXT + myIndex).html("");
						$( "#"+ITEM_STOCK_ACTUAL + myIndex).val ("");
						$( "#"+ITEM_STOCK_ACTUAL_UNIT + myIndex).html("");
						$( "#"+ITEM_STOCK_NUEVO + myIndex).val(0);
						$( "#"+ITEM_STOCK_NUEVO_UNIT + myIndex).html("");
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				if (item.stock == null) {
					item.stock = 0;
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
<%		end if	%>
		lastArticulos++;
		document.getElementById("cantArticulos").value = lastArticulos;
		return obj;
	}

	function updateLinkArticulo() {		
		var cmb = document.getElementById("idAlmacen");		
		for (k in ms) {
			ms[k].autocomplete("option", "source" , "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen="+cmb.options[cmb.selectedIndex].value	 );
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
	
	function fillArticulo(vss, linea, id, desc, stockActual, stockNuevo, unit) {
			$("#articuloItem"+linea+"_text").val(desc);
			
			$("#item"+linea).val(id);
			$("#itemDiv"+linea).html(id);
			
			$("#amount"+linea).val(stockActual);
			$("#amount_text"+linea).html(stockActual);
			$("#amount_unit"+linea).html("&nbsp;"+unit);
			
			$("#saldo"+linea).val(stockNuevo);
			$("#saldo_unit"+linea).html("&nbsp;"+unit),
			
			document.getElementById(ITEM_STOCK_NUEVO + linea).value = stockNuevo;
	}

	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function canSubmit(acc, btn) {		
			submitInfo(acc);		
	}
	
	function irvaleSalidas() {
		location.href = "almacenAdministrarvaleSalidas.asp";
	}
	function irHome() {
		location.href = "almacenIndex.asp";
	}	
	function irAjustes() {
		location.href = "almacenAjustes.asp";
	}	
	function bodyOnLoad() {	
		var myMS;
		tb = new Toolbar('toolbar', 6,'images/almacenes/');									
		tb.addButton("Home-16x16.png", "Home", "irHome()");				
		idBtnGuardar = tb.addButtonSAVE("Guardar", "canSubmit('<% =ACCION_GRABAR %>',0)");
		idBtnControl = tb.addButtonCONFIRM("Controlar", "canSubmit('<% =ACCION_CONTROLAR %>',1)");			
		tb.addButton("Setting_folder-16x16.png", "Ajustes", "irAjustes()");						
		tb.draw();
	<%	
		index = 0
			while (readNextArticuloVale(idVale))
 	%>
				myMS = agregarLineaArticulo();
 				fillArticulo(myMS, <% =index %>, '<% =vs_idArticulo %>', '<% =vs_dsArticulo %>', <%=VS_Cantidad%>, '<% =VS_Saldo%>', '<% =VS_abreviaturaUnidad %>');									
	<%
				index=index+1
			wend	
		while ((index < 5))%>
			agregarLineaArticulo();
		<%index=index+1
		wend	
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
		if (document.getElementById('item' + i ).value == '<%=arrArticulosConErrores(iArticulos)%>') tblArticulos.rows[i+1].className = 'reg_Header_Error';
	  }
    <%next%>
}
function closeWin() {
	parent.location.reload();
}

function keyPressed(e) {
	key=(document.all) ? e.keyCode : e.which;
	if(key==13) return false;
}
</script>
</head>

<script>
</script>
</head>
<body onLoad="bodyOnLoad()" onkeypress="return keyPressed(event)">	
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="almacenValesTitulo.asp?cdVAle=<% =CODIGO_VS_AJUSTE_STOCK %>" method="POST">	
	<table class="reg_Header" align="center" width="70%" border="0">				
		<tr>
			<td colspan="3">
				<%call showErrors()%>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" align="center"><font class="big"><% =ucase(VS_cdVale) %></font></td>
			<td align="center" class="reg_Header_nav" colspan="2"><font class="big"><% =getLeyendaCdVale(ucase(VS_cdVale)) %></font></td>
		</tr>

		<tr>	
		
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
			<td colspan="2">										
				<select id="idAlmacen" name="idAlmacen" onChange="updateLinkArticulo()">							
			<%	while (not rsAlmacenes.eof)	%>
					<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = vs_idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
			<%		rsAlmacenes.MoveNext()
				wend 	%>
				</select>										
			</td>		
		</tr>			

		<tr>
			<td width="20%" class="reg_Header_navdos"><% =GF_TRADUCIR("Responsable") %></td>
			<td width="30%" colspan="1">
				<%response.write VS_dsSolicitante%>
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =VS_cdSolicitante %>"/>
			</td>
			<td width="50%">
				<table cellpadding=0 cellspacing=0 border=0 width=100%>
					<tr>
						<td width="40%" class="reg_Header_navdos"><% =GF_TRADUCIR("Fecha Ajuste") %></td>
						<td width="30%" align=center>
							<div id="issuedateDiv"><% =VS_FechaSolicitud %></div>
							<input type="hidden" id="issuedate" name="issuedate" value="<% =VS_FechaSolicitud %>">
						</td>
					</tr>
				</table>	
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" colspan="3"><% =GF_TRADUCIR("Comentario") %></td>
		</tr>
		<tr>
			<% if idVale <> 0 then %>
				<td colspan="3"><%=getComentarioVale(idVale)%></td>
			<% else	%>
				<td colspan="3" align=center><textarea name="comentario" id="comentario" cols="100"><%=VS_comentario%></textarea>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" colspan="3"><% =GF_TRADUCIR("Detalle") %></td>
		</tr>
		<tr>
			<td colspan="6">
				<table class="reg_Header" width="100%" id="tblArticulos">
					<tr class="reg_Header_nav">
						<td align="center" width="10%"><% =GF_TRADUCIR("Codigo") %></td>												
						<td align="center" width="60%"><% =GF_TRADUCIR("Descripcion") %></td>
						<td align="center" width="15%"><% =GF_TRADUCIR("Stock Actual")%></td>
						<td align="center" width="15%"><% =GF_TRADUCIR("Nuevo Stock") %></td>							
					</tr>
					<tr>
						<td colspan="4" align="right">						
							<img src="images/add.gif" onClick="agregarLineaArticulo();" style="cursor:pointer">						
						</td>					
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id="accion" name="accion" value="">
	<input type="hidden" id="idVale" name="idVale" value="<% =idVale %>">
	<input type="hidden" id="cdVale" name="cdVale" value="<% =VS_cdVale %>">	
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">
	<input type="hidden" name="resp" id="resp" value="MAYBE">		
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
	PM_idBudgetArea = VS_idBudgetArea
	PM_idBudgetDetalle = VS_idBudgetDetalle
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
end sub
'---------------------------------------------------------------------------------------------
sub PM2VS()
	'VS = PM
	VS_FechaSolicitud = PM_FechaSolicitud
	VS_FechaRequerido = PM_FechaRequerido
	VS_cdSolicitante = PM_cdSolicitante
	VS_dsSolicitante = PM_dsSolicitante
	VS_idPedido = PM_idPedido
	VS_idAlmacen = PM_idAlmacen
	VS_idAlmacenDest = PM_idAlmacenDest
	VS_idObra = PM_idObra
	VS_idBudgetArea = PM_idBudgetArea
	VS_idBudgetDetalle = PM_idBudgetDetalle
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
end sub
%>