<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
Function getCantidadArticulosCTST(pIdControl)
	Dim strSQL,rs,rtrn
	strSQL = "SELECT COUNT(IdControl) AS CANT FROM TBLCSTKDETALLE WHERE IDCONTROL = "&pIdControl	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.eof)then rtrn = rs("CANT")	
	getCantidadArticulosCTST = rtrn
end function 
'***************************************************************************************
'*************************** COMIENZO DE LA PAGINA *************************************
'***************************************************************************************
Dim idVale, index, accion ,rsAlmacenes,idControl_new,lastArticulos,cantArt,myIdAlmacen,pSeleccion, flagCargaRes,cdResponsable
Dim myTitulo,rsCab

myIdAlmacen = GF_PARAMETROS7("IdAlmacen",0,6)
idControl_new = GF_PARAMETROS7("IdControl",0,6)
pSeleccion = GF_PARAMETROS7("tipoReporte","",6)
VS_cdVale = GF_PARAMETROS7("cdVale","",6)
lastArticulos = GF_PARAMETROS7("cantArticulos", 0, 6)
idVale = GF_PARAMETROS7("idVale",0,6)
accion = GF_PARAMETROS7("accion","",6)
cdResponsable = GF_PARAMETROS7("cdResponsable","",6)
flagGrabar = false
flagCargaRes = false
cantArt = 10
VS_idAlmacen = myIdAlmacen


if(idControl_new > 0)then	
	myTitulo = "Carga de Resultados"
	cantArt = getCantidadArticulosCTST(idControl_new)
	flagCargaRes = true
	' Lee los datos de la Cabcera del control desde la DB (carga de resultados)
	Set rsCab = leerCabeceraCtSt(idControl_new)  
	if (not rsCab.EoF)then		
		cdResponsable = rsCab("CDRESPONSABLE")
		myIdAlmacen   = rsCab("IDALMACEN")
		pSeleccion    = rsCab("TIPO")
	end if	
else
	myTitulo = "Seleccion Manual de Articulos"
end if

%>
<html>
<head>
<title><%=myTitulo%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
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
	.leyenda {
		font-weight: bold;
		font-size: 8px;
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
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var ch = new channel();
	var ITEM_ID = "item";
	var ITEM_DESC = "articuloItem";
	var ITEM_DIV = "itemDiv";
	var ITEM_STOCK_ACTUAL = "amount";
	var ITEM_STOCK_ACTUAL_TEXT = "amount_text";
	var ITEM_STOCK_ACTUAL_UNIT = "amount_unit";		
	var ITEM_STOCK_NUEVO = "saldo";
	var ITEM_STOCK_NUEVO_UNIT = "saldo_unit";
	var myAutoCompletesIndexs = {};
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var tb;	
	var lastArticulos = 0;
	var idBtnGuardar = 0;	
	var ms = new Array();
	var lastCategory = "";			

	function bodyOnLoad() {		
		var myMS;
		document.getElementById("cantArticulos").value = lastArticulos;
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		idBtnGuardar = tb.addButtonSAVE("Guardar", "InitManual()");
		tb.draw();
		<%index = 0
		while (index < cantArt)	%>
			myMS = agregarLineaArticulo();
		<%  index=index + 1
		wend 
		if(idControl_new > 0)then%>
			obtenerListResultado();
		<%end if%>
		pngfix();
	}

	function InitManual(){	
		document.getElementById("frmSel").action="almacenCtStGrabar.asp";		
		document.getElementById("frmSel").target= "ifrmSelec";		
		document.getElementById("frmSel").submit();		
	}
	
	function obtenerListResultado(){
		ch.bind("almacenCtSt_Ajax.asp?IdControl=<%=idControl_new%>&IdAlmacen=<%=myIdAlmacen%>","obtenerListResultado_callback()");
		ch.send();
	}
	
	function obtenerListResultado_callback(){		
		var rtrn = ch.response();
		var arr = rtrn.split(";");			
		for (i in arr) {
			if(i > 0){
				var vals = arr[i].split("|");
				fillArticulo( i-1, vals[0], vals[1], vals[3], vals[3], vals[4])
			}
		}							
	}
	
	function resultadoCarga_callback(pMsj,idControl){
		if(pMsj == 1)
		{			
			window.close();
		}
		else	
		{			
			document.getElementById("avisoCtSk").innerHTML="<% =GF_TRADUCIR("Se cargo correctamente.") %>";
			document.getElementById("avisoCtSk").className = "TDSUCCESS";
			window.returnValue = true;
			if(idControl > 0){
				var pp = new winPopUp('popupNuevoCtSt', 'almacenCtStConfirmarControl.asp?idcontrol='+idControl, '220', '120', 'Nuevo Control de Stock', 'CerrarVentana()');
			}			
		}
	}	
	
	function CerrarVentana()
	{		
		window.close();
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
		<%if(flagCargaRes)then%>
			var cStockNuevo  = rArticulo.insertCell(3);
		<%end if%>

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
		
		<%if(flagCargaRes)then%>
			var iDescripcion = document.createElement('div');		
			iDescripcion.id = ITEM_DESC + lastArticulos;				
			cDescripcion.appendChild(iDescripcion);	
		<%else%>	
			//TEXTO
			var iText = document.createElement('input');		
			iText.type="text";
			iText.id = ITEM_DESC + lastArticulos + "_text" ;
			iText.name = ITEM_DESC + lastArticulos + "_text";
			iText.size = 60;
			
			//DESCRIPCION
			var iDescripcion = document.createElement('div');		
			iDescripcion.id = ITEM_DESC + lastArticulos;				
			iDescripcion.appendChild(iText);
			cDescripcion.appendChild(iDescripcion);		
		<%end if%>	
		
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
		<%if(flagCargaRes)then%>
			var iStockNuevo = document.createElement('input');
			iStockNuevo.name = ITEM_STOCK_NUEVO + lastArticulos;
			iStockNuevo.id = ITEM_STOCK_NUEVO + lastArticulos;
			iStockNuevo.size= 8;
			iStockNuevo.value = 0;
			var dStockNuevoUnidad = document.createElement('span');
			dStockNuevoUnidad.id = ITEM_STOCK_NUEVO_UNIT + lastArticulos;
			dStockNuevoUnidad.align = 'right';
			cStockNuevo.align = 'right';
			iStockNuevo.style.textAlign = 'right';
			cStockNuevo.appendChild(iStockNuevo);
			cStockNuevo.appendChild(dStockNuevoUnidad);
			iStockNuevo.type = 'text';
			dStockNuevoUnidad.style.display = 'none';
		<%end if%>
		
		<%if ((idVale = 0)and(flagCargaRes = false)) then	%>
			myAutoCompletesIndexs[ITEM_DESC + lastArticulos + "_text"] = lastArticulos
			$( "#"+ITEM_DESC + lastArticulos + "_text" ).autocomplete({
				minLength: 2,
				//El source se setea al seleccionar un almacen
				source: "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen=<%=VS_idAlmacen%>",
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
					<%if(flagCargaRes)then%>
						$( "#"+ITEM_STOCK_NUEVO + myIndex).val(0);
						$( "#"+ITEM_STOCK_NUEVO_UNIT + myIndex).html("&nbsp;"+ui.item.abreviatura);
					<%end if%>
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
						<%if(flagCargaRes)then%>
							$( "#"+ITEM_STOCK_NUEVO + myIndex).val(0);
							$( "#"+ITEM_STOCK_NUEVO_UNIT + myIndex).html("");
						<%end if%>
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

	function fillArticulo(linea, id, desc, stockActual, stockNuevo, unit) {			
			$("#item"+linea).val(id);
			$("#itemDiv"+linea).html(id);
			$("#amount"+linea).val(stockActual);
			$("#amount_text"+linea).html(stockActual);
			$("#amount_unit"+linea).html("&nbsp;"+unit);			
			$("#saldo"+linea).val(stockNuevo);
			$("#saldo_unit"+linea).html("&nbsp;"+unit),			
			document.getElementById(ITEM_DESC + linea).innerHTML = desc;
			document.getElementById(ITEM_STOCK_NUEVO + linea).value = stockNuevo;			
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
<% call GF_TITULO2("kogge64.gif",myTitulo) %>
<div id="toolbar"></div>
<br>		

<form id="frmSel" name="frmSel" action="almacenCtStSelecciones.asp" method="POST"  >		
	
	<table class="reg_Header" align="center" width="80%" >	
	<tr><td colspan="4"><div id="avisoCtSk" align="center" class="TDBAJAS"></div></td></tr>				
	<tr>		
		<td colspan="4"><% call showErrors() %></td>
	</tr>		
		<tr>
			<td align="center" class="reg_Header_nav" width="20%"><%=GF_TRADUCIR("Almacen")%>:</td>
			<%Set rsAlmacenes = obtenerListaAlmacenes(myIdAlmacen)%>
			<td align="center" class="reg_header_navdos" width="30%"><%=rsAlmacenes("CDAlmacen")%></td>
			<td align="center" class="reg_Header_nav" width="20%"><%=GF_TRADUCIR("Responsable")%>:</td>
			<td align="center" class="reg_header_navdos" width="30%"><%=getUserDescription(cdResponsable)%></td>
		</tr>				
		<tr>
			<td colspan="4">
				<table  width="100%" id="tblArticulos">
					<tr class="reg_Header_nav">
						<td align="center" width="10%"><% =GF_TRADUCIR("Codigo") %></td>
						<td align="center" width="60%"><% =GF_TRADUCIR("Descripcion") %></td>
						<td align="center" width="15%"><% =GF_TRADUCIR("Stock Actual")%></td>						
						<%if(flagCargaRes)then%>
							<td align="center" width="15%"><% =GF_TRADUCIR("Stock Nuevo")%></td>
						<%end if%>
					</tr>
					<%if(flagCargaRes = false)then%>
					<tr>
						<td colspan="3" align="right">
							<img src="images/add.gif" onClick="agregarLineaArticulo();" style="cursor:pointer">							
						</td>					
					</tr>
					<%end if%>
					<% if(idControl_new > 0)then %>
					<tr>
						<td align="right" class="leyenda" colspan="4"><% =GF_TRADUCIR("El Stock de Sistema se calculó al momento de emitirse el control de stock.") %></td>
					</tr>
					<% end if %>	
				</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id="accion" name="accion" value="<%=ACCION_GRABAR%>">
	<input type="hidden" id="tipoReporte" name="tipoReporte" value="<%=pSeleccion%>">
	<input type="hidden" id="IdAlmacen" name="IdAlmacen" value="<% =myIdAlmacen %>">
	<input type="hidden" id="IdControl" name="IdControl" value="<% =idControl_new %>">
	<input type="hidden" id="idVale" name="idVale" value="<% =idVale %>">
	<input type="hidden" id="cdVale" name="cdVale" value="<% =VS_cdVale %>">
	<input type="hidden" id="CtSt_cdResponsable" name="CtSt_cdResponsable" value="<% =cdResponsable %>">	
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">
</form>
<iframe name="ifrmSelec" id="ifrmSelec" width="1px" height="1px" style="visibility:hidden"></iframe>
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