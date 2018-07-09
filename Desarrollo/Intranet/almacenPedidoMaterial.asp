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
Const LINEAS_DETALLE_DEFAULT = 5
'------------------------------------------------------------------------------------------------
'************************************************************************************************
'*************************************** COMIENZO DE LA PAGINA **********************************
'************************************************************************************************
Call controlAccesoAL("")
Dim IdAlmacen,comentario,fechaSolicitud, idVale, index, myDivision,idPMReferencia,constolOK,myJSClose, isTransferencia,almacenOld
idVale = 0
index = 0
Call GP_ConfigurarMomentos()
flagSubmit = false
	
accion = GF_PARAMETROS7("accion","",6)
PM_idAlmacen	  = GF_PARAMETROS7("cmbIdAlmacen", 0, 6)
PM_comentario	  = GF_PARAMETROS7("comentario","",6)
PM_FechaSolicitud = GF_PARAMETROS7("issuedate", "", 6)
if (PM_FechaSolicitud = "") then PM_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))
PM_FechaRequerido = GF_PARAMETROS7("closingdate","",6)
if(PM_FechaRequerido = "") Then PM_FechaRequerido = GF_FN2DTE(Left(session("MmtoSistema"),8))
PM_idAlmacenDest  = GF_PARAMETROS7("idAlmacenDest",0,6)
rowNum			  = GF_PARAMETROS7("rowNum",0,6)
myDivision		  = GF_PARAMETROS7("idDivision",0,6)
PM_Transferir	  = GF_PARAMETROS7("chkTransferir",0,6)
PM_cdSolicitante  = GF_PARAMETROS7("cdSolicitante","",6)
PM_dsSolicitante  = getUserDescription(PM_cdSolicitante)
PM_CantDetalle    = GF_PARAMETROS7("rowDetalle",0,6)
if(PM_CantDetalle = 0) Then PM_CantDetalle = 1 
artDefault		  = GF_PARAMETROS7("artDefault",0,6)
almacenOld		  = GF_PARAMETROS7("AlmacenOld",0,6)
if(almacenOld = 0)then almacenOld = PM_idAlmacen
PM_articuloError = 0
if(artDefault = 0)then artDefault = LINEAS_DETALLE_DEFAULT
estaPMReferencia = false
flagTransferencia = false

Set rsAlmacenes = obtenerListaAlmacenesSolicitud()
if (PM_idAlmacen = 0) then 
	if (not rsAlmacenes.Eof) then 
		PM_idAlmacen = rsAlmacenes("IDALMACEN")
		myDivision   = rsAlmacenes("IDDIVISION")
	end if	
end if
if tieneAccesoTransferenciaPM(PM_idAlmacen) then flagTransferencia = true
if Not flagTransferencia then PM_Transferir = 0 
if PM_Transferir = 0 then PM_idAlmacenDest = 0
if isFormSubmit() then	
	flagSubmit = true
	estaPMReferencia = true
	constolOK = controlarPM()
	if ((accion = ACCION_GRABAR) and (constolOK)) then
		while (readNextDetalleParamsPM())
			idPMReferencia = grabarHeaderPMInsert()			
			PM_ArticuloActual = 1			
			while (readNextArticuloParamsPM())				
				if (PM_idArticulo > 0) then
					Call grabarPMDetalle(idPMReferencia, PM_idArticulo, PM_cantidad, PM_saldo)
				end if				
				PM_ArticuloActual = PM_ArticuloActual + 1
			wend				
			PM_DetalleActual = PM_DetalleActual +  1
		wend
		myJSClose = "location.href= 'almacenAdministrarPedidosMateriales.asp'"
	end if
end if

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
	
	<%=myJSClose%>	
	var myAutoCompletesIndexs = {};
	var lastCategory = "";
	var articuloCounter = new Array();	//guarda en cada indice(numero de seccion) la cantidad de lineas de articulo que tiene		
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");		
	var rowDetalle = 0;	//es el numero de seccion que hay actualmente   
	var rowArticulo = 0; 	
	var rowNum = 0;
	var idBtnGuardar = 0;
	var idBtnControl = 0;	
	var ch = new channel();		
	var ms = new Array();
	var IdObra;
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
	var strMasterSelectDefault="";
	var strAreaDetalleDefault="";
	var valMasterSelectDefault = "";
	
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		idBtnGuardar = tb.addButtonSAVE("Guardar", "canSubmit('<% =ACCION_GRABAR %>')");
		idBtnControl = tb.addButtonCONFIRM("Controlar", "canSubmit('<% =ACCION_CONTROLAR %>')");	
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irTDC()");			
		tb.addButton("../compras/refresh-16x16.png", "Recargar", "canSubmit('<% =ACCION_PROCESAR %>')");
		tb.draw();		 
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 4, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;
		msSolicitante.setValue('<% =PM_dsSolicitante %>');	  		
		var idDivision = $("#idDivision").val();
		<%if PM_Transferir = 1 then%>			
			document.getElementById("linkAddLinea").style.display = "none";
			articuloCounter[0] = 0;
			crearTablaArticulo(0);
		<%	if not (estaPMReferencia) then %>				
			<%	PM_ArticuloActual = 1
				while (PM_ArticuloActual < 6) %>
					newLineaArticulo(0);
				<%  PM_ArticuloActual = PM_ArticuloActual + 1
				wend
			else %>				
			<%	PM_ArticuloActual = 1
				PM_DetalleActual  = 0 %>				
			<%	while (readNextArticuloParamsPM()) %>
					newLineaArticulo(0);
					loadLineaArticulo(0, <%=PM_ArticuloActual%>, <%=PM_idArticulo%>,'<%= PM_dsArticulo%>', '<%= PM_abreviaturaUnidad%>','<%= PM_cdInterno%>', '<%= PM_saldo%>', <%= PM_cantidad%>);
			<%		PM_ArticuloActual = PM_ArticuloActual + 1
				wend
			end if %>
		<%else%> 
			/*Si no es una transferencia al iniciar la pagina lo primero que tengo es el Master Combo (Partida/Sector) de la Division seleccionada*/
			$.ajax({
			url: "almacenObtenerBudget.asp?accion=<%=ACCION_CONTROLAR%>&idObra=0&idDivision="+idDivision+"&idBudgetArea=0&idBudgetDetalle=0",
			success: function(data){
				strMasterSelectDefault = data;
				document.getElementById("divMastesSelectOld").innerHTML = data;
				$("#masterSelect").each(function(){
					valMasterSelectDefault = $(this).val();					
				})				
			<%  PM_DetalleActual = 0
				while (readNextDetalleParamsPM())%>
					newLineaDetalle(<%=PM_DetalleActual%>);
					loadLineaDetalle(<%=PM_DetalleActual%>,<%= PM_idObra%>, <%= PM_idSector%>,<%= PM_idBudgetArea%>, <%= PM_idBudgetDetalle%>, <%= PM_Tipo%>);
					articuloCounter[<%=PM_DetalleActual%>] = 0;
					crearTablaArticulo(<%=PM_DetalleActual%>);
				<%	if not (estaPMReferencia) then %>
					<%	PM_ArticuloActual = 1
						while ((PM_ArticuloActual < 6)) %>
							newLineaArticulo(<%=PM_DetalleActual%>);
						<%  PM_ArticuloActual = PM_ArticuloActual + 1
						wend
					else %>
					<%	PM_ArticuloActual = 1
						while (readNextArticuloParamsPM()) %>
							newLineaArticulo(<%=PM_DetalleActual%>);
							loadLineaArticulo(<%=PM_DetalleActual%>, <%=PM_ArticuloActual%>, <%=PM_idArticulo%>,'<%= PM_dsArticulo%>', '<%= PM_abreviaturaUnidad%>','<%= PM_cdInterno%>', '<%= PM_saldo%>', <%= PM_cantidad%>);
					<%		PM_ArticuloActual = PM_ArticuloActual + 1
						wend
					end if
					PM_DetalleActual = PM_DetalleActual + 1 %>
					document.getElementById("rowDetalle").value = <%=PM_DetalleActual%>;
			<%	wend  %>
				}
			});
		<%end if%>
		pngfix();		
	}
		
	
	/* function loadLineaDetalle : carga los Combo Box de  Partida/Sector y Area/Detalle.
								   La carga inicial trae por Ajax el combo Area/Detalle de la obra por defecto(Mantenimiento)
								   y lo guarda para su posible uso en nuevas lineas.
								   En caso de que se submitio la pagina, toma los valores de cada linea(Sector/Obra y Area/Detalle)
								   y los carga al combo.
							pObra: es la obra que debe buscar el detalle si viene cargado		    
							pRowDetalle: es la linea de detelle actual						
							pSector: es el sector que viene cargado
							pArea: area que viene cargada
							pDetalle: detalle que viene cargado	*/
	function loadLineaDetalle(pRowDetalle ,pObra, pSector,pArea, pDetalle, pTipo){
		var valMasterSelect;		
		document.getElementById("divMastesSelect_" + pRowDetalle).innerHTML = strMasterSelectDefault;
		renameComboBox(pRowDetalle, "divMastesSelect_","cmbPartidaSector_");
		document.getElementById("hidBoleanPartidaSector_" + pRowDetalle).value = pTipo;		
		<% if (almacenOld <> PM_idAlmacen) then %>
			valMasterSelect = valMasterSelectDefault;
			$("#cmbPartidaSector_"+ pRowDetalle +" option[value="+ valMasterSelect +"]").attr("selected",true);
		<% else %>
			if((pObra > 0)||(pSector > 0)){
				var valOption = pSector;
				if(pTipo == 0) valOption = pObra;
				$("#cmbPartidaSector_"+ pRowDetalle +" option[value="+ valOption +"]").attr("selected",true);
				valMasterSelect = valOption;
			}
			else{
				valMasterSelect = valMasterSelectDefault;
				$("#cmbPartidaSector_"+ pRowDetalle +" option[value="+ valMasterSelect +"]").attr("selected",true);
			}
		<% end if 
		   if isFormSubmit() then %>
			 document.getElementById("hiddenValueArea_" + pRowDetalle).value = pArea;			 
			 if(pTipo == 0){
				getAreaDetalleAjax(valMasterSelect, pRowDetalle,pArea, pDetalle);
				if((pDetalle > 0)&&(pArea > 0)) $("#cmbAreaDetalle_" + pRowDetalle + " option[value="+ pArea +"][alt="+ pDetalle +"]").attr("selected",true);
			 }
	   <% else	%>
			if(pRowDetalle == 0){
				getAreaDetalleAjax(valMasterSelectDefault, pRowDetalle,0, 0);
			}
			else{
				document.getElementById("divPPAreaDetalle_" + pRowDetalle).innerHTML = strAreaDetalleDefault;
				renameComboBox(pRowDetalle, "divPPAreaDetalle_", "cmbAreaDetalle_");
			}
	   <% end if %>
	   document.getElementById("AlmacenOld").value = <%=PM_idAlmacen%>;
	}
	
	/* function getAreaDetalleAjax : trae por ajax el combo box del Area/Detalle de una determinada obra. */
	function getAreaDetalleAjax(pIdObra,pRowDetalle,pDetalle,pArea){
		$.ajax({				
			url: "almacenObtenerBudget.asp?accion=<%=ACCION_PROCESAR%>&idObra=" + pIdObra + "&idBudgetArea=0&idBudgetDetalle=0&readOnly=0",
			success: function(data){					
					if(strAreaDetalleDefault.length == 0) strAreaDetalleDefault = data;
					document.getElementById("divPPAreaDetalle_" + pRowDetalle).innerHTML = data;
					renameComboBox(pRowDetalle, "divPPAreaDetalle_", "cmbAreaDetalle_");
					if((pDetalle > 0)&&(pArea > 0))  $("#cmbAreaDetalle_" + pRowDetalle + " option[value="+ pArea +"][alt="+ pDetalle +"]").attr("selected",true);
				}
		});		
	}
	
	/* function renameComboBox : a un ComboBox le asigna un id y name nuevo con Indice asociativo a cada seccion.
								Se lo hace debido a que es traido por ajax con Id y Name generico (todos iguales)
						pIndex = indice de la Partida/Sector que pertenece 
						pIdParent = Id del elemento Padre que lo contiene (es el div creado)
						pIdNew = Id nuevo que va a tener el combo box */
	function renameComboBox(pIndex, pIdParent, pIdNew){
		$("#"+ pIdParent + pIndex).each(function(){			
			var id = "#" + this.id;
			$(id + " select").attr("id", pIdNew + pIndex);
			$(id + " select").attr("name", pIdNew + pIndex);
		})
	}
	
	/*function loadLineaArticulo: Carga los valores de los articulos correspondiente a cada fila,
					 pRowDetalle: Numero de fila de la linea del detalle (Sector/Obra y Area/Detalle) 
					pRowArticulo: Numero de fila de la linea del articulo
				     pIdArticulo: id Articulo
				     pDsArticulo: Descripcion del Articulo
				    pAbreviatura: Descripcion de la Abreviatura de la unidad
					  pCdInterno: Codigo Interno
					      pStock: Stock Articulo
					   pCantidad: Cantidad solicitada					*/
	function loadLineaArticulo(pRowDetalle, pRowArticulo, pIdArticulo, pDsArticulo, pAbreviatura, pCdInterno, pStock, pCantidad){
		if(pIdArticulo > 0){
			document.getElementById(ITEM_ID + pRowDetalle + "_" + pRowArticulo).value = pIdArticulo;
			document.getElementById(ITEM_DESC + pRowDetalle + "_" + pRowArticulo + "_text").value = pDsArticulo;
			document.getElementById(ITEM_DIV + pRowDetalle + "_" + pRowArticulo).innerHTML = pIdArticulo;
			document.getElementById(ITEM_AMOUNT_UNIT + pRowDetalle + "_" + pRowArticulo).innerHTML = pAbreviatura;
			document.getElementById(ITEM_CD_INTERNO + pRowDetalle + "_" + pRowArticulo).innerHTML = pCdInterno;
			document.getElementById(ITEM_STOCK_TEXT + pRowDetalle + "_" + pRowArticulo).value = pStock;
			document.getElementById(ITEM_STOCK + pRowDetalle + "_" + pRowArticulo).innerHTML = pStock;			
			document.getElementById(ITEM_AMOUNT+ pRowDetalle + "_" + pRowArticulo).value = pCantidad;
			if ("<%=PM_articuloError%>" ==  pIdArticulo ){
				var tblArticulos = document.getElementById("tblArticulo_" + pRowDetalle);
				tblArticulos.rows[pRowArticulo].className = 'reg_Header_Error';
			}
			
		}	
	}
		

	/* function addLinea : agrega una linea nueva, va a contener la linea de Partida/Sectror y 5 lineas de articulos*/
	function addLinea(){
		var rowDetalle = document.getElementById("rowDetalle").value;
		newLineaDetalle(rowDetalle);
		loadLineaDetalle(rowDetalle,0,0,0,0,0)		
		articuloCounter[rowDetalle] = 0;
		crearTablaArticulo(rowDetalle);
	<%	PM_ArticuloActual = 0
		while ((PM_ArticuloActual < 5)) %>
			newLineaArticulo(rowDetalle);
	<%		PM_ArticuloActual = PM_ArticuloActual + 1
		wend %>			
		document.getElementById("rowDetalle").value = parseInt(rowDetalle) + 1;
	}
	
	/*function resetForm(me) : Esta funcion se produce cuando se cambia el Almacen de la cabecera o cuando se carga
							   una transferencia, lo que hace es volver la pagina a la condicion inicial 
							   de carga(borra todo el Detalle)     */
	function resetForm(me){
		if(me.tagName == "SELECT") {
			document.getElementById("idDivision").value = $("#" + me.id + " option:selected").attr("alt");						
		}			
		document.getElementById("rowDetalle").value = 0;
		document.getElementById("rowNum").value = 0;
		document.getElementById("rowArticulo").value = 0;
		submitInfo('<% =ACCION_PROCESAR %>');
	}
	
	/*function actualizarAreaDetalle : Actualiza el estado del Combo Box de Area/Detalle y sus valores, si el valor selccionado 
									   en el Master Select es una Partida, llama por ajax y trae el nuevo Combo con 
									   las Areas/Detalles actualizadas. Si eligio un Sector oculta el Combo Box de Area/Detalle  
				Parametro: 	me (es el evento del Combo Box cuando se produce un cambio)	*/
	function actualizarAreaDetalle(me){
		var valTipoMasterSelect = me.options[me.selectedIndex].parentNode.id;
		var row = document.getElementById(me.parentNode.id).nextSibling.value;		
		$("#hidBoleanPartidaSector_" + row).val(valTipoMasterSelect);
		var obj = document.getElementById("cmbPartidaSector_" + row);
		var seleccionado = obj.options[obj.options.selectedIndex];		
		if(seleccionado.parentNode.getAttribute("id") == 0){
			//El nuevo item seleccionado es una Partida Presupuestaria
			document.getElementById("divPPAreaDetalle_" + row).style.display = "block";			
			getAreaDetalleAjax(obj.value, row,0,0);
		}
		else{
			//El nuevo item selccionado es un Sector			
			document.getElementById("divPPAreaDetalle_" + row).style.display = "none";				
		}
	}
	
	/* function newLineaDetalle : agrega una linea de Partida/Sector dentro de la tabla tblDetalle
							pRowDetalle = es el numero de fila a agregar	*/
	function newLineaDetalle(pRowDetalle) {		
		var tblDetalle = document.getElementById("tblDetalle");		
		var rDetalle   = tblDetalle.insertRow(parseInt(rowNum));		
		rDetalle.id= "filaD_" + rowNum;		
		//Creamos la fila Partida/Sector de cada seccion
		var cTitle 	   = rDetalle.insertCell(0);  // Partida/sector
		var cTipo  	   = rDetalle.insertCell(1);  // Partida/sector
		var cPartida   = rDetalle.insertCell(2);  // En caso de que sea Partida, es el detalle		
		/*	Div divMastesSelect_X : guardamos el ComboBox de la Partida/Sector */
		var iTitle = document.createElement('span');
		iTitle.id = "spanTitlePartidaSector_"  + pRowDetalle;
		iTitle.name = "spanTitlePartidaSector_"  + pRowDetalle;
		cTitle.appendChild(iTitle);
		cTitle.className = "reg_Header_nav";
		document.getElementById("spanTitlePartidaSector_" + pRowDetalle).innerHTML = "Partida/Sector:"		
		/*	Div divMastesSelect_X : guardamos el ComboBox de la Partida/Sector */
		var iTipo = document.createElement('div');
		iTipo.id = "divMastesSelect_"  + pRowDetalle;
		iTipo.name = "divMastesSelect_"  + pRowDetalle;
		cTipo.appendChild(iTipo);
		cTipo.setAttribute('width',"40%");						
		/*	hidFilaDet _X : guardamos el valor de la fila actual de Partida/Sector */
		var hidFilaDet = document.createElement('input');
		hidFilaDet.type = 'hidden';
		hidFilaDet.id = "hidRowDetalle_"  + pRowDetalle;
		hidFilaDet.name = "hidRowDetalle_"  + pRowDetalle;
		hidFilaDet.value = pRowDetalle;
		cTipo.appendChild(hidFilaDet);		
		/*	Input Hidden hidBoleanPartidaSector_X : guardaremos en un input hidden el valor 0 caso de que sea Partida,
													y 1	en caso de que se seleccioe un sector */
		var hidRow = document.createElement('input');
		hidRow.type = 'hidden';
		hidRow.id = "hidBoleanPartidaSector_"  + pRowDetalle;
		hidRow.name = "hidBoleanPartidaSector_"  + pRowDetalle;		
		cTipo.appendChild(hidRow);		
		/*	Input Hidden hiddenTipoSel_X : guardamos el indice del valor seleccionado
										-si es 0 es Partida 
										-si es 1 es Sector					*/
		var hidTipo = document.createElement('input');
		hidTipo.type = 'hidden';
		hidTipo.id = "hiddenTipoSel_"  + pRowDetalle;
		hidTipo.name = "hiddenTipoSel_"  + pRowDetalle;
		iTipo.appendChild(hidTipo);		
		cPartida.setAttribute('width',"50%");
		/*	Div divPPAreaDetalle_X : guardamos el ComboBox del Area - Detalle (en caso que selecciono Partida Pres.) */
		var iPartida = document.createElement('div');
		iPartida.id = "divPPAreaDetalle_"  + pRowDetalle;
		iPartida.name = "divPPAreaDetalle_"  + pRowDetalle;		
		cPartida.appendChild(iPartida);		
		/*	hiddenValueArea_X : guardaremos el valor del Area selccionado (en caso que selecciono Partida Pres.)*/
		var hidAreaDet = document.createElement('input');
		hidAreaDet.type = 'hidden';
		hidAreaDet.id = "hiddenValueArea_"  + pRowDetalle;
		hidAreaDet.name = "hiddenValueArea_"  + pRowDetalle;		
		cPartida.appendChild(hidAreaDet);		
		rowNum++;
		document.getElementById("rowNum").value = rowNum
	}
	
	/* function newLineaArticulo : agrega una linea de articulos dentro de la tabla tblArticulos 
							pNumTable = es el indice de la tabla que se agrega la linea 				*/
	function newLineaArticulo(pNumTable){	
		var tblArticulo  = document.getElementById("tblArticulo_"  + pNumTable);		
		var rArticulo	 = tblArticulo.insertRow(articuloCounter[pNumTable]);		
		var cCodigo		 = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cInterno	 = rArticulo.insertCell(2);
		var cStock		 = rArticulo.insertCell(3);
		var cPedido		 = rArticulo.insertCell(4);
		var cImgAddArt	 = rArticulo.insertCell(5);
		var myRowArticulo = pNumTable + "_" + articuloCounter[pNumTable];
					
		//CODIGO
		var iCodigo = document.createElement('input');
		iCodigo.type = "hidden";
		iCodigo.id = ITEM_ID + myRowArticulo;
		iCodigo.name = ITEM_ID + myRowArticulo;
		iCodigo.size= 7;
		iCodigo.maxLength = 5;
		cCodigo.appendChild(iCodigo);
		var dCodigo = document.createElement('div');
		dCodigo.className = "labelStyle";
		dCodigo.id = ITEM_DIV + myRowArticulo;
		cCodigo.appendChild(dCodigo);
		
		//DESCRIPCION
		var iDescripcion = document.createElement('input');		
		iDescripcion.type = "text";
		iDescripcion.id   = ITEM_DESC + myRowArticulo + "_text" ;
		iDescripcion.name = ITEM_DESC + myRowArticulo + "_text";
		iDescripcion.size = 50;
		var iDescripcionDiv = document.createElement('div');		
		iDescripcionDiv.id = ITEM_DESC + myRowArticulo;				
		cDescripcion.appendChild(iDescripcion);
		cDescripcion.appendChild(iDescripcionDiv);		
		
		//CODIGO INTERNO
		cInterno.align = 'center';
		var iCdInterno = document.createElement('div');		
		iCdInterno.id = ITEM_CD_INTERNO + myRowArticulo;
		cInterno.appendChild(iCdInterno);

		//STOCK
		cStock.align = 'center';
		var iStock = document.createElement('span');		
		iStock.id = ITEM_STOCK + myRowArticulo;
		cStock.appendChild(iStock);	
		var iTxtStock = document.createElement('input');
		iTxtStock.type = "hidden";
		iTxtStock.name = ITEM_STOCK_TEXT + myRowArticulo;
		iTxtStock.id = ITEM_STOCK_TEXT + myRowArticulo;
		cStock.appendChild(iTxtStock);	
		
		//PEDIDOS
		cPedido.align = 'right';
		var iCantidad = document.createElement('input');
		iCantidad.id = ITEM_AMOUNT + myRowArticulo;
		iCantidad.name = ITEM_AMOUNT + myRowArticulo;
		iCantidad.size= 4;
		iCantidad.align = 'center';
		iCantidad.style.textAlign = "right";
		if (isFirefox) {
			iCantidad.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");			
		} else {
			iCantidad['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}		
		cPedido.appendChild(iCantidad);
		var dCantidadUnidad = document.createElement('span');
		dCantidadUnidad.id = ITEM_AMOUNT_UNIT + myRowArticulo;
		dCantidadUnidad.style.textAlign = "right";
		cPedido.appendChild(dCantidadUnidad);				
		
		//	IMAGEN PARA AGREGAR DETALLE (PARTIDA O SECTORES)				
		if(articuloCounter[pNumTable] > 0) $("#imgAddArt_" + pNumTable + "_" + (articuloCounter[pNumTable]-1)).remove();
		var iImgAddArt = document.createElement('img');
		iImgAddArt.id = "imgAddArt_"  + myRowArticulo;
		iImgAddArt.name = "imgAddArt_"  + myRowArticulo;
		iImgAddArt.src = "images/compras/add-16x16.png";
		iImgAddArt.title = "Agregar Articulo";
		iImgAddArt.alt = "ImgArticulo";
		iImgAddArt.setAttribute('style', "cursor:pointer;");
		if (isFirefox) {
			iImgAddArt.setAttribute('onclick', "newLineaArticulo("+pNumTable+")");			
		} else {
			iImgAddArt['onclick'] = new Function("newLineaArticulo("+pNumTable+")");			
		}
		cImgAddArt.align = "center";
		cImgAddArt.appendChild(iImgAddArt);
		//SALDO		
		var iSaldo = document.createElement('input');	
		var dSaldo = document.createElement('span');	
		iSaldo.name = ITEM_SALDO + myRowArticulo;
		iSaldo.size= 4;
		if (isFirefox) {
			iSaldo.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");					
		} else {
			iSaldo['onkeypress'] = new Function("return controlIngreso(this, event, 'N')");			
		}		
		iSaldo.id = ITEM_SALDO + myRowArticulo;				
		cImgAddArt.appendChild(iSaldo);		
		iSaldo.type = "hidden";
		dSaldo.style.display = 'none';		
		//SALDO TEXT
		dSaldo.id = ITEM_SALDO_TEXT + myRowArticulo;
		iSaldo.style.textAlign = "right";		
		cImgAddArt.appendChild(dSaldo);
		
		myAutoCompletesIndexs[ITEM_DESC + myRowArticulo  + "_text"] = myRowArticulo
		    var link = "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen=" + document.getElementById("cmbIdAlmacen").value;
			$( "#"+ ITEM_DESC + myRowArticulo  + "_text").autocomplete({
				minLength: 2,				
				source: link,
				focus: function( event, ui ) {
					$( "#"+ ITEM_DESC + myAutoCompletesIndexs[this.id]  + "_text").val(ui.item.dsarticulo);
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
		ms.push(myRowArticulo)
		if(pNumTable==0) document.getElementById("artDefault").value = articuloCounter[pNumTable];
		articuloCounter[pNumTable]++;
		document.getElementById("rowTblArticulos_" + pNumTable).value = articuloCounter[pNumTable];
	}
	
	
	/*	function crearTablaArticulo() : crea una linea de la tabla Detalle y le agrega una tabla de articulos.
									Allí van a estar los articulos de cada Partida/sector  */
	function crearTablaArticulo(pRowDetalle){
		var tblDetalle = document.getElementById("tblDetalle");
		var detailRow = tblDetalle.insertRow(parseInt(rowNum));
		detailRow.id= "filaA_" + rowNum		
		var detailCell = detailRow.insertCell(0);
		detailCell.setAttribute("colspan","3");
		detailCell.setAttribute("align","center");		
		// creo por cada Seccion Partida/Sector una tabla donde estaran sus articulos (tblArticulo_X)
		var tTable = document.createElement('table');
		tTable.id   = "tblArticulo_"  + pRowDetalle;
		tTable.name = "tblArticulo_"  + pRowDetalle;
		tTable.setAttribute("width","80%");
		detailCell.appendChild(tTable);
		var rTable = document.createElement('input');
		rTable.type = "hidden";
		rTable.id   = "rowTblArticulos_"  + pRowDetalle;
		rTable.name = "rowTblArticulos_"  + pRowDetalle;		
		detailCell.appendChild(rTable);		
		agregarTituloArticulo(pRowDetalle)		
		rowNum++;
		document.getElementById("rowNum").value = rowNum
	}	
	
	/* function agregarTituloArticulo : agrega el titulo de la tabla de los articulos */
	function agregarTituloArticulo(pRowDetalle){
		var tblArticulo  = document.getElementById("tblArticulo_"  + pRowDetalle);
		var rArticulo	 = tblArticulo.insertRow(articuloCounter[pRowDetalle]);
		rArticulo.setAttribute('class',"reg_Header_nav")
		var cCodigo		 = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cInterno	 = rArticulo.insertCell(2);
		var cStock		 = rArticulo.insertCell(3);
		var cPedido		 = rArticulo.insertCell(4);
		var cImgAddArt	 = rArticulo.insertCell(5);		

		var iCodigo = document.createElement('span');
		iCodigo.id = "cdArticulo_titulo_" + pRowDetalle;		
		cCodigo.appendChild(iCodigo);
		cCodigo.align = "center";
		document.getElementById("cdArticulo_titulo_" + pRowDetalle).innerHTML = "Codigo";
		var iDescripcion = document.createElement('span');
		iDescripcion.id   = "description_titulo_" + pRowDetalle;		
		cDescripcion.appendChild(iDescripcion);
		cDescripcion.align = "center";
		document.getElementById("description_titulo_" + pRowDetalle).innerHTML = "Descripcion";
		var iCdInterno = document.createElement('span');
		iCdInterno.id   = "cdInterno_" + pRowDetalle;
		cInterno.appendChild(iCdInterno);
		cInterno.align = "center";
		document.getElementById("cdInterno_" + pRowDetalle).innerHTML = "Cod.Interno";
		var iStock = document.createElement('span');
		iStock.id   = "stock_span_" + pRowDetalle;
		cStock.appendChild(iStock);
		cStock.align = "center";
		document.getElementById("stock_span_" + pRowDetalle).innerHTML = "Stock";
		var iCantidad = document.createElement('span');
		iCantidad.id   = "pedido_" + pRowDetalle;
		cPedido.appendChild(iCantidad);
		cPedido.align = "center";
		document.getElementById("pedido_" + pRowDetalle).innerHTML = "Pedidos";
		articuloCounter[pRowDetalle]++;
		document.getElementById("rowTblArticulos_" + pRowDetalle).value = articuloCounter[pRowDetalle];
	}
		
	/* function readBudgetArea(me) : Esta funcion asigna a un elemento Hidden el valor del Area selccionada, se dispara cuando
								el combo box del Area/Detalle pierde el focus */
	function readBudgetArea(me){		
		document.getElementById(me.parentNode.id).nextSibling.value = $("#" + me.id + " option:selected").attr("alt");;		
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
	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}
	
	function canSubmit(acc) {
		submitInfo(acc);
		
	}
	
	
	function cerrar() {	
		var refPopUpArt;
		refPopUpArt = startIWin('popupArt');
		refPopUpArt.hide();
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
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Almacen - Pedido de Materiales") %>
<div id="toolbar"></div>
<br>
<form id="frmSel" name="frmSel" method="POST">
	<table class="reg_Header" align="center" width="95%" border="0">
		<tr>
			<td colspan="5">
				<%call showErrors()%>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Almacen") %></td>
			<td width="35%">				
				<select id="cmbIdAlmacen" name="cmbIdAlmacen" onchange="javascript:resetForm(this);">					
				<%	while (not rsAlmacenes.eof)	%>
						<option value="<% =rsAlmacenes("IDALMACEN") %>" alt="<% =rsAlmacenes("IDDIVISION") %>" <%if(rsAlmacenes("IDALMACEN")=PM_idAlmacen)then response.write "selected='true'" end if%>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %>
				<%		rsAlmacenes.MoveNext()
					wend %>
				</select>
				<input type="hidden" name="idDivision" id="idDivision" value="<% =myDivision %>">
			</td>
			<% if flagTransferencia then %>
			<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Transferir desde:") %>&nbsp&nbsp
				<input style="cursor:pointer;" onclick="javascript:resetForm(this);" type="checkBox" value="1" <%if (PM_Transferir = 1) then %>checked<%end if%> name="chkTransferir">
			</td>			
			<td colspan="2" width="35%">
				<div id="almTransferencia" <% if PM_Transferir = 0 then Response.Write " style='display: none;'" else Response.Write " style='display: block;'" end if %> >
					<table border=0 cellspacing=1 cellpadding=1 width=100%>
						<tr>
							<td>
							<%	Set rsAlmacenesDest = obtenerListaAlmacenesSolicitud()%>
								<select id="idAlmacenDest" name="idAlmacenDest">								
									<option value="0">- Seleccione -</option>
								<%	while (not rsAlmacenesDest.eof)	 %>								
										<option value="<% =rsAlmacenesDest("IDALMACEN") %>" <% if (rsAlmacenesDest("IDALMACEN") = PM_idAlmacenDest) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenesDest("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenesDest("DSALMACEN")) %></option>
								<%		rsAlmacenesDest.MoveNext()
									wend   %>
								</select>	
							</td>
						</tr>
					</table>
				</div>
			</td>
			<% end if %>
		</tr>			
		<tr>
			<td class="reg_Header_navdos" ><% =GF_TRADUCIR("Fecha del Vale")%></td>
			<td >
				<div id="issuedateDiv"><% =PM_FechaSolicitud %></div>															
				<input type="hidden" id="issuedate" name="issuedate" value="<% =PM_FechaSolicitud %>"/>
			</td>
			<td class="reg_Header_navdos" ><% =GF_TRADUCIR("Fecha Requerido") %></td>			
			<td align="center" width="10%">
				<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/DATE.gif"></a>
			</td>	
			<td width="25%" align="left">				
				<div id="closingdateDiv"><% =PM_FechaRequerido %></div>				
				<input type="hidden" id="closingdate" name="closingdate" value="<% =PM_FechaRequerido %>" />
			</td>
		</tr>
		<tr>		
			<td class="reg_Header_navdos"><% =GF_TRADUCIR("Solicitante") %></td>
			<td>
				<div id="divSolicitante"></div>
				<input type="hidden" id="cdSolicitante" name="cdSolicitante" value="<% =PM_cdSolicitante %>"/>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" colspan="5"><% =GF_TRADUCIR("Comentario") %></td>
		</tr>
		<tr>
			<td colspan="5" align=center>
				<textarea name="comentario" id="comentario" cols="100"><%=PM_comentario%></textarea>
			</td>
		</tr>
		<tr>
			<td class="reg_Header_nav" colspan="5"><% =GF_TRADUCIR("Detalle") %></td>
		</tr>
		<tr>
			<td colspan="5">
				<table class="reg_Header" width="100%" border=0 id="tblDetalle">
					<tr>
						<td colspan="3"></td>
					</tr>
					<tr>
						<td colspan="3" align="center">							
							<a href="javascript:addLinea();" id="linkAddLinea" style="display:'block';">Agregar Partida/Sector</a>
						</td>                        
					</tr> 
				</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id="accion" name="accion" value="">	
	<input type="hidden" id="TC" name="TC" value="<% =fromTC %>">
	<input type="hidden" id="cantArticulos" name="cantArticulos"  value="0">	
	<input type="hidden" id="cdVale" name="cdVale" value="<%= cdVale%>">
	<input type="hidden" id="rowNum" name="rowNum" value="<%= rowNum%>">
	<input type="hidden" id="rowDetalle" name="rowDetalle" value="<%= rowDetalle%>">
	<input type="hidden" id="rowArticulo" name="rowArticulo" value="<%= rowArticulo%>">	
	<input type="hidden" id="artDefault" name="artDefault" value="<%= artDefault%>">		
	<input type="hidden" id="AlmacenOld" name="AlmacenOld" value="<%= almacenOld%>">		
	<div id="divMastesSelectOld" name="divMastesSelectOld" style="display:none;"><div>	
	<div id="divAreaDetalletOld" name="divAreaDetalletOld" style="display:none;"><div>	
</form>	
		
</body>
</html>
