<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)

Function actualizarElemento(seccion, id, estado)
	Dim strSQL, rs, conn
	
	actualizarElemento = false
	Select Case(seccion) 		
		Case 1: 
			strSQL="UPDATE TBLARTCATEGORIAS SET ESTADO=" & estado & " where IDCATEGORIA=" & id		
		Case 2: 
			strSQL="UPDATE TBLUNIDADES SET ESTADO=" & estado & " where IDUNIDAD=" & id		
		Case 3: 
			strSQL="UPDATE TBLARTICULOS SET ESTADO=" & estado & " where IDARTICULO=" & id		
		Case 5: 
			strSQL="UPDATE TBLBUDGETAREAS SET IDESTADO=" & estado & " where IDAREA=" & id		
		Case 6: 
			strSQL="UPDATE TBLBUDGETDETALLES SET IDESTADO=" & estado & " where IDDETALLE=" & id		
	End Select
	if (strSQL <> "") then
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		actualizarElemento=true
	end if
End Function
'-----------------------------------------------------------------------------------------------
Dim classResponsables, classCategorias, classUnidades, classArticulos, classNormas, classPresupuestos
Dim seccion, classEmpresas, idElemento, id, myEstado, titulo, imagen

seccion = GF_PARAMETROS7("seccion",0,6)
accion = GF_PARAMETROS7("accion","",6)
idElemento = GF_PARAMETROS7("id",0,6)

if (accion = ACCION_BORRAR or accion = ACCION_ACTIVAR) then 
	if accion = ACCION_BORRAR then 
		myEstado = ESTADO_BAJA
	else
		myEstado = ESTADO_ACTIVO
	end if	
	if (actualizarElemento(seccion, idElemento, myEstado)) then
		Response.Write "OK"
	else
		Response.Write "ERROR - No se ejecuto la operacion: Seccion Incorrecta(" & seccion & ")"
	end if
	Response.end
end if

classResponsables = "tabbertab"
classCategorias = "tabbertab"
classUnidades = "tabbertab"
classArticulos = "tabbertab"
'classEmpresas = "tabbertab"
classPresupuestos = "tabbertab"

Select Case(seccion) 
	Case 0: classResponsables = classResponsables & " tabbertabdefault"
	Case 1: classCategorias = classCategorias & " tabbertabdefault"
	Case 2: classUnidades = classUnidades & " tabbertabdefault"	
	Case 3: classArticulos = classArticulos & " tabbertabdefault"
	'Case 4: classEmpresas = classEmpresas & " tabbertabdefault"	
	Case 5: classPresupuestos = classPresupuestos & " tabbertabdefault"		
End Select

%>
<html>
<head>
<title>Sistema de Compras</title>
<link rel="stylesheet" href="css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="css/tabs-print.css" TYPE="text/css" MEDIA="print">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>

<script type="text/javascript">
		
	/* Barra de herramientas de almacenes */
	var allButtonId= new Array(6);
	var toolBarResponsables = new Toolbar("toolBarResponsables", 8, "images/compras/");
	toolBarResponsables.addButtonREFRESH("Refrescar", "obtenerSecciones('0')");
	//toolBarResponsables.addSwitcher("Ver.gif", "Todos", "obtenerSecciones('0', 'todos=1')", "obtenerSecciones('0', '')");
	toolBarResponsables.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(0)", "buscarOff(0)");
	/* Barra de herramientas de categorias */
	var toolBarCategorias = new Toolbar("toolBarCategorias", 8, "images/compras/");
	<% if puedeCrear then %>
		toolBarCategorias.addButton("categories_new-16x16.png", "Nueva", "loadPopUpCategorias(0)");	
	<% end if %>
	toolBarCategorias.addButtonREFRESH("Refrescar", "obtenerSecciones('1')");
	allButtonId[1] = toolBarCategorias.addSwitcher("see_all-16x16.png", "Todos", "obtenerSecciones('1', obtenerParametrosBusquedas(1,1))", "obtenerSecciones('1', obtenerParametrosBusquedas(1,0))");
	toolBarCategorias.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(1)", "buscarOff(1)");
	/* Barra de herramientas de unidades */
	var toolBarUnidades = new Toolbar("toolBarUnidades", 8, "images/compras/");
	<% if puedeCrear then %>
		toolBarUnidades.addButton("units_new-16x16.png", "Nueva", "loadPopUpUnidades(0)");
	<% end if %>	
	toolBarUnidades.addButtonREFRESH("Refrescar", "obtenerSecciones('2')");
	allButtonId[2] = toolBarUnidades.addSwitcher("see_all-16x16.png", "Todos", "obtenerSecciones('2', obtenerParametrosBusquedas(2,1))", "obtenerSecciones('2', obtenerParametrosBusquedas(2,0))");	
	toolBarUnidades.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(2)", "buscarOff(2)");
	/* Barra de herramientas de articulos */
	var toolBarArticulos = new Toolbar("toolBarArticulos", 8, "images/compras/");
	<% if puedeCrear then %>
		toolBarArticulos.addButton("items_new-16x16.png", "Nuevo", "loadPopUpArticulos(0)");
	<% end if %>	
	toolBarArticulos.addButtonREFRESH("Refrescar", "obtenerSecciones('3')");
	allButtonId[3] =  toolBarArticulos.addSwitcher("see_all-16x16.png", "Todos", "obtenerSecciones('3', obtenerParametrosBusquedas(3,1))", "obtenerSecciones('3', obtenerParametrosBusquedas(3,0))");
	toolBarArticulos.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(3)", "buscarOff(3)");
	toolBarArticulos.addButton("changeUnit-16x16.png", "Cambiar Unidad", "loadPopUpCambioUnidad()");
	/* Barra de herramientas de Empresas */		
	//var toolBarEmpresas = new Toolbar("toolBarEmpresas", 8, "images/compras/");
	<% if puedeCrear then %>
	//	toolBarEmpresas.addButton("company_new-16x16.png", "Nuevo", "loadPopUpEmpresas(0)");
	<% end if %>
	//toolBarEmpresas.addButtonREFRESH("Refrescar", "obtenerSecciones('4')");		
	//toolBarEmpresas.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(4)", "buscarOff(4)");		

	/* Barra de herramientas de presupuestos */
	var toolBarPresupuestos = new Toolbar("toolBarPresupuestos", 6, "images/compras/");
	var swt, btn;
	<% if puedeCrear then %>
		toolBarPresupuestos.addButton("Budget_Area_new-16x16.png", "Nueva Area", "loadPopUpPresupuestos(0,'A')");
		toolBarPresupuestos.addButton("Budget_Item_new-16x16.png", "Nuevo Detalle", "loadPopUpPresupuestos(0,'D')");
	<% end if %>
	toolBarPresupuestos.addButtonREFRESH("Refrescar", "obtenerSecciones('5')");
	toolBarPresupuestos.addButton("Budget_Area-16x16.png", "Areas", "obtenerSecciones('5', '', 1);");
	toolBarPresupuestos.addButton("Budget_Item-16x16.png", "Items", "obtenerSecciones('5', 'detalles=1', 0)");		
	allButtonId[5] = toolBarPresupuestos.addSwitcher("see_all-16x16.png", "Todos", "obtenerSeccionesPre('5', obtenerParametrosBusquedas(5,1), 1, swt)", "obtenerSeccionesPre('5', obtenerParametrosBusquedas(5,0), 0, swt)");	
	toolBarPresupuestos.addSwitcher("Search-16x16.png", "Buscar", "buscarOn(5)", "buscarOff(5)");


	/* Funciones de busqueda */
	function buscarOn(id) {
		document.getElementById("busqueda" + id).className = "";		
	}
	
	function buscarOff(id) {
		document.getElementById("busqueda" + id).className = "divOculto";		
	}
	
	function doBuscar(seccion) {
		//alert("aa");
		switch (seccion) {
			case 0:
				buscarResponsables();
				break;
			case 1:
				buscarCategorias(1);
				break;
			case 2:
				buscarUnidades(2);
				break;
			case 3:
				buscarArticulos(3);
				break;
			//case 4:
			//	buscarEmpresas();
			//	break;
			case 5:
				buscarPresupuestos(5);
				break;
		}		
	}
	function obtenerParametrosBusquedas(idSeccion,todos){
		var param="";
		if ( idSeccion == 0 && document.getElementById("busqueda0").className != "divOculto" ) {
			param = "&cdResponsable=" + document.getElementById("cdResponsable").value;
			param += "&dsResponsable=" + document.getElementById("dsResponsable").value;
			param += "&hkResponsable=" + document.getElementById("hkResponsable").value;
		}
		if ( idSeccion == 1 && document.getElementById("busqueda1").className != "divOculto" ) {
			param = "&cdCategoria=" + document.getElementById("cdCategoria").value;
			param += "&dsCategoria=" + document.getElementById("dsCategoria").value;		
		}
		if ( idSeccion == 2 && document.getElementById("busqueda2").className != "divOculto" ) {
			param = "&cdUnidad=" + document.getElementById("cdUnidad").value;
			param += "&dsUnidad=" + document.getElementById("dsUnidad").value;		
		}
		if ( idSeccion == 3 && document.getElementById("busqueda3").className != "divOculto" ) {
			param = "&idArticulo=" + document.getElementById("idArticulo").value;
			param += "&dsArticulo=" + document.getElementById("dsArticulo").value;
		}
		/*if ( idSeccion == 4 && document.getElementById("busqueda4").className != "divOculto" ) {
			param = "&idEmpresa=" + document.getElementById("idEmpresa").value;
			param += "&dsEmpresa=" + document.getElementById("dsEmpresa").value;
			param += "&cuit=" + document.getElementById("cuit").value;
			param += "&listaNegra=" + document.getElementById("listaNegra").value;
		}*/
		if ( idSeccion == 5 && document.getElementById("busqueda5").className != "divOculto" ) {
			if(document.getElementById("dsPresupuestoDetalle")){
				param = "&detalles=1&dsPresupuestoDetalle=" + document.getElementById("dsPresupuestoDetalle").value;
			}
			else{
				param = "&dsPresupuestoArea=" + document.getElementById("dsPresupuestoArea").value;
			}	
		}
		return "todos=" + todos + param ;
	}
	function buscarResponsables() {
		var param = "cdResponsable=" + document.getElementById("cdResponsable").value;
		param += "&dsResponsable=" + document.getElementById("dsResponsable").value;
		param += "&hkResponsable=" + document.getElementById("hkResponsable").value;
		if (document.getElementById("verEmpleados1").checked) param += "&verEmpleados=" + document.getElementById("verEmpleados1").value;
		if (document.getElementById("verEmpleados2").checked) param += "&verEmpleados=" + document.getElementById("verEmpleados2").value;
		obtenerSecciones('0', param);
	}
	function buscarCategorias(idSeccion) {
		var param = "cdCategoria=" + document.getElementById("cdCategoria").value;
		param += "&dsCategoria=" + document.getElementById("dsCategoria").value;		
		param += "&todos=" + toolBarCategorias.buttons[allButtonId[idSeccion]].status ;
		obtenerSecciones('1', param);
	}
	function buscarUnidades(idSeccion) {
		var param = "cdUnidad=" + document.getElementById("cdUnidad").value;
		param += "&dsUnidad=" + document.getElementById("dsUnidad").value;		
		param += "&todos=" + toolBarUnidades.buttons[allButtonId[idSeccion]].status ;
		obtenerSecciones('2', param);
	}
	function buscarArticulos(idSeccion) {
		var param = "idArticulo=" + document.getElementById("idArticulo").value;
		param += "&dsArticulo=" + document.getElementById("dsArticulo").value;
		param += "&cdArtCategoria=" + document.getElementById("cdArtCategoria").value;
		param += "&todos=" + toolBarArticulos.buttons[allButtonId[idSeccion]].status ;
		obtenerSecciones('3', param);
	}
	/*function buscarEmpresas() {
		var param = "idEmpresa=" + document.getElementById("idEmpresa").value;
		param += "&dsEmpresa=" + document.getElementById("dsEmpresa").value;
		param += "&cuit=" + document.getElementById("cuit").value;
		param += "&listaNegra=" + document.getElementById("listaNegra").checked;
		obtenerSecciones('4', param);
	}*/
	function buscarPresupuestos(idSeccion) {
		var param;
		if(document.getElementById("dsPresupuestoDetalle")){
			param = "detalles=1&dsPresupuestoDetalle=" + document.getElementById("dsPresupuestoDetalle").value;
		}
		else{
			param = "dsPresupuestoArea=" + document.getElementById("dsPresupuestoArea").value;
		}	
		param += "&todos=" + toolBarPresupuestos.buttons[allButtonId[idSeccion]].status ;
		obtenerSecciones('5', param);
	}	
	/* Manejo de PopUps */		
	function loadPopUpResponsablesApertura(id) {				
		var puw = new winPopUp('popupResponsable','comprasPropResponsable.asp?idResponsable=' + id,'400','420','Propiedades Responsable Apertura', 'obtenerSecciones("0")');
	}
	
	function loadPopUpResponsablesAccesos(cd) {
		var puw = new winPopUp('popupResponsableAccesos','comprasPropResponsableAccesos.asp?cdResponsable=' + cd,'500','400','Propiedades Responsable Accesos', 'obtenerSecciones("0")');
	}

    function loadPopUpResponsablesRoles(cd) {
    	var puw = new winPopUp('popupResponsableRoles','comprasPropResponsableRoles.asp?cdResponsable=' + cd,'600','600','Propiedades Responsable Roles', 'obtenerSecciones("0")');
    }
    
	function loadPopUpCambioUnidad()
	{
		winPopUp('Cambiar Unidad', 'almacenCambioUnidad.asp', '600', '500', 'Cambiar Unidad', '');
	}

	function reloadPage(){
		document.location.reload();
	}

	function loadPopUpCategorias(id) {				
		var puw = new winPopUp('popupCategoria','comprasPropCategoria.asp?idCategoria=' + id,'400','300','Propiedades Categoria', 'obtenerSecciones("1")');
	}		
	
	function loadPopUpUnidades(id) {				
		var puw = new winPopUp('popupUnidad','comprasPropUnidad.asp?idUnidad=' + id,'420','350','Propiedades Unidad', 'obtenerSecciones("2")');
	}
	function loadPopUpPresupuestos(id, tipo) {			
		var aux = "";
		var puw = new winPopUp('popupPresupuesto','comprasPropPresupuesto.asp?idItem=' + id + "&tipo=" + tipo,'320','150','Propiedades Presupuesto', 'obtenerSecciones("5")');
	}		
	function loadPopUpArticulos(id) {
		var puw = new winPopUp('popupArticulo','comprasPropArticulo.asp?idArticulo=' + id, '500', '300','Propiedades Articulo', 'obtenerSecciones("3")');
	}
	
	/*function loadPopUpEmpresas(id) {	
		var puw = new winPopUp('popupEmpresa','comprasPropEmpresa.asp?idEmpresa=' + id,'360','350','Propiedades Empresa', 'obtenerSecciones("4")');	
	}*/
	
	/* Codigo para el manejo de las secciones */
	var ch = new channel();	
	var vParams = new Array();	//Vector con los ultimos parametros utilizados para una seccion.		
	var vRegs = new Array();	//Vector con los objetos de paginacion de cada seccion.	
	
	function addParam(seccion, name, value) {			
		if (vParams[seccion]) {
			var arrKV = vParams[seccion].split("&");		
			var pos = vParams[seccion].indexOf(name);
		} else {
			var arrKV = new Array();
			var pos = -1;
		}		
		if (pos > -1) {
			//El parametro existe, se reemplaza.
			for(var k in arrKV) {
				var u = arrKV[k].indexOf(name);
				if (u > -1) arrKV[k] = name + "=" + value;
			} 
		} else {			
			arrKV.push(name + "=" + value);
		}
		vParams[seccion] = (arrKV.toString()).replace(/,/g, "&");
	}
	 
	function seccionCallback(index) {		
		var div = document.getElementById("seccion" + index);
		var divB = document.getElementById("busqueda" + index + "TD");
		//recibo los datos del canal			
		var chData = ch.response();
		var arr = chData.split("-h-");
		if (arr[1].indexOf("-#-") != -1) {
			var arr2 = arr[1].split("-#-");
			divB.innerHTML = arr2[0];
			div.innerHTML = arr2[1];			
		} else {
			div.innerHTML = arr[1];
		}				
		//Verifico si hay que paginar		
		if (vRegs[index]) {			
			paginarSeccion(index, arr[0]);			
		}		
	}	
	 
	function paginarSeccion(seccion, cantLineas) {
		if (vRegs[seccion].cantidadLineas != cantLineas) {
			if (cantLineas==0) cantLineas=1;
			vRegs[seccion].paginar(1, cantLineas, 10, 50, "paginarCall(" + seccion + ")");					
		}
	}
	
	function paginarCall(seccion, pagina, regsXPag) {
		//Remuevo los parametros de paginacion del cache de parametros, esto evita que se dupliquen al cambiar de pagina. 
		addParam(seccion, PGN_ACTUAL_PAGE, pagina);
		addParam(seccion, PGN_LINES_X_PAGE, regsXPag);
		obtenerSecciones(seccion);
	}
	function obtenerSeccionesPre(seccion, params, state, button) {
		var aux = "";
		if (toolBarPresupuestos.buttons[button].status == 1){
			aux = "&detalles=1";
		}
		else{
			aux = "&detalles=0";
		}
		
		obtenerSecciones(seccion, params + aux);
	}
	
	function obtenerSecciones(seccion, params) {
		var d = new Date();
		var p = "";
		if ((params) || (params == "")) vParams[seccion] = params;
		if (vParams[seccion]) p = "&" + vParams[seccion];		
		//Se solicita la seccion
		seccionCache[seccion] = d.getMinutes();
		ch.bind("comprasArmarSecciones.asp?seccion=" + seccion + p, "seccionCallback(\"" + seccion + "\")");
		ch.send();	
	}
	
	function altaBajaCallback(seccion) {		
		var errMsg = ch.response();
		if (errMsg != "OK") {			
			alert(errMsg);
		} else {
			obtenerSecciones(seccion);
		} 		
	}
	
	function deleteElemento(seccion, id) {
		if (confirm("Esta seguro que desea eliminar el elemento seleccionado?")) {
			//Al borrar puede cambiar la cantidad de paginas, se reincian los parametros.
			addParam(seccion, PGN_ACTUAL_PAGE, 1);
			addParam(seccion, PGN_LINES_X_PAGE, 10);
			var link = "comprasSecciones.asp?accion=<% =ACCION_BORRAR %>"; 
			link += "&seccion=" + seccion;
			link += "&id=" + id;
			if (seccion==6) seccion = 5;	
			ch.bind(link, "altaBajaCallback(\"" + seccion + "\")");
			ch.send();
		}
	}
		
	function habilitarElemento(seccion, id, estado) {
		var link = "comprasSecciones.asp?accion=<%= ACCION_ACTIVAR %>"; 
		link += "&seccion=" + seccion;
		link += "&id=" + id;
		if (seccion==6) seccion = 5;			
		ch.bind(link, "altaBajaCallback(\"" + seccion + "\")");
		ch.send();
	}
	function irObras() {
		location.href = "comprasObras.asp";
	}
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	
	function loadFunc() {
 		toolBarResponsables.draw();
		toolBarCategorias.draw();
		toolBarUnidades.draw();
		toolBarArticulos.draw();		
		//toolBarEmpresas.draw();
		toolBarPresupuestos.draw();
		//Habilito las secciones que quiero paginar.
		vRegs[0] = new Paginacion("paginacion0");	//Responsables.
		vRegs[1] = new Paginacion("paginacion1");	//Categorias.
		vRegs[2] = new Paginacion("paginacion2");	//Unidades.
		vRegs[3] = new Paginacion("paginacion3");	//Articulos.
		//vRegs[4] = new Paginacion("paginacion4");	//Empresas.
		vRegs[5] = new Paginacion("paginacion5");	//Presupuestos.
		
		var tb = new Toolbar('toolbar', 6, "images/compras/");	
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButton("OBR-16X16.png", "Obras", "irObras()");				
		tb.addButton("Bid_purchase-16x16.png", "Concurso", "irConcurso()");
		tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irComparativa()");
		tb.addButton("Direct_purchase-16x16.png", "Directa", "irDirecta()");
		tb.draw();
		obtenerSecciones("<% =seccion %>");
	}

	function irDirecta(){
		location.href = "comprasAdministrarCotizaciones.asp";
	}
	function irComparativa() {
		location.href = "comprasAdministrarPedidos.asp?tipoCompra=P";
	}	
	function irConcurso() {
		location.href = "comprasAdministrarPedidos.asp?tipoCompra=C";
	}	
	
	/* Esto se hace asi ya que lo necesita el tabber */
	//Vector para determinar el momento de carga de una seccion, si pasaron mas de 5 minutos, la carga de nuevo.
	var seccionCache = new Array();	
	var tabberOptions = {'onClick':	function(argsObj) {    					
										var load = true;
										//Se verifica si la seccion ya fue cargada.													
										if (seccionCache[argsObj.index] != undefined) {
											//Si la seccion ya fue cargada, no la recarga hasta que pasen por lo menos 5 minutos.
											var d = new Date();
											var diff = d.getMinutes() - seccionCache[argsObj.index];  
											if ((diff < 5) && (diff >= 0)) load = false;											
										}	
										if (load) obtenerSecciones(argsObj.index);		
  									}};
	window.onload = loadFunc;
</script>
<script language="javascript" src="scripts/tabber.js"></script>
</head>
<body>
<% call GF_TITULO2("kogge64.gif","Administración") %>
<div id="toolbar"></div><br>
<div class="tabber">
    <div class="<%= classResponsables %>">
		<%  seccion= "0"
		imagen = "users-16x16.png"
		titulo = "Responsables"	%>
		<!--#include file="comprasSeccionDefault.asp"-->    	
	</div>     
    <div class="<%= classCategorias %>">
		<%  seccion= "1"
		imagen = "categories-16x16.png"
		titulo = "Categorias"	%>
		<!--#include file="comprasSeccionDefault.asp"-->    	
	</div>  
    <div class="<%= classUnidades %>">
		<%  seccion= "2"
		imagen = "units-16x16.png"
		titulo = "Unidades"	%>
		<!--#include file="comprasSeccionDefault.asp"-->    	
	</div> 
    <div class="<%= classArticulos %>">	  	
		<%  seccion= "3"
		imagen = "items-16x16.png"
		titulo = "Articulos"	%>		
		<!--#include file="comprasSeccionDefault.asp"-->    	
	</div> 	
	<!--
	<div class="<%= classEmpresas %>">	  	
		<%  'seccion= "4"
		'imagen = "Company-16x16.png"
		'titulo = "Empresas"	%>		
	</div> 	-->
	<div class="<%= classPresupuestos %>">	  	
		<%  seccion= "5"
		imagen = "budget_view-16x16.png"
		titulo = "Presupuestos"	%>		
		<!--#include file="comprasSeccionDefault.asp"-->    	
	</div> 	
</div>
</body>
</html>
