<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
	Dim strSQL,rs,rs2,conn
	Dim idObra,accion
	Dim areaDetalle,controlOK,importeDL,motivo,mostrarImporteMotivo
	Dim obraCD, obraDS, obraDivID, obraDivDS, obraImporte, obraFechaBudget, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS
	Dim importePIC, dsUltimaArea 'dicVales,
	dim gastosFacturados, totalVales, totalPedido, totalFacturado, totalBudget
	dim totalGralVales, totalGralPedido, totalGralFacturado, totalGralBudget
	dim idAreaActual, dsAreaActual,dicGrabarPartidas,g_idArea,g_detalleNew,tipoFormularioPartida
	
Function grabarNuevasCuentasCCosto(idObra)
	Dim areaDetalle,mySets,rs2
	strSQL="SELECT * FROM TBLBUDGETOBRAS WHERE IDOBRA=" & idObra & " ORDER BY IDAREA, IDDETALLE"
	Call executeQueryDb(DBSITE_SQL_INTRA, rsInt, "OPEN", strSQL)
	while not rsInt.EOF
		areaDetalle	 = rsInt("IDAREA") & rsInt("IDDETALLE")
		mySets = ""
		if (rsInt("IDDETALLE") <> 0 ) then
			IF ( GF_Parametros7("cuenta_"&areaDetalle,"",6) <> "") then
				mySets = mySets  & "CDCUENTA = '" & GF_Parametros7("cuenta_"&areaDetalle,"",6) & "',"
			end if
			if (GF_Parametros7("ccosto_"&areaDetalle,0,6) <> 0) then
				mySets = mySets  & "CCOSTOS = " & GF_Parametros7("ccosto_"&areaDetalle,"",6) & ","
			end if
			if (mySets <> "") then
				mySets = left(mySets,len(mySets)-1) 'le quito la ultima coma
				strSQL = "UPDATE TBLBUDGETOBRAS SET " & mySets & " WHERE idobra = " & idObra & " and idarea = " & rsInt("IDAREA") & " and iddetalle = " & rsInt("IDDETALLE")
				'Response.Write strSQL
				Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "UPDATE", strSQL)
			end if
		end if
		rsInt.MoveNext
	wend
	'Response.End 
end Function
'---------------------------------------------------------------------------------------
'Controla las partidas nuevas que el usuario agrega
Function controlarNuevasPartidas(idObra)
	Dim dsDetallePartida,nuevoId,dicDsDetalle,rsInt
	Set dicDsDetalle = Server.CreateObject("Scripting.Dictionary")
	nuevoId = getNuevoIdPartidaDetalle(idObra)
	strSQL="SELECT DISTINCT(IDAREA) AS IDAREA FROM TBLBUDGETOBRAS WHERE IDOBRA=" & idObra & " ORDER BY IDAREA"
    Call executeQueryDb(DBSITE_SQL_INTRA, rsInt, "OPEN", strSQL)
	while not rsInt.EOF		
		i  = 1
		'Obtengo la cantidad de filas que se agregaron por Area de partida
		indexDetalle = Cdbl(GF_Parametros7("rowDetalle_"& rsInt("IDAREA"),0,6))
		while (i <= indexDetalle)and(not hayError())
			'obtengo y valido la descripcion del detalle de la partida
			dsDetallePartida = Ucase(Trim(GF_Parametros7("textDsDetalle_"& rsInt("IDAREA") & "_" & i,"",6)))
			if (dsDetallePartida <> "") then
				'consulto si la descripcion ya fue controlada en Areas anteriores(para no generar el nuevo Id)
				if (not dicDsDetalle.Exists(dsDetallePartida)) then
					nuevoId = nuevoId + 1
					dicDsDetalle.Add dsDetallePartida,rsInt("IDAREA") &"|"& nuevoId
				else
					'Si la descripcion del Detalle ya fue cargada, verifico que no se encuentre repetida para el AREA que se esta analizando
					auxItem = split(dicDsDetalle.item(dsDetallePartida),"|")
					if (not dicGrabarPartidas.Exists(rsInt("IDAREA") &"|"& auxItem(1) &"|"& dsDetallePartida)) then
						nuevoId = Cdbl(auxItem(1))
					else
						Call setError(DETALLE_DUPLICADO)
					end if
				end if
				if not hayError() then dicGrabarPartidas.Add rsInt("IDAREA") &"|"& nuevoId &"|"& dsDetallePartida,""
			else
				Call setError(DESC_BUDGET_OBLIGATORIA)
			end if
			i = i + 1
		wend
		rsInt.MoveNext()
	wend
End Function
'---------------------------------------------------------------------------------------
'Encargada de devolver el ultimo Id del budget
Function getNuevoIdPartidaDetalle(idObra)
	Dim nuevoId,strSQL,rsDetalle,rsArea
	nuevoId=0
	strSQL= "Select max(IDDETALLE) IDDETALLE from TBLBUDGETOBRAS where IDOBRA=" & idObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rsDetalle, "OPEN", strSQL)
	strSQL= "Select max(IDAREA) IDAREA from TBLBUDGETOBRAS where IDOBRA=" & idObra
    Call executeQueryDb(DBSITE_SQL_INTRA, rsArea, "OPEN", strSQL)
	if (not rsDetalle.eof) then
		if (CInt(rsDetalle("IDDETALLE")) > CInt(rsArea("IDAREA"))) then
			nuevoId = CInt(rsDetalle("IDDETALLE"))
		else
			nuevoId = CInt(rsArea("IDAREA"))
		end if
		if (nuevoId < 1000) then nuevoId = 1000
	end if
	getNuevoIdPartidaDetalle = nuevoId
End Function
'------------------------------------------------------------------------------------------------------
Function grabarNuevasPartidas(idObra)
	Dim tipoCambioBgt,rsInsert
	tipoCambioBgt = getTipoCambioBudget(idObra)
	for each key in dicGrabarPartidas.Keys
		auxItem = split(key,"|")
		strSQL="Insert into TBLBUDGETOBRAS values(" & idObra & ", " & auxItem(0) & ", " & auxItem(1) & ", '" & auxItem(2) & "', 0, 0, " & tipoCambioBgt & ", '" & session("Usuario") & "', " & session("MmtoSistema") & ", '', '', 0)"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsInsert, "EXEC", strSQL)
	next
end Function
'------------------------------------------------------------------------------------------------------
'Procesa la nueva linea del Budget
Function readNextAreaPartida()
	readNextAreaPartida = false
	if (not rs.Eof) then
		g_idArea = rs("IDAREA")
		g_detalleNew = 0
		'Si luego de submitir presenta errores el control, debo guardar el numero de filas que se agregaron para cada Area
		if hayError() then g_detalleNew = GF_Parametros7("rowDetalle_"& rs("IDAREA"),0,6)
		readNextAreaPartida = true
	end if	
End function
'----------------------------------------------------------------------------------------------------------
'********************************************************************
'***********************    INICIO DE LA PAGINA    ******************
'********************************************************************
	Call GP_ConfigurarMomentos
	
	idObra = GF_Parametros7("idObra",0 ,6)
	accion = GF_Parametros7("accion","",6)
    tipoFormularioPartida = getFormTypeByIdObra(idObra)
	if (accion=ACCION_GRABAR) then
		Set dicGrabarPartidas = Server.CreateObject("Scripting.Dictionary")
		Call controlarNuevasPartidas(idObra)
		if not hayError() then
			Call grabarNuevasCuentasCCosto(idObra)
			Call grabarNuevasPartidas(idObra)
		end if
	end if
	strSQL="SELECT IDAREA, DSBUDGET FROM TBLBUDGETOBRAS WHERE IDOBRA=" & idObra & " AND IDDETALLE=0 GROUP BY IDAREA, DSBUDGET order by IDAREA"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

Call loadDatosObra(idObra, obraCD, obraDS, obraDivID, obraDivDS, obraImporte, obraFechaBudget, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS)
fechaFin = obraFechaFin
if (CDbl(obraFechaAjustada) <> 0) then fechaFin = obraFechaAjustada			
	
%>
<HTML>
<HEAD>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link type="text/css" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

	<style type="text/css">
		body { font-size: 62.5%; }

		.titulo{
			color:#FFFFFF;
			font-weight:bold;
		}
		.cursorStyle {
			cursor: pointer;
		}		
	</style>
	
<script type="text/javascript">
	var ch = new channel();	
	var lastRow = 3;
	var totalGralVales = 0;
	var totalGralPedidos = 0;
	var totalGralFacturado = 0;
	var totalGralBudget = 0;
	var cantidadRegistros = 0;
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	var lastPositionDetalle = new Array();
	var counterNewDetalle = new Array();	
	function bodyOnLoad(){
		var tb = new Toolbar('toolbar', 6, 'images/compras/');
		tb.addButton("save-16x16.png", "<%=GF_Traducir("Guardar")%>", "grabarReasignacion()");
		tb.addButtonREFRESH("<%=GF_Traducir("Recargar")%>", "refreshPage()");
		tb.addButton("printer-16x16.png", "<%=GF_Traducir("Imprimir")%>", "imprimir()");
		tb.addButton("previous-16x16.png", "<%=GF_Traducir("Volver")%>", "volver()");
		tb.addButton("see_all-16x16.png", "<%=GF_Traducir("Tablero Obra")%>", "abrirTableroObra()");
		tb.draw();
		cargar();
		//LEER LAS AREAS Y PEDIR LOS DETALLES POR AJAX and i<=2 
		<%	while (readNextAreaPartida())%>
				counterNewDetalle[<%=g_idArea%>] = 0;
				cargarDetalles(<%=idObra%>, <%=g_idArea%>, <%=g_detalleNew%>);
				<% for i=1 to g_detalleNew %>
					restablecerDescripcionDetalle(<%=g_idArea%>,<%=i%>,'<%=Ucase(Trim(GF_Parametros7("textDsDetalle_"& g_idArea & "_" & i,"",6)))%>');
				<% next %>
		<%		rs.MoveNext()
			wend %>			
	}
	function cargar() {
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel2").style.visibility = 'visible';		
	}
	
	function ocultar_cargar() {
		document.getElementById("actionLabel").style.visibility = 'hidden';
		document.getElementById("actionLabel2").style.visibility = 'hidden';		
	}
	function cargarDetalles(pIdObra, pIdArea, pCountDet){
		ch.bind("comprasBudgetGetAreaDetalle_AJAX.asp?idObra=" + pIdObra + "&idBudgetArea=" + pIdArea + "&obraFechaInicio=<%=obraFechaInicio%>&fechaFin=<%=fechaFin%>", "cargarDetalles_Callback(" + pIdArea + "," + pCountDet + ")");
		ch.send();			
	}
	function cargarDetalles_Callback(pIdArea,pCountDet){
		var totalVales = 0;
		var totalPedidos = 0;
		var totalFacturado = 0;
		var totalBudget = 0;
		document.getElementById("image" + pIdArea).style.visibility = "hidden";
		document.getElementById("image" + pIdArea).style.position = "absolute";
		var result = ch.response();
		//alert(result);
		var myDescomposicion;
		var myDescomposicionInterna;
		myDescomposicion = result.split("//");
		lastRow = document.getElementById("AREA_" + pIdArea).rowIndex+1
		lastPositionDetalle[pIdArea] = 0;
		for (var i=1;i<myDescomposicion.length;i++){			
			myDescomposicionInterna = myDescomposicion[i].split(";");
			totalVales = totalVales + parseFloat(myDescomposicionInterna[2]);
			totalPedidos = totalPedidos + parseFloat(myDescomposicionInterna[3]);
			totalFacturado = totalFacturado + parseFloat(myDescomposicionInterna[4]);
			totalBudget = totalBudget + parseFloat(myDescomposicionInterna[5]);
			agregarDetalle(lastRow,pIdArea,myDescomposicionInterna[0], myDescomposicionInterna[1], myDescomposicionInterna[2], myDescomposicionInterna[3], myDescomposicionInterna[4], myDescomposicionInterna[5], myDescomposicionInterna[6], myDescomposicionInterna[7])
			lastRow = lastRow + 1;
		}
		//NUEVO: SE AGREGA LA OPCION PARA QUE SE AGREGE UN DETALLE NUEVO AL AREA DE LA PARTIDA (SIEMPRE LO HAR� CON MONTO 0)
		dibujarItemNuevaPartida(lastRow,pIdArea);
		if (pCountDet != 0){
			for (var index=1;index<=pCountDet;index++){
				var auxDsDetalle = "";
				if (document.getElementById("hiddenDsDetalle_"+pIdArea+"_"+index) != null){
					 auxDsDetalle = document.getElementById("hiddenDsDetalle_"+pIdArea+"_"+index).value;
					 var padre = document.getElementById("frm");
					 padre.removeChild(document.getElementById("hiddenDsDetalle_"+pIdArea+"_"+index));
				}
				agregarDetallePartida(pIdArea,auxDsDetalle);
			}
		}
		document.getElementById("VlVales_" + pIdArea).innerHTML = formatearImporte(totalVales.toString(),2); 
		document.getElementById("VlPedidos_" + pIdArea).innerHTML = formatearImporte(totalPedidos.toString(),2); 
		document.getElementById("VlFacturado_" + pIdArea).innerHTML = formatearImporte(totalFacturado.toString(),2); 
		document.getElementById("VlBudget_" + pIdArea).innerHTML = formatearImporte(totalBudget.toString(),2); 
		totalGralVales = totalGralVales + totalVales;
		totalGralPedidos = totalGralPedidos + totalPedidos;
		totalGralFacturado = totalGralFacturado + totalFacturado;
		totalGralBudget = totalGralBudget + totalBudget;
		//CARGAR TOTALES GENERALES
		document.getElementById("VlGralVales").innerHTML = formatearImporte(totalGralVales.toString(),2); 
		document.getElementById("VlGralPedidos").innerHTML = formatearImporte(totalGralPedidos.toString(),2); 
		document.getElementById("VlGralFacturado").innerHTML = formatearImporte(totalGralFacturado.toString(),2); 
		document.getElementById("VlGralBudget").innerHTML = formatearImporte(totalGralBudget.toString(),2);
		cantidadRegistros = cantidadRegistros + 1;
		if (cantidadRegistros == <%=rs.recordCount%>)
			ocultar_cargar();
	}
	function agregarDetalle(pNbrRow,pBudgetArea,pBudgetDetalle, pBudgetDs, pVlVales, pVlPedidos, pVlFacturado, pVlBudget, pCuenta, pCentroCosto){
		var tableRef = document.getElementById("tableBudget");
		var rowDetalle = tableRef.insertRow(pNbrRow);
		rowDetalle.id = "detalle_" + pBudgetArea + "_" + pBudgetDetalle
		var cellBlank = rowDetalle.insertCell(0);
		var cellIdBudget = rowDetalle.insertCell(1);
		var cellDsBudget = rowDetalle.insertCell(2);
		var cellVlVales = rowDetalle.insertCell(3);
		var cellVlPedidos = rowDetalle.insertCell(4);
		var cellVlFacturado = rowDetalle.insertCell(5);
		var cellVlBudget = rowDetalle.insertCell(6);
		var cellCuenta = rowDetalle.insertCell(7);
		var cellCentroCostos = rowDetalle.insertCell(8);
		var cellLink = rowDetalle.insertCell(9);
		var cellAjuste = rowDetalle.insertCell(10);
		/*General*/
		if (parseFloat(pVlBudget) < parseFloat(pVlPedidos) + parseFloat(pVlVales)){
			//alert(pVlBudget + "-" + pVlPedidos)
			rowDetalle.className = "reg_header_warning";
			}
		if (parseFloat(pVlBudget) == 0) 
			rowDetalle.className = "reg_header_error";
		
		/*Blank*/
		
		/*IdBudget*/
		cellIdBudget.className = "reg_header_nav round_border_left";
		cellIdBudget.width = "2%";
		cellIdBudget.style.textAlign= "right";
		var divIdBudget = document.createElement('div');
		divIdBudget.innerHTML = pBudgetDetalle;
		cellIdBudget.appendChild(divIdBudget);
		/*DsBudget*/
		var divDsBudget = document.createElement('div');
		divDsBudget.innerHTML = pBudgetDs;
		cellDsBudget.appendChild(divDsBudget);	
		/*VlVales*/
		var divVlVales = document.createElement('div');
		divVlVales.style.textAlign= "right";
		divVlVales.innerHTML = formatearImporte(pVlVales,2);
		cellVlVales.appendChild(divVlVales);	
		/*VlPedidos*/
		var divVlPedidos = document.createElement('div');
		divVlPedidos.style.textAlign= "right";
		divVlPedidos.innerHTML = formatearImporte(pVlPedidos,2);
		cellVlPedidos.appendChild(divVlPedidos);			
		/*VlPedidos*/
		var divVlFacturado = document.createElement('div');
		divVlFacturado.style.textAlign= "right";
		divVlFacturado.innerHTML = formatearImporte(pVlFacturado,2);
		cellVlFacturado.appendChild(divVlFacturado);	
		/*VlBudget*/
		var divVlBudget = document.createElement('div');
		divVlBudget.style.textAlign= "right";
		divVlBudget.innerHTML = formatearImporte(pVlBudget,2);
		cellVlBudget.appendChild(divVlBudget);	
		/*Cuenta*/
		cellCuenta.align="center";
		var txtCuenta = document.createElement('input');
		txtCuenta.style.textAlign= "center";
		txtCuenta.size = 12;
		txtCuenta.value = pCuenta;
		txtCuenta.maxLength = 12;
		txtCuenta.name = 'cuenta_' + pBudgetArea + pBudgetDetalle;
		txtCuenta.id = 'cuenta_' + pBudgetArea + pBudgetDetalle;
		//}
		cellCuenta.appendChild(txtCuenta);	
		
		/*Centro de Costos*/
		cellCentroCostos.align="center";
		var txtCCostos = document.createElement('input');
		txtCCostos.style.textAlign= "center";
		txtCCostos.size = 10;
		if (pCentroCosto != '' && pCentroCosto != 0){ 
			var divCentroCostos = document.createElement('div');
			divCentroCostos.style.textAlign= "center";
			divCentroCostos.innerHTML = pCentroCosto;
			cellCentroCostos.appendChild(divCentroCostos);			
			txtCCostos.type = 'hidden';
			txtCCostos.value = '';
			txtCCostos.name = 'ccosto_' + pBudgetArea + pBudgetDetalle;
			txtCCostos.id = 'ccosto_' + pBudgetArea + pBudgetDetalle;
		}
		else{
			txtCCostos.value = '';
			txtCCostos.maxLength = 6;
			txtCCostos.name = 'ccosto_' + pBudgetArea + pBudgetDetalle;
			txtCCostos.id = 'ccosto_' + pBudgetArea + pBudgetDetalle;
		}	
		cellCentroCostos.appendChild(txtCCostos);	
				
	
		/*Link*/
		var imgReasignar = document.createElement("img");
		imgReasignar.src="images/compras/loop-16X16.png";
		imgReasignar.className = "cursorStyle";
		imgReasignar.title = "Reasignar";
		cellLink.align = "center";
		if (isFirefox) {
			imgReasignar.setAttribute('onclick', "reasignar(<%=idObra%>," + pBudgetArea + "," + pBudgetDetalle + ")");
		} else {
			imgReasignar['onclick'] = new Function("reasignar(<%=idObra%>," + pBudgetArea + "," + pBudgetDetalle + ")");
		}
		cellLink.appendChild(imgReasignar);

		/*Ajuste Partida Presupuestaria*/
		var imgReasignar = document.createElement("img");
		imgReasignar.src="images/compras/budget_item-16x16.png";
		imgReasignar.className = "cursorStyle";
		imgReasignar.title = "Ajuste";
		cellAjuste.align = "center";
		if (isFirefox) {
		    imgReasignar.setAttribute('onclick', "ajustePartidaPresupuestaria(<%=idObra%>," + pBudgetArea + "," + pBudgetDetalle + ")");
		} else {
		    imgReasignar['onclick'] = new Function("ajustePartidaPresupuestaria(<%=idObra%>," + pBudgetArea + "," + pBudgetDetalle + ")");
		}
		cellAjuste.appendChild(imgReasignar);
		lastPositionDetalle[pBudgetArea] = pBudgetDetalle;
	}
	
	function reasignar(idObra, idArea, idDetalle){
		var param = '?idobra=' + idObra + '&idarea=' + idArea + '&iddetalle=' + idDetalle
		var url = 'comprasBudgetPopUpReasignacion.asp' + param
		var puw = new winPopUp('popUpReasignaciones', url, 700, 450, 'Reasignacion', 'location.reload()');
	}
    
	function ajustePartidaPresupuestaria(idObra, idArea, idDetalle){
	    var param = '?idobra=' + idObra + '&idarea=' + idArea + '&iddetalle=' + idDetalle
	    var url = 'comprasAjusteBudgetPopUp.asp' + param
	    var puw = new winPopUp('popUpAjuste', url, 700, 375, 'Reasignacion', 'location.reload()');
	}

	function volver() {
		document.location.href = "comprasObras.asp";
	}

	function abrirTableroObra() {
		document.location.href = 'comprasTableroObra.asp?idObra=<%=idObra%>';
	}

	function imprimir(){
		var idObra = document.getElementById('idobra').value;
		window.open("comprasbudgetobrafilter.asp?idobra=" + idObra);
	}

	function refreshPage() {
		document.getElementById("accion").value = '';
		document.getElementById("frm").submit();
	}
	
	function grabarReasignacion(){
		document.getElementById("accion").value = '<%=ACCION_GRABAR%>';
		document.getElementById("frm").submit();
	}

	function seleccionarCombo(me){
		var nombre = 'combovalue_'+me.name;
		document.getElementById(nombre).value = me.value;
	}
	function dibujarItemNuevaPartida(pNbrRow,pBudgetArea){
		var tableRef = document.getElementById("tableBudget");
		var rowDetalle = tableRef.insertRow(pNbrRow);
		var cellAdd1 = rowDetalle.insertCell(0);
		cellAdd1.colSpan = "9";
		var cellAddPP = rowDetalle.insertCell(1);
		var imgAddPP = document.createElement("img");
		imgAddPP.src = "images/add.gif";
		imgAddPP.title = "Nueva Partida";
		imgAddPP.className = "cursorStyle";
		cellAddPP.align = "center";
		if (isFirefox) {
			imgAddPP.setAttribute('onclick', "agregarDetallePartida(" + pBudgetArea + ",'')");
		} else {
			imgAddPP['onclick'] = new Function("agregarDetallePartida(" + pBudgetArea + ",'')");
		}
		cellAddPP.appendChild(imgAddPP);
	}
	
	function agregarDetallePartida(pArea,pDsDetalle){
		counterNewDetalle[pArea]++;
		document.getElementById("rowDetalle_"+pArea).value = counterNewDetalle[pArea];
		var objLastRow = document.getElementById('detalle_'+ pArea +'_'+ lastPositionDetalle[pArea]);
		var ne = document.createElement("tr");
		var td1 = document.createElement("td");
		ne.appendChild(td1);
		td1.setAttribute('colspan','2');
		var td2 = document.createElement("td");
		var input2 = document.createElement("input");
		input2.type = "text";
		input2.id = "textDsDetalle_"+pArea+"_"+counterNewDetalle[pArea];
		input2.name = "textDsDetalle_"+pArea+"_"+counterNewDetalle[pArea];
		if (pDsDetalle != "") input2.value = pDsDetalle;
		input2.size = "50";		
		input2.maxLength = 100;		
		td2.appendChild(input2);
		ne.appendChild(td2);
		var td3 = document.createElement("td");
		ne.appendChild(td3);
		var div3 = document.createElement("div");
		div3.innerHTML = '0,00';
		div3.style.textAlign= "right";
		td3.appendChild(div3);
		var td4 = document.createElement("td");
		ne.appendChild(td4);
		var div4 = document.createElement("div");
		div4.innerHTML = '0,00';
		div4.style.textAlign= "right";
		td4.appendChild(div4);
		var td5 = document.createElement("td");
		ne.appendChild(td5);
		var div5 = document.createElement("div");
		div5.innerHTML = '0,00';
		div5.style.textAlign= "right";
		td5.appendChild(div5);
		var td6 = document.createElement("td");
		ne.appendChild(td6);
		var div6 = document.createElement("div");
		div6.innerHTML = '0,00';
		div6.style.textAlign= "right";
		td6.appendChild(div6);
		var td8 = document.createElement("td");
		ne.appendChild(td8);
		td8.setAttribute('colspan','3');
		objLastRow.parentNode.insertBefore(ne,objLastRow.nextSibling);
	}
	function restablecerDescripcionDetalle(pArea,pDetalle,pDsDescripcion){
		var myForm = document.getElementById("frm");
		var hidDsDet = document.createElement("input");
		hidDsDet.type = "hidden";
		hidDsDet.id = "hiddenDsDetalle_"+pArea+"_"+pDetalle;
		hidDsDet.name = "hiddenDsDetalle_"+pArea+"_"+pDetalle;
		hidDsDet.value = pDsDescripcion;
		myForm.appendChild(hidDsDet);
	}
</script>
<style type="text/css">
	select{
		width:100px;
	}
</style>
</HEAD>
<BODY onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Reasignacion de Presupuestos") %>	
	<div id="toolbar"></div>
	<br>
	<form name="frm" id="frm" method="POST" action="comprasBudgetReasignaciones.asp?idobra=<%=idobra%>">
		<input type="hidden" id="mensaje" name="menasje" value="Cargando">
		<input type="hidden" name="accion" id="accion" value='<%=accion%>'>
		<input type="hidden" name="idobra" id="idobra" value='<%=idObra%>'>
		<% if (hayError()) then Call showErrors() 
			if (hayError() = false AND accion = ACCION_GRABAR) then 
				reasignacionOK = true
			end if
		%>
		<div id="actionLabel" class="round_border_top TDSUCCESS" style="width:100%;visibility:hidden;">
			<script> document.write(document.getElementById("mensaje").value);</script>...
		</div>
		<br>
		<table id="tableBudget" width="100%" border="0" align="center" class="reg_header">
	  <tr><td class="titu_header  round_border_all" colspan="12">
			<table border="0"><tr><td align="center"><img src='images/compras/OBR-48X48.png' width="32" height="32"></td>
			<td>&nbsp;</td>
			<td class="titu_header" style="border:none;"><% =GF_TRADUCIR("OBRA : " & getDescripcionObra(idObra)) %></td></tr>
			</table>
		</td></tr>
		<tr>
	      	<td colspan="3" rowspan="2" class="titu_header round_border_top" ><%=GF_TRADUCIR("Detalle")%></td>
	        <td colspan="3" align="center" width="24%" class="titu_header round_border_top"><%=GF_TRADUCIR("Gastos (USD)")%></td>
	        <td rowspan="2" align="center" width="8%" class="titu_header round_border_top_left" ><%=GF_TRADUCIR("Budget (USD)")%></td>
            <% if(CInt(tipoFormularioPartida) = OBRA_FORM_ANUAL)then %>
				<td rowspan="2" width="15%" align="center" class="titu_header"><%=GF_TRADUCIR("Subproyecto")%></td>
			<% else %>
				<td rowspan="2" width="8%" align="center" class="titu_header"><%=GF_TRADUCIR("Cuenta")%></td>
				<td rowspan="2" width="5%" align="center" class="titu_header round_border_top_right"><%=GF_TRADUCIR("CC")%></td>
	        <% end if %>
	        <td width="3%" rowspan="2" align="center" class="titu_header round_border_top">.</td>
            <td width="3%" rowspan="2" align="center" class="titu_header round_border_top">.</td>
            <td width="3%" rowspan="2" align="center" class="titu_header round_border_top">.</td>
	    </tr>
		  <tr class="titu_header">
		    	<td width="8%"><div align="center"><%=GF_TRADUCIR("Vales")%></div></td>
				<td width="8%"><div align="center"><%=GF_TRADUCIR("Pedido")%></div></td>
				<td width="8%"><div align="center"><%=GF_TRADUCIR("Facturado")%></div></td>
       	  </tr>
			<% 
			rs.movefirst
			while not rs.eof
					%>
					<tr>
						<td class="reg_header_nav round_border_left" width="2%" align="right"><%=rs("IDAREA")%></td>
						<td class="reg_header_navdos " width="40%" colspan="2"> <%=rs("DSBUDGET")%> </td>	
						<td class="reg_header_navdos" align="right"><div id='VlVales_<%=rs("IDAREA")%>'></div></td>	
						<td class="reg_header_navdos" align="right"><div id='VlPedidos_<%=rs("IDAREA")%>'></div></td>	
						<td class="reg_header_navdos" align="right"><div id='VlFacturado_<%=rs("IDAREA")%>'></div></td>	
						<td class="reg_header_navdos" align="right"><div id='VlBudget_<%=rs("IDAREA")%>'></div></td>	
						<td class="reg_header_navdos" colspan="5" align="right">&nbsp;</td>	
					</tr>
					</tr>
					<tr id='AREA_<%=rs("IDAREA")%>' >
					<td align="center" width="100%" colspan="8">
						<img id="image<%=rs("IDAREA")%>" src="images/Loading1.gif">
						<input type="hidden" id="rowDetalle_<%=rs("IDAREA")%>" name="rowDetalle_<%=rs("IDAREA")%>"/>
					</td>
					</tr>
			<%
				rs.movenext
			wend
			%>
			<tr>
				<td class="reg_header_nav" colspan="3" align="right"><%=GF_TRADUCIR("Total")%>&nbsp;</td>
				<td class="reg_header_nav" align="right"><div id='VlGralVales'></div></td>	
				<td class="reg_header_nav" align="right"><div id='VlGralPedidos'></div></td>	
				<td class="reg_header_nav" align="right"><div id='VlGralFacturado'></div></td>	
				<td class="reg_header_nav" align="right"><div id='VlGralBudget'></div></td>	
				<td class="reg_header_nav" colspan="5" align="right">&nbsp;</td>
			</tr>			
		</table>
	</form>
	<br>
		<div id="actionLabel2" class="round_border_bottom TDSUCCESS" style="width:100%;visibility:hidden;">
			<script> document.write(document.getElementById("mensaje").value);</script>...
		</div>
</BODY>
</HTML>