<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->

<%
Const XLS_REPORT = "XLS"
Const TXT_REPORT = "TXT"
'******************************************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
dim pto,rsGeneral,accion,strSQLPro,paginaActual,mostrar,lineasTotales,flagReport

totalVagones = 0
totalKilosNetos = 0
Call GP_CONFIGURARMOMENTOS()

g_strPuerto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", g_strPuerto, params)
accion = GF_PARAMETROS7("accion", "", 6)



call getParametros()
if not hayError() then
	Set rsGeneral = loadOperativosPuertos()
	paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
	if (paginaActual = 0) then paginaActual = 1
	mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
	if (mostrar = 0) then mostrar = 10	
	Call setupPaginacion(rsGeneral, paginaActual, mostrar)
	lineasTotales = rsGeneral.recordcount
end if

%>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=9">

<title><%=GF_TRADUCIR("Puertos - Administrar Operativos")%></title>
<link rel="stylesheet" href="../../css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../../css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../../css/main.css" type="text/css"> 

<script type="text/javascript" src="../../scripts/formato.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="../../scripts/calendar.js"></script>
<script type="text/javascript" src="../../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../../scripts/iwin.js"></script>
<script type="text/javascript" src="../../scripts/paginar.js"></script>
<script type="text/javascript">	
	var ch = new channel();		
	var changeFilters = false;
	var maxSegments;
	var currSegment=0;
	var optionReport;	
	var MS_X_DAY = 86400000 //Milisegundos por d�a.	
	var d = new Date();	
	
	function bodyOnLoad() {		
		tb = new Toolbar('toolbar');
		tb.addButton("toolbar-excel", "Generar XLS", "cargarReporte('<%=XLS_REPORT%>')");
		tb.addButton("toolbar-excel", "Generar TxT", "cargarReporte('<%=TXT_REPORT%>')");
		tb.addButton("toolbar-excel", "Informe Camara", "cargarInformeCamara()");
		tb.draw();
		autoCompleteCoordinado();
		autoCompleteVendedor();
		autoCompleteCorredor();
		autoCompleteEntregador();		
		<% 	if not hayError() then
				if (not rsGeneral.eof) then %>
					var pgn = new Paginacion("paginacion");
					pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "AdministracionOperativos.asp<% =params %>");
		<%		end if
			end if	%>
		
	}
	
	
	function cargarInformeCamara(){
	    window.open('vagonesInformeCamara.asp?pto=<%=g_strPuerto%>&firstTime=1','<%=g_strPuerto%>','width=1200,height=800,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');	   
    }		
			
	function cargarReporte(pOpcion){
		<% if not hayError() then %>			
			if (changeFilters) {  
				alert ("Atencion!\nSe cambiaron los filtros de b�squeda, por favor genere nuevamente el informe.");
				return 0;
			}
			optionReport = pOpcion;			
			generateReport();
		<% end if %>
	}
	
	function generateReport() {
		document.getElementById("fileCode").value = "_" + document.getElementById("usr").value + "_" + currSegment;
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel").innerHTML = "Inicializando... ";
		calculateSegments();
		document.getElementById("frmSel").action="OperativosInformePrintE1.asp";
		document.getElementById("frmSel").target="ifrmReport";		
		generateSegment(currSegment)
	}	
	
	function restartAttForm(){
		document.getElementById("accion").value = '<%=ACCION_SUBMITIR%>';
		var obj = document.getElementById("frmSel");
		obj.removeAttribute('target');
		obj.removeAttribute('action');
		var myFechaContD = document.getElementById("dtFechaDesde").value;		
		str = document.getElementById("dtFechaDesde").value;
		currSegment = 0;
	    document.getElementById("fecContableD").value = str.substr(0,2);
	    document.getElementById("fecContableM").value = str.substr(3,2);
	    document.getElementById("fecContableA").value = str.substr(6,4);
	    strH = document.getElementById("dtFechaHasta").value;
	    document.getElementById("fecContableDH").value = strH.substr(0,2);
	    document.getElementById("fecContableMH").value = strH.substr(3,2);
	    document.getElementById("fecContableAH").value = strH.substr(6,4);
	}		
	function generateSegment(currSegment) {
		document.getElementById("actionLabel").innerHTML = "Recopilando datos...  ( " + (currSegment+1) + " / " + (maxSegments+1) + " )";
		var d = document.getElementById("fecContableD").value;
		var m = document.getElementById("fecContableM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fecContableA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		var d = new Date(fd.getTime() + (MS_X_DAY*currSegment));
		document.getElementById("fecContableDS").value = d.getDate();
		document.getElementById("fecContableMS").value = d.getMonth()+1;	//getMonth() entrega el nro de mes de 0 a 11.
		document.getElementById("fecContableAS").value = d.getFullYear();		
		document.getElementById("frmSel").submit();
	}
	
	function generateSegment_callback(){
		if (currSegment < maxSegments) {
			currSegment += 1; 
			document.getElementById("fileCode").value = "_" + document.getElementById("usr").value + "_" + currSegment;
			generateSegment(currSegment);
			
		} else {
			document.getElementById("maxSegment").value = currSegment;			
			document.getElementById("accion").value = '<%=ACCION_PROCESAR%>';
			if (optionReport == '<%=XLS_REPORT%>') generateExcel();
			if (optionReport == '<%=TXT_REPORT%>') generateTextFile();
		}
	 }
	
	function generateExcel() {		
		document.getElementById("actionLabel").innerHTML = "Generando Excel...";
		setTimeout("document.getElementById('actionLabel').style.visibility = 'hidden'", 3000);
		document.getElementById("frmSel").action="operativosInformeXLS.asp";
		document.getElementById("frmSel").target="";		
		document.getElementById("frmSel").submit();
		restartAttForm();
	}
	function generateTextFile(){
		document.getElementById("actionLabel").innerHTML = "Generando Archivo de Texto...";
		ch.bind("OperativosInformeTXT.asp<%=params%>&maxSegment="+document.getElementById("maxSegment").value, "generateTextFile_Callback()");
		ch.send();
		document.getElementById("frmSel").target="";
		restartAttForm();
	}	
	function generateTextFile_Callback(){		
		document.getElementById("actionLabel").innerHTML = "Archivo generado con exito<br>Click <u><a href='" + ch.response() + "' style='cursor:pointer;' >aqui</a></u> para ir al archivo.";	
	}
	function calculateSegments() {
		var d = document.getElementById("fecContableD").value;
		var m = document.getElementById("fecContableM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fecContableA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		d = document.getElementById("fecContableDH").value;
		m = document.getElementById("fecContableMH").value-1; //El Month de Date trabaja de 0 a 11
		y = document.getElementById("fecContableAH").value;
		var fh = new Date(y, m, d, 0, 0, 0, 0);		
		maxSegments = Math.round((fh.getTime() - fd.getTime())/MS_X_DAY)		
	}	
	
	function volver() {	
		location.href = "../puertosReportes.asp?pto=<%=g_strPuerto%>";
	}

	function habilitarLoading(pVisibility, pPosition){
		document.getElementById("imgLoading").style.position = pPosition;
		document.getElementById("imgLoading").style.visibility  = pVisibility;
		document.getElementById("lblLoading").style.position = pPosition;
		document.getElementById("lblLoading").style.visibility  = pVisibility;
		if (pVisibility=='visible')
			document.getElementById("actionLabel").style.visibility  = "hidden";
		else	
			document.getElementById("actionLabel").style.visibility  = "visible";
	}

	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function cambioBusqueda(){
		changeFilters = true;
	}			
	function setSortBy(pInput, pCol, pOrder){	
		document.getElementById("myOrder").value = pCol + ' ' + pOrder;		
		document.getElementById(pInput).value = pOrder;
		submitInfo();
	}	
	function submitInfo(){
		document.getElementById("frmSel").submit();
	}
	function verVagones(cdOperativo, cartaPorte, dtContable, cdproducto){
		myPopUp = new PopUpWindow('Iframe', 'operativosPopUp.asp?Pto=<%=g_strPuerto%>&cdOperativo=' + cdOperativo + '&fecha=' + dtContable + '&cartaPorte=' + cartaPorte +"&cdProducto=" + cdproducto, '700', '600', 'Informacion de Vagones');
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
	
	function SeleccionarCalDesde(cal, date) {
		var str= new String(date);		
		document.getElementById("dtFechaDesde").value = str;
	    document.getElementById("fecContableD").value = str.substr(0,2);
	    document.getElementById("fecContableM").value = str.substr(3,2);
	    document.getElementById("fecContableA").value = str.substr(6,4);
		if (cal) cal.hide();
	}	
	function SeleccionarCalHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("dtFechaHasta").value = str;	    
	    document.getElementById("fecContableDH").value = str.substr(0,2);
	    document.getElementById("fecContableMH").value = str.substr(3,2);
	    document.getElementById("fecContableAH").value = str.substr(6,4);	    
		if (cal) cal.hide();	
	}	
	
	
	function autoCompleteVendedor(){
		$( "#dsVendedor" ).autocomplete({
			minLength: 2,
			source: "../puertosStreamElementos.asp?tipo=JQVendedores&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsVendedor").val(ui.item.dsvendedor);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsVendedor"    ).val (ui.item.dsvendedor);
				$( "#cdVendedor"    ).val (ui.item.cdvendedor);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsVendedor").val ("");
					$( "#cdVendedor").val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {			
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdvendedor + " - <font style='font-size:10;'>" + item.dsvendedor + "</font></a>" )
				.appendTo( ul );
		};
	}		
		
	function autoCompleteCoordinado(){
		$( "#dsCoordinado" ).autocomplete({
			minLength: 2,
			source: "../puertosStreamElementos.asp?tipo=JQClientes&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsCoordinado").val(ui.item.dscliente);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCoordinado"    ).val (ui.item.dscliente);
				$( "#cdCoordinado"    ).val (ui.item.cdcliente);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsCoordinado").val ("");
					$( "#cdCoordinado").val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdcliente + " - <font style='font-size:10;'>" + item.dscliente + "</font></a>" )
				.appendTo( ul );
		};
	}
	
	function autoCompleteCorredor(){
		$( "#dsCorredor" ).autocomplete({
			minLength: 2,
			source: "../puertosStreamElementos.asp?tipo=JQCorredores&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsCorredor").val(ui.item.dscorredor);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCorredor"    ).val (ui.item.dscorredor);
				$( "#cdCorredor"    ).val (ui.item.cdcorredor);
				return false;
			},
			change: function( event, ui ) {				
				if (!ui.item) {					
					$( "#dscorredor").val ("");
					$( "#cdcorredor").val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdcorredor + " - <font style='font-size:10;'>" + item.dscorredor + "</font></a>" )
				.appendTo( ul );
		};
	}	
	function autoCompleteEntregador(){
		$( "#dsEntregador" ).autocomplete({
			minLength: 2,
			source: "../puertosStreamElementos.asp?tipo=JQEntregadores&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsEntregador").val(ui.item.dsentregador);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsEntregador"    ).val (ui.item.dsentregador);
				$( "#cdEntregador"    ).val (ui.item.cdentregador);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsEntregador").val ("");
					$( "#cdEntregador").val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdentregador + " - <font style='font-size:10;'>" + item.dsentregador + "</font></a>" )
				.appendTo( ul );
		};
	}	
</script>
</head>

<body onLoad="bodyOnLoad()">

<div id="toolbar"></div>

<form name="frmSel" id="frmSel">	
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> filtro - <%=GF_Traducir("Consulta de Operativos")%> </h3>
    
    <div id="searchfilter" class="tableasidecontent">
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Operativo") %> </div>
        <div class="col16"> <input type="text" onChange="cambioBusqueda();" id="Operativo" maxLength="12" size="18" name="Operativo" value="<% =myOperativo %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Carta de Porte") %> </div>
        <div class="col16"> <input type="text" onChange="cambioBusqueda();" id="CartaPorte" maxLength="16" size="18" name="CartaPorte" value="<% =myCartaPorte %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Turno") %> </div>
        <div class="col16"> <input type="text" onChange="cambioBusqueda();" id="turno" maxLength="6" size="5" name="turno" value="<% =myTurno %>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Inicio Desde") %> </div>
        <div class="col16"> <input type="text" name="dtFechaDesde" id="dtFechaDesde" readonly onClick="javascript:MostrarCalendario('dtFechaDesde', SeleccionarCalDesde)" value="<% =myFecContableD &"/"& myFecContableM &"/"& myFecContableA%>">
        		<input type="hidden" id="fecContableD" name="fecContableD" value="<%=myFecContableD%>">
				<input type="hidden" id="fecContableM" name="fecContableM" value="<%=myFecContableM%>">
				<input type="hidden" id="fecContableA" name="fecContableA" value="<%=myFecContableA%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Inicio Hasta") %> </div>        
        <div class="col16"> <input type="text" name="dtFechaHasta" id="dtFechaHasta" readonly onClick="javascript:MostrarCalendario('dtFechaHasta', SeleccionarCalHasta)" value="<% =myFecContableDH &"/"& myFecContableMH &"/"& myFecContableAH%>">
        		<input type="hidden" id="fecContableDH" name="fecContableDH" value="<%=myFecContableDH%>">
				<input type="hidden" id="fecContableMH" name="fecContableMH" value="<%=myFecContableMH%>">
				<input type="hidden" id="fecContableAH" name="fecContableAH" value="<%=myFecContableAH%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Cordinado") %> </div>
        <div class="col16">
				<input type="hidden" id="cdCoordinado" name="cdCoordinado" value="<%=myCdCoordinado%>">
				<input type="text"   id="dsCoordinado" name="dsCoordinado" value="<%=myDsCoordinado%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Corredor") %> </div>
        <div class="col16">
				<input type="hidden" id="cdCorredor" name="cdCorredor" value="<%=myCdCorredor%>">
				<input type="text" id="dsCorredor" name="dsCorredor" value="<%=myDsCorredor%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Vendedor") %> </div>
        <div class="col16">
				<input type="hidden" id="cdVendedor" name="cdVendedor" value="<%=myCdVendedor%>">
				<input type="text" id="dsVendedor" name="dsVendedor" value="<%=myDsVendedor%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Entregador") %> </div>
        <div class="col16">
				<input type="hidden" id="cdEntregador" name="cdEntregador" value="<%=myCdEntregador%>">
				<input type="text" id="dsEntregador" name="dsEntregador" value="<%=myDsEntregador%>"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Estado") %> </div>
        <div class="col16"> 
			<% strSQLEst = "SELECT * FROM ESTADOSOPERATIVOS ORDER BY CDESTADO"
				 call GF_BD_Puertos(g_strPuerto, rsEstado, "OPEN",strSQLEst) %>
				<select onChange="cambioBusqueda();" name="cmbEstado" id="cmbEstado" value="<%=myEstado%>">
					<option value=""> <%=GF_Traducir("TODOS")%></option>
					<%while not rsEstado.eof
						mySelected = ""
						if trim(rsEstado("CDESTADO")) = trim(myEstado) then mySelected = "SELECTED"%>
						<option value="<%=rsEstado("CDESTADO")%>" <%=mySelected%>> <%=Trim(Ucase(rsEstado("DSESTADO")))%></option>
						<%
						rsEstado.movenext
					 wend%>
				</select>
        </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Nro.Vag�n") %> </div>
        <div class="col16"> <input type="text" id="nroVagon" name="nroVagon" onChange="cambioBusqueda();" value="<%=myIdVagon%>" onKeyPress="return controlIngreso (this, event, 'N');"> </div>
        
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Producto") %> </div>
        <div class="col16"> 
				<% strSQLPro = "SELECT * FROM PRODUCTOS ORDER BY DSPRODUCTO"
				 call GF_BD_Puertos(g_strPuerto, rsProducto, "OPEN",strSQLPro)
				 %>
					<select onChange="cambioBusqueda();" name="cdProducto" value="<%=myCdProducto%>">
						<option value=""> <%=GF_Traducir("TODOS")%></option>
						<%while not rsProducto.eof
							mySelected = ""
							if trim(rsProducto("CDPRODUCTO")) = trim(myCdProducto) then mySelected = "SELECTED"%>
							<option value="<%=rsProducto("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProducto("DSPRODUCTO")%></option>
							<%
							rsProducto.movenext
						 wend%>
				</select>
        </div>
        
        <span class="btnaction"><input type="submit" value="Buscar" id=submit1 name=submit1></span>
    </div>
</div><!-- END BUSCAR -->

<div class="col66"></div>


	<% Call showErrors() %>
  	
	<%
	if not hayError() then %>
		<TABLE class="datagrid" id="TAB1" align="center" width="100%">
	<%	if not rsGeneral.eof then%>
    <thead>
			<TR class="reg_Header_nav">
				<th align="center"><%=GF_Traducir("Turno")%> 
					<%if orderByTurno = "ASC" then 
						myTitle="Descendiente"
						orderByTurno="DESC"
					  else
						myTitle="Ascendiente"
						orderByTurno="ASC"
					  end if %>
					<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByTurno','T1.SQTURNO','<%=orderByTurno%>')">				
					<input type="hidden" id="orderByTurno" name="orderByTurno" value="<%=orderByTurno%>">
				</th>
				<th align="center"><%=GF_Traducir("Operativo")%> 
				<%if orderByOperativo = "ASC" then 
					myTitle="Descendiente"
					orderByOperativo="DESC"
				  else
					myTitle="Ascendiente"
					orderByOperativo="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByOperativo','T1.CDOPERATIVO','<%=orderByOperativo%>')">				
				<input type="hidden" id="orderByOperativo" name="orderByOperativo" value="<%=orderByOperativo%>">
				</th>
				<th align="center"><%=GF_Traducir("Carta Porte")%> 
				<%if orderByCartaPorte = "ASC" then 
					myTitle="Descendiente"
					orderByCartaPorte="DESC"
				  else
					myTitle="Ascendiente"
					orderByCartaPorte="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByCartaPorte','T1.cdoperativoserie','<%=orderByCartaPorte%>')">				
				<input type="hidden" id="orderByCartaPorte" name="orderByCartaPorte" value="<%=orderByCartaPorte%>">				
				</th>
				<th align="center"><%=GF_Traducir("Fecha Inicio")%> 
				<%if orderByFechaInicio = "ASC" then 
					myTitle="Descendiente"
					orderByFechaInicio="DESC"
				  else
					myTitle="Ascendiente"
					orderByFechaInicio="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByFechaInicio','T1.DTINICIO','<%=orderByFechaInicio%>')">
				<input type="hidden" id="orderByFechaInicio" name="orderByFechaInicio" value="<%=orderByFechaInicio%>">
				</th>
				<th align="center"><%=GF_Traducir("Coordinado")%> 
				<%if orderByCoordinado = "ASC" then 
					myTitle="Descendiente"
					orderByCoordinado="DESC"
				  else
					myTitle="Ascendiente"
					orderByCoordinado="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByCoordinado','CL.DSCLIENTE','<%=orderByCoordinado%>')">				
				<input type="hidden" id="orderByCoordinado" name="orderByCoordinado" value="<%=orderByCoordinado%>">
				</th>
				<th align="center"><%=GF_Traducir("Producto")%> 
				<%if orderByProducto = "ASC" then 
					myTitle="Descendiente"
					orderByProducto="DESC"
				  else
					myTitle="Ascendiente"
					orderByProducto="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByProducto','P.dsproducto','<%=orderByProducto%>')">				
				<input type="hidden" id="orderByProducto" name="orderByProducto" value="<%=orderByProducto%>">
				</th>
				<th align="center"><%=GF_Traducir("Corredor")%> 
				<%if orderByCorredor = "ASC" then 
					myTitle="Descendiente"
					orderByCorredor="DESC"
				  else
					myTitle="Ascendiente"
					orderByCorredor="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByCorredor','CO.DSCORREDOR','<%=orderByCorredor%>')">				
				<input type="hidden" id="orderByCorredor" name="orderByCorredor" value="<%=orderByCorredor%>">
				</th>
				<th align="center"><%=GF_Traducir("Vendedor")%> 				
				<%if orderByVendedor = "ASC" then 
					myTitle="Descendiente"
					orderByVendedor="DESC"
				  else
					myTitle="Ascendiente"
					orderByVendedor="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByVendedor','VE.DSVENDEDOR','<%=orderByVendedor%>')">				
				<input type="hidden" id="orderByVendedor" name="orderByVendedor" value="<%=orderByVendedor%>">				
				</th>
				<th align="center"><%=GF_Traducir("Estado")%> 
				<%if orderByEstado = "ASC" then 
					myTitle="Descendiente"
					orderByEstado="DESC"
				  else
					myTitle="Ascendiente"
					orderByEstado="ASC"
				  end if %>
				<img src="../../images/orderlist.png" title="<%=myTitle%>" style="cursor:pointer;" onClick="setSortBy('orderByEstado','EST.DSESTADO','<%=orderByEstado%>')">
				<input type="hidden" id="orderByEstado" name="orderByEstado" value="<%=orderByEstado%>">
				</th>
				<th align="center" colspan="2">Vagones</th>
			</TR>
            </thead>
            <tbody>
		<%	reg = 0
			while not rsGeneral.eof and (reg < mostrar) 
				reg = reg + 1 %>
			<tr>
				<TD align="right"><%=rsGeneral("SQTURNO")%></TD>
				<TD align="center"><%=Left(rsGeneral("OPERATIVO"), 12) %></TD>
				<TD align="center"><%=GF_EDIT_CTAPTE(Left(rsGeneral("CARTAPORTE"), 12))%></TD>
				
				<TD align="center"><%If(not IsNull(rsGeneral("DTINICIO")))then Response.Write GF_FN2DTE(rsGeneral("DTINICIO"))%> </TD>				
				<TD align="center"><%=rsGeneral("DSCLIENTE")%></TD>
				<TD align="left"><%=rsGeneral("DSPRODUCTO")%></TD>
				<TD align="left"><%=rsGeneral("DSCORREDOR")%></TD>
				<TD align="left"><%=rsGeneral("DSVENDEDOR")%></TD>
				<TD align="left"><%=rsGeneral("DSESTADO")%> </TD>				
				<%'if (not isNull(rsGeneral("DTINICIO"))) then dtContable = Year(rsGeneral("DTINICIO")) & "-" & GF_nDigits(Month(rsGeneral("DTINICIO")), 2) & "-" & GF_nDigits(Day(rsGeneral("DTINICIO")), 2) %>				
				<td align="center"><% =rsGeneral("QTVAGONES") %></td>
				<TD align="center"><img title="Detalle de Vagones" src="../../images/buscar-16.png" style="cursor:pointer;" onClick="javascript:verVagones('<%=rsGeneral("CDOPERATIVO")%>','<%= rsGeneral("CARTAPORTE")%>','<%=rsGeneral("DTCONTABLE")%>','<%=rsGeneral("CDPRODUCTO")%>');"></TD>
			</tr>
             </tbody>
            
            <tfoot>
                <tr>
                    <td colspan="12"><div id="paginacion"></div></td>
                </tr>
            </tfoot>            
                        
		<%
		rsGeneral.movenext
		wend %>	  	
	<%	else  %>
   	<tfoot>
		<tr>
        	<td align="center" colspan="12" class="reg_Header_Warning"><%=GF_TRADUCIR("No se encontraron datos")%></td>
        </tr>
    </tfoot>
	<%	end if	%>
   
	  </TABLE>	
<%	end if%>
	<input type="hidden" id="sortBy" name="sortBy" VALUE="<%=sortBy%>">
	<input type="hidden" id="fileCode" name="fileCode" value="">
	<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">
	<input type="hidden" id="pto" name="pto" value="<% =g_strPuerto %>">	
	<input type="hidden" id="fecContableDS" name="fecContableDS">
	<input type="hidden" id="fecContableMS" name="fecContableMS">
	<input type="hidden" id="fecContableAS" name="fecContableAS">
	<input type="hidden" id="maxSegment" name="maxSegment">
	<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>">
	<input type="hidden" id="sector" name="sector" value="<% =sector %>">		
	<div id="actionLabel" class="confirmsj" style="width:80%;visibility:hidden;"></div>
	<input type="hidden" name="myOrder" id="myOrder" value="<%=myOrder%>">	
</form>
<iframe name="ifrmReport" id="ifrmReport" width="1px" height="1px" style="visibility:hidden"></iframe>
</body>
</html>
