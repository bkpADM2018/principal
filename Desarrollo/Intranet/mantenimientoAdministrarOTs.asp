<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
Call initAccessInfo(RES_OT_SM)

Dim txtIdEquipo, txtDsEquipo, myWhere, strSQL, rs, conn, ref, listaDivisionesDefault, filtroDivision

ref = GF_PARAMETROS7("ref","",6)
if trim(ref) = "" then ref = "mantenimientoOTIndex.asp"
call addParam("ref", ref, params)

txtIdOrder = GF_PARAMETROS7("txtIdOrder", 0, 6)
call addParam("txtIdOrder", txtIdOrder, params)

txtNroOrder = GF_PARAMETROS7("txtNroOrder", "", 6)
call addParam("txtNroOrder", txtNroOrder, params)

txtDsOrder = GF_PARAMETROS7("txtDsOrder", "", 6)
call addParam("txtDsOrder", txtDsOrder, params)

listaDivisionesDefault = getListaCargosAdmin()
filtroDivision = listaDivisionesDefault
txtIdDivision = GF_PARAMETROS7("txtIdDivision", 0, 6)
if (txtIdDivision <> 0) then filtroDivision = txtIdDivision
call addParam("txtIdDivision", txtIdDivision, params)

txtIdEquipoActivo = GF_PARAMETROS7("txtIdEquipoActivo", 0, 6)
call addParam("txtIdEquipoActivo", txtIdEquipoActivo, params)

txtIdSector = GF_PARAMETROS7("txtIdSector", 0, 6)
call addParam("txtIdSector", txtIdSector, params)

solicitante = GF_PARAMETROS7("solicitante","",6)
call addParam("solicitante", solicitante, params)
if(solicitante <> "")then cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)

responsable = GF_PARAMETROS7("responsable","",6)
call addParam("responsable", responsable, params)
if responsable = "" then 
	cdResponsable = 0
else	
	cdResponsable = GF_PARAMETROS7("cdResponsable",0,6)
end if	

txtSoloPlanificadas = GF_PARAMETROS7("txtSoloPlanificadas", "" ,6)
if txtSoloPlanificadas = "" then txtSoloPlanificadas = 1
call addParam("txtSoloPlanificadas", txtSoloPlanificadas, params)


txtTipoMantenimiento = GF_PARAMETROS7("txtTipoMantenimiento","",6)
if(txtTipoMantenimiento = "") then txtTipoMantenimiento = "T"
call addParam("txtTipoMantenimiento", txtTipoMantenimiento, params)

txtTipoOrden = GF_PARAMETROS7("txtTipoOrden","",6)
if(txtTipoOrden = "") then txtTipoOrden = "T"
call addParam("txtTipoOrden", txtTipoOrden, params)

dtProgramada = GF_PARAMETROS7("dtProgramada","",6)
call addParam("dtProgramada", dtProgramada, params)

txtCdState = GF_PARAMETROS7("txtCdState", 0, 6)
call addParam("txtCdState", txtCdState, params)

myOrder = GF_PARAMETROS7("myOrder", "" ,6)
call addParam("myOrder", myOrder, params)
if myOrder = "" then myOrder = " ORD.CDSTATE ASC, ORD.SCHEDULEDDATE ASC "
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
call addParam("paginaActual", paginaActual, params)
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
call addParam("mostrar", mostrar, params)
myDtProgrmada = GF_DTE2FN(dtProgramada)
if myDtProgrmada = "" then myDtProgrmada = 0

Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_GET_BY_PARAMETERS", txtIdOrder & "||" & txtNroOrder & "||" & txtDsOrder & "||" & cdSolicitante & "||" & filtroDivision & "||" & txtIdEquipoActivo & "||" & txtIdSector & "||" & txtTipoOrden & "||" & txtTipoMantenimiento & "||" & myDtProgrmada & "||" & txtCdState & "||" & cdResponsable & "||" & txtSoloPlanificadas & "||" & myOrder & "||" & paginaActual & "||" & mostrar & "$$lineasTotales")
lineasTotales = sp_ret("lineasTotales")

'Parametros del ordenamiento
orderById = GF_PARAMETROS7("orderById", "" ,6)
orderByNro = GF_PARAMETROS7("orderByNro", "" ,6)
orderByTit = GF_PARAMETROS7("orderByTit", "" ,6)
orderByDiv = GF_PARAMETROS7("orderByDiv", "" ,6)
orderByEqu = GF_PARAMETROS7("orderByEqu", "" ,6)
orderBySec = GF_PARAMETROS7("orderBySec", "" ,6)
orderBySol = GF_PARAMETROS7("orderBySol", "" ,6)
orderByTiM = GF_PARAMETROS7("orderByTiM", "" ,6)
orderByTiO = GF_PARAMETROS7("orderByTiO", "" ,6)
orderByMan = GF_PARAMETROS7("orderByMan", "" ,6)
orderByFcP = GF_PARAMETROS7("orderByFcP", "" ,6)
orderByEst = GF_PARAMETROS7("orderByEst", "" ,6)
'-----------------------------------------------------------------------
sub setSortParams(byref pOrdenActual, byref pTitle)
if pOrdenActual = "ASC" then 
	pTitle="Descendiente"
	pOrdenActual="DESC"
else
	pTitle="Ascendiente"
	pOrdenActual="ASC"
end if 
end sub				
%>

<html>
<head>
<title>Sistema de Mantenimiento - Administrar Ordenes de Trabajo</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
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
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>


<script type="text/javascript">
	var ch = new channel();	
	/* Barra de herramientas de almacenes */
	function bodyOnLoad() {
		var toolBarEquipos = new Toolbar("toolBarEquipos", 8, 'images/');
		toolBarEquipos.addButtonRETURN("Volver", "irA('<%=ref%>')");	
		<% if isAdmin(SIN_DIVISION) then %>	
		toolBarEquipos.addButton("ot-add-16.png", "Nueva OT", "irA('mantenimientoAgregarOT.asp')");
		<% end if %>
		toolBarEquipos.addButtonREFRESH("Refrescar", "submitInfo()");
		toolBarEquipos.addButton("toolbar-print", "Imprimir Seleccionadas", "printMultipleOTs()");
 		toolBarEquipos.draw();
		<%if (not rs.eof) then	%>
			var pgn = new Paginacion("paginacion");
			pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "mantenimientoAdministrarOTs.asp<% =params %>");
		<%end if %>
		autoCompleteSolicitante();
		autoCompleteEmpresaResponsable();
	}
	function irA(pLink) {
		location.href = pLink;
	}	
	function submitInfo() {		
		document.getElementById("frmSel").submit();
	}

	function setOrder(pInput, pCol, pOrder){
		document.getElementById("myOrder").value =  pCol + ' ' + pOrder;
		document.getElementById(pInput).value = pOrder;
		submitInfo();
	}
	function editarOT(pIdOrder){
		document.location.href = "mantenimientoAgregarOT.asp?idOT=" + pIdOrder + "&modoEdicion=1";
	}	

	function autoCompleteEmpresaResponsable()
		{
			$( "#responsable" ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQEmpresas",
					focus: function( event, ui ) {
						$( "#responsable").val(ui.item.idempresa);
						return false;
					},
					select: function( event, ui ) {
						$( "#responsable"    ).val (ui.item.dsempresa);
						$( "#cdResponsable"    ).val (ui.item.idempresa);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#responsable").val ("");
							$( "#cdResponsable").val ("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.idempresa + " - <font style='font-size:10;'>" + item.dsempresa + "</font></a>" )
						.appendTo( ul );
				};
		}			
	function autoCompleteSolicitante()
		{
			$( "#solicitante" ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQPersonas",
					focus: function( event, ui ) {
						$( "#solicitante").val(ui.item.nombre);
						return false;
					},
					select: function( event, ui ) {
						$( "#solicitante"    ).val (ui.item.nombre);
						$( "#cdSolicitante"    ).val (ui.item.cdusuario);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#soliciante").val ("");
							$( "#cdSolicitante").val ("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
						.appendTo( ul );
				};
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
	function SeleccionarCalLimite(cal, date) {
		var str= new String(date);		
		document.getElementById("dtProgramadaDiv").innerHTML = str + " <a href='javascript:QuitarFecha()'><img src='images/button_cancel.png'></a>";
	    document.getElementById("dtProgramada").value = str;
		if (cal) cal.hide();	
	}		
	function QuitarFecha(){
		document.getElementById("dtProgramadaDiv").innerHTML = "";
	    document.getElementById("dtProgramada").value = "";
	}
	function cambiarEstadoOT(pIdOT, pCdEstado, pObservaciones){
		ch.bind("mantenimientoOtABMAJAX.asp?idOrder=" + pIdOT + "&cdEstado=" + pCdEstado + "&observations=" + pObservaciones + "&tipoOpr=UO", "cambiarEstadoOT_Callback()");
		ch.send();			
	}
	function cambiarEstadoOT_Callback(){
		submitInfo();
	}
	function cancelarOT(pIdOT, pNroOt) {				
		var puw = new winPopUp('popupEquipo','mantenimientoCancelarOTPopUp.asp?idOT=' + pIdOT + '&nroOT=' + pNroOt,'500','230','Cancelar Orden de Trabajo', "submitInfo()");
	}	
	function iniciarOT(pIdOT, pNroOt) {				
		var puw = new winPopUp('popupEquipo','mantenimientoIniciarOTPopUp.asp?idOT=' + pIdOT + '&nroOT=' + pNroOt,'500','180','Iniciar Orden de Trabajo', "submitInfo()");
	}	
	function adjuntarArchivoOT(pIdOT, pNroOt) {				
		var puw = new winPopUp('popupEquipo','mantenimientoOTFilesPopUp.asp?idOT=' + pIdOT + '&nroOT=' + pNroOt,'550','400','Adjuntar Archivo a Orden de Trabajo', "");
	}	
	function imprimirOT(pIdOT) {
		window.open('mantenimientoOtPrint.asp?idOT=' + pIdOT,"_new" + pIdOT);
	}	
	function verPlanificacionOT(pIdOT) {		
		document.location.href = "mantenimientoAdministrarPlanificacion.asp?txtIdOrder=" + pIdOT;
	}	
	
	function printMultipleOTs() {	    
	    var otList = "";
	    $(":checkbox:checked").each(function () {
	        if (otList != "") otList += ",";
	        otList += $(this).val();
	    });	    
	    window.open('mantenimientoOtPrint.asp?idOT=' + otList,"_new" + Math.floor((Math.random() * 1000)));	    
	}
	
	function finalizarOT(pIdOT, pAcc) {		
		var auxURL;
		auxURL = "";
		if (pAcc){
			auxURL = "&accion=" + pAcc
		}
		document.location.href = "mantenimientoFinalizarOT.asp?idOT=" + pIdOT + auxURL;
	}	
</script>
</head>
<body onLoad="bodyOnLoad()">
<div id="toolBarEquipos"></div>
<form id="frmSel" name="frmSel">
	<input type="hidden" name="accion" id="accion" value="">
	<div class="tableaside size100"> 
		<h3> <%=GF_Traducir("Filtros")%> </h3>
		  
		<div id="searchfilter" class="tableasidecontent">
	        <!--
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("ID:")%> </div>
	        <div class="col16"> <input SIZE="2" type="text" value="<%=txtIdOrder%>" id="txtIdOrder" name="txtIdOrder"> </div>
	        -->
	        <input SIZE="2" type="hidden" value="<%=txtIdOrder%>" id="txtIdOrder" name="txtIdOrder">
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("N�mero:")%> </div>
	        <div class="col16"> <input type="text" size="11" value="<%=txtNroOrder%>" id="txtNroOrder" name="txtNroOrder"> </div>
		       
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("T�tulo:")%> </div>
	        <div class="col16"> <input type="text" size="25" value="<%=txtDsOrder%>" id="txtDsOrder" name="txtDsOrder"> </div>

	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Ver solo planificadas:")%> </div>
	        <div class="col16"> <input style="cursor:pointer;" type="checkbox" value="<%=ESTADO_BAJA%>" <%if cint(txtSoloPlanificadas) = cint(ESTADO_BAJA) then Response.Write "Checked" %> id="txtSoloPlanificadas" name="txtSoloPlanificadas"> </div>

		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Divisi�n:")%> </div>
	        <div class="col16"> 
				<select name="txtIdDivision" id="txtIdDivision">
					<option value=0><%=GF_Traducir("Todas...")%></option>
					<%
					call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLDIVISIONES_GET_BY_LIST", listaDivisionesDefault)
					while not rsList.eof
						%>	
							<option value="<%=rsList("IDDIVISION")%>" <%if cint(txtIdDivision)=cint(rsList("IDDIVISION")) then Response.Write "Selected"%>><%=rsList("DSDIVISION")%></option>
						<%	
						rsList.movenext
					wend	
					%>
				</select>	
	        </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Equipo:")%> </div>
	        <div class="col16">
				<select name="txtIdEquipoActivo" id="txtIdEquipoActivo">
					<option value=0><%=GF_Traducir("Todos...")%></option>
					<%
					call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_PARAMETERS", "0||0||" & txtIdDivision & "||0|| || || ||1|| ")
					while not rsList.eof
						%>	
							<option value="<%=rsList("IDACTIVEEQUIPMENT")%>" <%if cint(txtIdEquipoActivo)=cint(rsList("IDACTIVEEQUIPMENT")) then Response.Write "Selected"%>><%=trim(rsList("CDACTIVATION")) & "/" & trim(rsList("DSACTIVATION"))%></option>
						<%	
						rsList.movenext
					wend	
					%>
				</select>
	        </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Sector:")%> </div>
	        <div class="col16">
				<select name="txtIdSector" id="txtIdSector">
					<option value=0><%=GF_Traducir("Todos...")%></option>
					<%
					call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMSECTOR_GET", "")
					while not rsList.eof
						%>	
							<option value="<%=rsList("IDSECTOR")%>" <%if cint(txtIdSector)=cint(rsList("IDSECTOR")) then Response.Write "Selected"%>><%=rsList("DSSECTOR")%></option>
						<%	
						rsList.movenext
					wend	
					%>
				</select>	
	        </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Solicitante:")%> </div>
	        <div class="col16"> 								
				<input name="solicitante" type="text" id="solicitante" value="<%=solicitante%>" style="width:150px">
				<input type="hidden" name="cdSolicitante" id="cdSolicitante" value="<%=cdSolicitante%>">
			</div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Tipo Mant:")%> </div>
	        <div class="col16"> 
				<select name="txtTipoMantenimiento" id="txtTipoMantenimiento">
					<option value="<%=MAIN_TYPE_ALLS%>"><%=GF_Traducir("Todos...")%></option>
					<option value="<%=MAIN_TYPE_PREVENTIVE%>" <%if txtTipoMantenimiento=MAIN_TYPE_PREVENTIVE then Response.Write "Selected"%>><%=GF_Traducir("Preventivo")%></option>
					<option value="<%=MAIN_TYPE_PREDICTIVE%>" <%if txtTipoMantenimiento=MAIN_TYPE_PREDICTIVE then Response.Write "Selected"%>><%=GF_Traducir("Predictivo")%></option>
					<option value="<%=MAIN_TYPE_CORRECTIVE%>" <%if txtTipoMantenimiento=MAIN_TYPE_CORRECTIVE then Response.Write "Selected"%>><%=GF_Traducir("Correctivo")%></option>
				</select>
	        </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Tipo Orden:")%> </div>
	        <div class="col16">
				<select name="txtTipoOrden" id="txtTipoOrden">
					<option value="<%=ORDER_TYPE_ALLS%>"><%=GF_Traducir("Todos...")%></option>
					<option value="<%=ORDER_TYPE_MECHANICAL%>"	<%if txtTipoOrden=ORDER_TYPE_MECHANICAL then Response.Write "Selected"%>><%=GF_Traducir("Mec�nica")%></option>
					<option value="<%=ORDER_TYPE_ELECRONIC%>"	<%if txtTipoOrden=ORDER_TYPE_ELECRONIC	then Response.Write "Selected"%>><%=GF_Traducir("Electrica")%></option>
					<option value="<%=ORDER_TYPE_CIVIL%>"		<%if txtTipoOrden=ORDER_TYPE_CIVIL		then Response.Write "Selected"%>><%=GF_Traducir("Civil")%></option>
					<option value="<%=ORDER_TYPE_SECURITY%>"	<%if txtTipoOrden=ORDER_TYPE_SECURITY	then Response.Write "Selected"%>><%=GF_Traducir("Seguridad")%></option>
					<option value="<%=ORDER_TYPE_OPERATIVE%>"	<%if txtTipoOrden=ORDER_TYPE_OPERATIVE	then Response.Write "Selected"%>><%=GF_Traducir("Operativa")%></option>
					<option value="<%=ORDER_TYPE_SYSTEM%>"		<%if txtTipoOrden=ORDER_TYPE_SYSTEM		then Response.Write "Selected"%>><%=GF_Traducir("Sistemas")%></option>															
				</select>
	        </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Prog:")%> </div>
	        <div class="col16"> 
   				<table>
					<tr>
						<td>
							<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/calendar-16.png"></a>
						</td>	
						<td>
					<div id="dtProgramadaDiv"><% =dtProgramada %>				
						<%if dtProgramada <> "" then %>
							<a href="javascript:QuitarFecha()"><img src="images/button_cancel.png"></a>
						<%end if%>	
					</div>
					<input type="hidden" id="dtProgramada" name="dtProgramada" value="<%=dtProgramada%>">
				</table>
	         </div>
		        
	        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Empresa Resp:")%> </div>
	        <div class="col16"> 
				<input name="responsable" type="text" id="responsable" value="<%=responsable%>" style="width:150px">
				<input type="hidden" name="cdResponsable" id="cdResponsable" value="<%=cdResponsable%>">
			</div>	

			<div class="col16 reg_header_navdos"> <%=GF_Traducir("Estado:")%> </div>
	        <div class="col16"> 
				<select name="txtCdState" id="txtCdState">
					<option value="<%=STATE_ALLS%>"><%=GF_Traducir("Todos...")%></option>
					<option value="<%=STATE_STAND_BY%>"	<%if txtCdState=STATE_STAND_BY then Response.Write "Selected"%>><%=GF_Traducir(getDsState(STATE_STAND_BY))%></option>
					<option value="<%=STATE_STARTED%>"	<%if txtCdState=STATE_STARTED  then Response.Write "Selected"%>><%=GF_Traducir(getDsState(STATE_STARTED))%></option>
					<option value="<%=STATE_FINISHED%>"	<%if txtCdState=STATE_FINISHED then Response.Write "Selected"%>><%=GF_Traducir(getDsState(STATE_FINISHED))%></option>
					<option value="<%=STATE_CANCELED%>"	<%if txtCdState=STATE_CANCELED then Response.Write "Selected"%>><%=GF_Traducir(getDsState(STATE_CANCELED))%></option>
				</select>
	        </div>
		    
	    	<span class="btnaction"><input type="submit" value="Buscar"></span>
		</div>
	</div>
    
	<div class="col66"></div>

<table class="datagrid" width="90%" align="center">
    <thead>
        <tr>
            <th align="center">.</th>	
			<th class="thiconac" align="center" nowrap>
				<% =GF_TRADUCIR("Nro") %>
				<% call setSortParams(orderByNro,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByNro', 'NROORDER','<%=orderByNro%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByNro" name="orderByNro" value="<%=orderByNro%>">
			</th>				
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("T�tulo") %>
				<% call setSortParams(orderByTit,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByTit', 'DSORDER','<%=orderByTit%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByTit" name="orderByTit" value="<%=orderByTit%>">
			</th>				
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("Equipo") %>		
				<% call setSortParams(orderByEqu,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByEqu', 'CDACTIVATION','<%=orderByEqu%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByEqu" name="orderByEqu" value="<%=orderByEqu%>">
			</th>
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("Sector") %>
				<% call setSortParams(orderBySec,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderBySec', 'DSSECTOR','<%=orderBySec%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderBySec" name="orderBySec" value="<%=orderBySec%>">
			</th>
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("Solicitante") %>
				<% call setSortParams(orderBySol,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderBySol', 'CDAPPLICANT','<%=orderBySol%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderBySol" name="orderBySol" value="<%=orderBySol%>">
			</th>
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("Tipo Mant.") %>	
				<% call setSortParams(orderByTiM,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByTiM', 'MAINTENANCETYPE','<%=orderByTiM%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByTiM" name="orderByTiM" value="<%=orderByTiM%>">
			</th>
			<th align="center" nowrap>
				<% =GF_TRADUCIR("Tipo Orden") %>
				<% call setSortParams(orderByTiO,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByTiO', 'ORDERTYPE','<%=orderByTiO%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByTiO" name="orderByTiO" value="<%=orderByTiO%>">
			</th>
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("M. Obra") %>		
				<% call setSortParams(orderByMan,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByMan', 'DSEMPRESA','<%=orderByMan%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByMan" name="orderByMan" value="<%=orderByMan%>">
			</th>
			<th class="thfecha" align="center" nowrap>	
				<% =GF_TRADUCIR("Fecha Prog.") %>	
				<% call setSortParams(orderByFcP,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByFcP', 'SCHEDULEDDATE','<%=orderByFcP%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByFcP" name="orderByFcP" value="<%=orderByFcP%>">
			</th>				
			<th align="center" nowrap>					
				<% =GF_TRADUCIR("Estado") %>
				<% call setSortParams(orderByEst,myTitle) %>
				<img style="cursor:pointer" title="<%=myTitle%>" onClick="setOrder('orderByEst', 'CDSTATE','<%=orderByEst%>')" src="images\orderlist.png"> 
				<input type="hidden" id="orderByEst" name="orderByEst" value="<%=orderByEst%>">
			<th class="thicon"> <% =GF_TRADUCIR("Acci�n") %> </th>
		</tr>
	<tbody> 	
	<%
	while ((not rs.eof)	and (reg < mostrar))
		reg = reg + 1			
		%>
		<tr>
		    <td class="thicon" align="center"><input type="checkbox" id="chk<% =reg %>" name="chk<% =reg %>" value="<% =rs("IDORDER") %>"></td>
			<td class="thicon" align="center"><% =rs("NROORDER")%></td>
			<td align="left" title="<%=rs("DSORDER")%>">	   
				<% 
				myText = rs("DSORDER")
				if len(myText) > 40 then myText = left(myText,40) & "..."
				Response.Write myText
				%>
			</td>
			<td align="center">	   <% =trim(rs("CDACTIVATION")) & "-" & trim(rs("DSACTIVATION"))%></td>
			<td align="center">	   <% =rs("DSSECTOR") %>								 </td>				
			<td align="center" title="<% =getUserDescription(rs("CDAPPLICANT")) %>">	   <% =rs("CDAPPLICANT") %>			 </td>				
			<td align="center">    <% =getDsMaintenanceType(rs("MAINTENANCETYPE")) %>	 </td>				
			<td align="center">    <% =getDsOrderType(rs("ORDERTYPE")) %>				 </td>				
			<td align="center">    <% =rs("DSEMPRESA") %>								 </td>				
			<td align="center">    <% =GF_FN2DTE(rs("SCHEDULEDDATE")) %>				 </td>
			<td align="center" title="<% =rs("OBSERVATIONS") %>">    <% =getDsState(rs("CDSTATE")) %>						 </td>
			<%	
				verPlanificacion = ""
				if not isNull(rs("IDOTPLAN")) then verPlanificacion = "<img src='images/calendar-16.png' style='cursor: pointer' title='Ir a planificaci�n' onClick=verPlanificacionOT('" & rs("IDORDER") & "')>"
			
				select case CINT(rs("CDSTATE"))
					case STATE_STAND_BY %>
						<td class="thiconac" align="center" nowrap>
						    <% if (isAdmin(rs("IDDIVISION"))) then %> 
							<img src="images/play-16.png" style="cursor: pointer" title="Iniciar" onClick="iniciarOT('<% =rs("IDORDER") %>','<% =rs("NROORDER") %>')">
							<img src="images/edit-16.png" style="cursor: pointer" title="Editar" onClick="editarOT('<% =rs("IDORDER") %>')">
							<% end if %>
							<img src="images/print-16.png" style="cursor: pointer" title="Imprimir" onClick="imprimirOT('<% =rs("IDORDER") %>')">
							<%=verPlanificacion%>
							<img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivo" onClick="adjuntarArchivoOT('<% =rs("IDORDER") %>','<% =rs("NROORDER") %>')">
							<% if (isAdmin(rs("IDDIVISION"))) then %> 
							<img src="images/cross-16.png" style="cursor: pointer" title="Cancelar" onClick="cancelarOT('<% =rs("IDORDER") %>','<% =rs("NROORDER") %>')">
							<% end if %>
						</td>
				<%	case STATE_STARTED %>
						<td class="thiconac" align="center" nowrap>
						    <% if (isAdmin(rs("IDDIVISION"))) then %> 
							<img src="images/finish-16.png" style="cursor: pointer" title="Finalizar" onClick="finalizarOT('<% =rs("IDORDER") %>')">
							<% end if %>
							<img src="images/see-16.png" style="cursor: pointer" title="Visualizar" onClick="finalizarOT('<% =rs("IDORDER") %>','<%=ACCION_VISUALIZAR%>')">
							<img src="images/print-16.png" style="cursor: pointer" title="Imprimir" onClick="imprimirOT('<% =rs("IDORDER") %>')">
							<%=verPlanificacion%>							
							<img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivo" onClick="adjuntarArchivoOT('<% =rs("IDORDER") %>','<% =rs("NROORDER") %>')">
							<% if (isAdmin(rs("IDDIVISION"))) then %> 
							<img src="images/cross-16.png" style="cursor: pointer" title="Cancelar" onClick="cancelarOT('<% =rs("IDORDER") %>','<% =rs("NROORDER") %>')">
							<% end if %>
						</td>
				<%	case STATE_FINISHED,STATE_CANCELED %>
						<td class="thiconac" align="center" nowrap>						
							<img src="images/see-16.png" style="cursor: pointer" title="Visualizar" onClick="finalizarOT('<% =rs("IDORDER") %>','<%=ACCION_VISUALIZAR%>')">
							<img src="images/print-16.png" style="cursor: pointer" title="Imprimir" onClick="imprimirOT('<% =rs("IDORDER") %>')">
							<%=verPlanificacion%>							
						</td>	
			<% end select %>								
		</tr>
		<%
		rs.MoveNext()
	wend
	%>
    <tfoot>
  		<td colspan="13"><div id="paginacion"></div></td>
  	</tfoot>	
  	<%
	if (reg = 0) then		
		%>			
		<tr>
			<td align="center" colspan="13"><% =GF_TRADUCIR("No se encontraron ordenes de trabajo.") %></td>
		</tr>
		<%
	end if 
	%>	
</table>
<input type="hidden" name="myOrder" id="myOrder" value="<%=myOrder%>">	
<input type="hidden" name="ref" id="ref" value="<%=ref%>">	
</form>	
</body>
</html>