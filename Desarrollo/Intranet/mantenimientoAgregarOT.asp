<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/ExternalFunctions.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<% 
Call initAccessInfo(RES_OT_SM)
dim params, myIndice, tituloAux
myIndice = 0
tituloAux = "Crear"
SM_idOrder = GF_PARAMETROS7("idOT",0,6)
call addParam("idOT", SM_idOrder, params)
accion = GF_PARAMETROS7("accion","",6)
call addParam("accion", accion, params)
modoEdicion = GF_PARAMETROS7("modoEdicion",0,6)
call addParam("modoEdicion", modoEdicion, params)

if SM_idOrder > 0 then tituloAux = "Editar"
call readHeaderOT(SM_idOrder)
'SM_type = TASK_TYPE_PROGRAMMING
if accion = ACCION_GRABAR then
	bOk = checkHeaderOT()
	if bok then
		bOk = checkItemsOT()
		if bOk then 
			call saveOT()
			accion = ""
		end if	
	end if
elseif accion = ACCION_CONTROLAR then	
	call checkHeaderOT()
	call checkItemsOT() 
end if
'---------------------------------------
sub leerDesdeDB()
		call readOTDb(idOT)
		call addParam("SM_idOrder", SM_idOrder, params)		
		call addParam("SM_nroOrder", SM_nroOrder, params)		
		call addParam("SM_dsOrder", SM_dsOrder, params)		
		call addParam("SM_idDivision", SM_idDivision, params)		
		call addParam("SM_idActiveEquipment", SM_idActiveEquipment, params)		
		call addParam("SM_scheduledDate", SM_scheduledDate, params)		
		call addParam("SM_startDate", SM_startDate, params)		
		call addParam("SM_finishedDate", SM_finishedDate, params)		
		call addParam("SM_maintenanceType", SM_maintenanceType, params)		
		call addParam("SM_cdState", SM_cdState, params)	
		call addParam("SM_cdApplicant", SM_cdApplicant, params)	
		call addParam("SM_dsApplicant", SM_dsApplicant, params)	
		call addParam("SM_idResponsableCompany", SM_idResponsableCompany, params)	
		call addParam("SM_dsResponsableCompany", SM_dsResponsableCompany, params)			
		call addParam("SM_cdUser", SM_cdUser, params)			
		call addParam("SM_moment", SM_moment, params)					
		call addParam("SM_observations", SM_observations, params)							
end sub	
'---------------------------------------------------------------------------------------------------------
%>
<html>
<head>
<title>Sistema de Mantenimiento - <%=tituloAux%></title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>

<script type="text/javascript">
	function irA(pLink) {
		location.href = pLink;
	}
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 7, "images/");	
		tb.addButtonRETURN("Volver", "irA('mantenimientoAdministrarOTs.asp')");
		tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
		tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
		tb.addButtonREFRESH("Refrescar", "submitInfo()");		
		tb.draw();		
		autoCompleteSolicitante();
		autoCompleteEmpresaResponsable();
		actualizarBudgets(<%=SM_idObra%>,<%=SM_idBudgetArea%>,<%=SM_idBudgetDetalle%>);	
		var myIndex = 0;
		<%
		call initItemsOT()
		while readNextItemOt()%>
			myIndex = parseInt(myIndex) + 1;
			autoCompleteItem(myIndex)    
		<%wend%>		
		
	}
	function submitInfo(pAccion) {		
		document.getElementById("accion").value = pAccion;
		document.getElementById("frmSel").submit();
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
		document.getElementById("dtProgramadaDiv").innerHTML = str;
	    document.getElementById("SM_scheduledDate").value = str;
		if (cal) cal.hide();	
	}	
	
	function autoCompleteEmpresaResponsable()
		{
			$( "#SM_dsResponsableCompany" ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQEmpresas",
					focus: function( event, ui ) {
						$( "#SM_dsResponsableCompany").val(ui.item.dsempresa);
						return false;
					},
					select: function( event, ui ) {
						$( "#SM_dsResponsableCompany"    ).val (ui.item.dsempresa);
						$( "#SM_idResponsableCompany"    ).val (ui.item.idempresa);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#SM_dsResponsableCompany").val ("");
							$( "#SM_idResponsableCompany").val ("");
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
			$( "#SM_dsApplicant" ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQPersonas",
					focus: function( event, ui ) {
						$( "#SM_dsApplicant").val(ui.item.nombre);
						return false;
					},
					select: function( event, ui ) {
						$( "#SM_dsApplicant"    ).val (ui.item.nombre);
						$( "#SM_cdApplicant"    ).val (ui.item.cdusuario);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#SM_dsApplicant").val ("");
							$( "#SM_cdApplicant").val ("");
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
	function autoCompleteItem(pIndex)
		{
			$( "#SM_dsItem" + pIndex ).autocomplete({
					minLength: 2,
					source: "comprasStreamElementos.asp?tipo=JQArticulos",
					focus: function( event, ui ) {
						$( "#SM_dsItem" + pIndex).val(ui.item.dsarticulo);
						return false;
					},
					select: function( event, ui ) {
						$( "#SM_dsItem" + pIndex).val (ui.item.dsarticulo);
						$( "#SM_idItem" + pIndex).val (ui.item.idarticulo);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#SM_dsItem" + pIndex).val ("");
							$( "#SM_idItem" + pIndex).val ("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.idarticulo + " - <font style='font-size:10;'>" + item.dsarticulo + "</font></a>" )
						.appendTo( ul );
				};
		}

	var ch = new channel();	

	var bgClass = false;
    function addTask(pIdOT) {
		var subIndice;
		subIndice = document.getElementById("SM_ActualTask").value;
        var className = "reg_header_navdos";
        $("#TASKS_TABLE")
			.find('tfoot:last')
                .append($('<tr>')
                    .append($('<td>')
                        .append($('<input type=\"hidden\" value=\"0\" id=\"SM_nroTask' + subIndice + '\" name=\"SM_nroTask' + subIndice + '\">')
                            .attr('size', 4))
                    )
                    .append($('<td>')
                        .append($('<input maxlength=\"100\" type=\"text\" id=\"SM_dsTask' + subIndice + '\" name=\"SM_dsTask' + subIndice + '\">')
                            .attr('size', 120))
                    )
                    .append($('<td>')
                        .append($('<input type=\"hidden\" id=\"SM_doneTask' + subIndice + '\" name=\"SM_result' + subIndice + '\">')
                            .attr('size', 4))
                    )
                );
       $('table#TASKS_TABLE tr:last').after($('#ACTION_ROW'));     
       subIndice = Number(subIndice) + 1;   
       document.getElementById("SM_ActualTask").value = subIndice
    }
    function addItem(pIdOT) {
		var subIndice;
		subIndice = document.getElementById("SM_ActualItem").value;
        var className = "reg_header_navdos";
        $("#ITEMS_TABLE")
			.find('tfoot:last')
                .append($('<tr>')
                    .append($('<td>')
                        .append($('<input type=\"hidden\" value=\"0\" id=\"SM_nroItem' + subIndice + '\" name=\"SM_nroItem' + subIndice + '\">')
                            .attr('size', 4)
                        )
                    )
                    .append($('<td>')
                        .append($('<input type=\"hidden\" id=\"SM_idItem' + subIndice + '\" name=\"SM_idItem' + subIndice + '\">')
                            .attr('size', 8)
                        )
                        .append($('<input type=\"text\" id=\"SM_dsItem' + subIndice + '\" name=\"SM_dsItem' + subIndice + '\">')
                            .attr('size', 120)
                        )
                    )
                    .append($('<td align=\"center\">')
                        .append($('<input onkeypress="return controlIngreso(this, event,\'N\')" style="text-align:right;" type=\"text\" id=\"SM_programQuantityItem' + subIndice + '\" name=\"SM_programQuantityItem' + subIndice + '\">')
                            .attr('size', 10)
                        )
                        .append($('<input type=\"hidden\" id=\"SM_realQuantityItem' + subIndice + '\" name=\"SM_realQuantityItem' + subIndice + '\">')
                            .attr('size', 10)
                        )
                    )
                    .append($('<td align=\"center\" >')
                        .append($('<input type=\"hidden\" id=\"SM_idPmItem' + subIndice + '\" name=\"SM_idPmItem' + subIndice + '\">')
                            .attr('size', 4)
                        )
                    )
                );
       $('table#ITEMS_TABLE tr:last').after($('#ACTION_ROW_ITEM')); 
       autoCompleteItem(subIndice)    
       subIndice = Number(subIndice) + 1;   
       document.getElementById("SM_ActualItem").value = subIndice
    }        
	function deleteOtTask(pIdOT, pNroTask){
		ch.bind("mantenimientoOtABMAJAX.asp?idOrder=" + pIdOT + "&nroTask=" + pNroTask + "&tipoOpr=DT", "deleteOtTask_Callback()");
		ch.send();			
	}
	function deleteOtTask_Callback(){
		submitInfo();
	}        
	function deleteOtItem(pIdOT, pNroItem){
		ch.bind("mantenimientoOtABMAJAX.asp?idOrder=" + pIdOT + "&nroItem=" + pNroItem + "&tipoOpr=DI", "deleteOtTask_Callback()");
		ch.send();			
	}
	function actualizarBudgets(idObra, idBudgetArea, idBudgetDetalle){
		myReadOnly = 0;
		ch.bind("almacenObtenerBudget.asp?idObra=" + idObra + "&idBudgetArea=" + idBudgetArea + "&idBudgetDetalle=" + idBudgetDetalle + "&readOnly=" + myReadOnly + "&accion=<%=ACCION_PROCESAR%>", "actualizarBudgetsCallback(" + idObra + ")");
		ch.send();	
	}        	
	function actualizarBudgetsCallback(idObra){
		document.getElementById("secBudgetDiv").innerHTML = ch.response(); 	
	}	
	function readBudgetArea() {
		document.getElementById('idBudgetArea').value=$("#idBudgetDetalle option:selected").attr("alt");
	}	
</script>
</head>
<%'Response.Buffer = False%>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="mantenimientoAgregarOT.asp">

<div id="toolbar"></div>

<div class="tableaside size100">
<%=showMessages()%>
	<%
	myText = "Nueva"
	if clng(SM_idOrder) <> 0 then myText = "Editar"
	%>
	<h3> <%=GF_Traducir(myText & " Orden de Trabajo")%> </h3>
  
	<div class="tableasidecontent">
		<input type="hidden" name="idOT" id="idOT" value="<%=SM_idOrder%>">
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Número")%> </div>
        <div class="col26"> 
			<%if SM_nroOrder <> "" then
				Response.Write SM_nroOrder	
			else%>
				<i><%=GF_Traducir("Pendiente")%></i>
			<%end if%>		
			<input type="hidden" name="SM_nroOrder" id="SM_nroOrder" value="<%=SM_nroOrder%>"> 
		</div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Título")%> </div>
        <div class="col26"> 
			<input type="text" name="SM_dsOrder" id="SM_dsOrder" value="<%=SM_dsOrder%>" maxlength="1000" size="30"> 
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("División")%> </div>
        <div class="col26"> 
			<% if cint(modoEdicion) = 1 then 
					Response.Write getDivisionDS(SM_idDivision)
				%>
				<input type="hidden" value="<%=SM_idDivision%>" name="SM_idDivision" id="SM_idDivision">
			<%else%>
			<select name="SM_idDivision" id="SM_idDivision" onChange="submitInfo('<%=ACCION_CONTROLAR%>')">
				<option value=0><%=GF_Traducir("Seleccione...")%></option>
				<%
				call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLDIVISIONES_GET_BY_LIST", getListaCargosAdmin())
				while not rsList.eof
					%>	
						<option value="<%=rsList("IDDIVISION")%>" <%if cint(SM_idDivision)=cint(rsList("IDDIVISION")) then Response.Write "Selected"%>><%=rsList("DSDIVISION")%></option>
					<%	
					rsList.movenext
				wend	
				%>
			</select>	
			<%end if%>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Equipo")%> </div>
        <div class="col26"> 
			<% if SM_idDivision <> 0 then %>
				<select name="SM_idActiveEquipment" id="SM_idActiveEquipment" >
					<option value=0><%=GF_Traducir("Seleccione...")%></option>
					<%
					call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_PARAMETERS", "0||0||" & SM_idDivision & "||0|| || || ||1|| ")
					while not rsList.eof
						%>	
							<option value="<%=rsList("IDACTIVEEQUIPMENT")%>" <%if cint(SM_idActiveEquipment)=cint(rsList("IDACTIVEEQUIPMENT")) then Response.Write "Selected"%>><%=trim(rsList("CDACTIVATION")) & " - " & trim(rsList("DSACTIVATION"))%></option>
						<%	
						rsList.movenext
					wend	
					%>
				</select>
			<%else
				Response.Write "<i>Seleccione una division</i>"
			  end if
   		    %>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Tipo Mant.")%> </div>
        <div class="col26">
			<select name="SM_maintenanceType" id="SM_maintenanceType">
				<option value="<%=MAIN_TYPE_ALLS%>"><%=GF_Traducir("Seleccione...")%></option>
				<option value="<%=MAIN_TYPE_PREVENTIVE%>" <%if SM_maintenanceType=MAIN_TYPE_PREVENTIVE then Response.Write "Selected"%>><%=GF_Traducir("Preventivo")%></option>
				<option value="<%=MAIN_TYPE_PREDICTIVE%>" <%if SM_maintenanceType=MAIN_TYPE_PREDICTIVE then Response.Write "Selected"%>><%=GF_Traducir("Predictivo")%></option>
				<option value="<%=MAIN_TYPE_CORRECTIVE%>" <%if SM_maintenanceType=MAIN_TYPE_CORRECTIVE then Response.Write "Selected"%>><%=GF_Traducir("Correctivo")%></option>
			</select>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Tipo OT")%> </div>
        <div class="col26"> 
			<select name="SM_orderType" id="SM_orderType">
				<option value="<%=ORDER_TYPE_ALLS%>"><%=GF_Traducir("Seleccione...")%></option>
				<option value="<%=ORDER_TYPE_MECHANICAL%>"	<%if SM_orderType=ORDER_TYPE_MECHANICAL then Response.Write "Selected"%>><%=GF_Traducir("Mecánica")%></option>
				<option value="<%=ORDER_TYPE_ELECRONIC%>"	<%if SM_orderType=ORDER_TYPE_ELECRONIC	then Response.Write "Selected"%>><%=GF_Traducir("Electrica")%></option>
				<option value="<%=ORDER_TYPE_CIVIL%>"		<%if SM_orderType=ORDER_TYPE_CIVIL		then Response.Write "Selected"%>><%=GF_Traducir("Civil")%></option>
				<option value="<%=ORDER_TYPE_SECURITY%>"	<%if SM_orderType=ORDER_TYPE_SECURITY	then Response.Write "Selected"%>><%=GF_Traducir("Seguridad")%></option>
				<option value="<%=ORDER_TYPE_OPERATIVE%>"	<%if SM_orderType=ORDER_TYPE_OPERATIVE	then Response.Write "Selected"%>><%=GF_Traducir("Operativa")%></option>
				<option value="<%=ORDER_TYPE_SYSTEM%>"		<%if SM_orderType=ORDER_TYPE_SYSTEM		then Response.Write "Selected"%>><%=GF_Traducir("Sistemas")%></option>															
			</select>
        </div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Prog.")%> </div>
        <div class="col26"> 
			<table>
				<tr>
					<td>
						<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/calendar-16.png"></a>
					</td>	
					<td>
				<div id="dtProgramadaDiv">
					<% 
					if not isdate(SM_scheduledDate) then SM_scheduledDate = GF_FN2DTE(SM_scheduledDate)
					Response.write SM_scheduledDate 
					%>
				</div>
				<input type="hidden" id="SM_scheduledDate" name="SM_scheduledDate" value="<%=SM_scheduledDate%>">
				<input type="hidden" id="SM_startDate" name="SM_startDate" value="<%=SM_startDate%>">
				<input type="hidden" id="SM_finishedDate" name="SM_finishedDate" value="<%=SM_finishedDate%>">
			</table>
		</div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Frecuencia")%> </div>
        <div class="col26"> 
			<input type="text" maxlength="3" onkeypress="return controlIngreso(this, event,'N')" size="1" id="SM_OTFrequencyQuantity" name="SM_OTFrequencyQuantity" value="<%=SM_OTFrequencyQuantity%>">
			<select name="SM_OTFrequencyUnit" id="SM_OTFrequencyUnit">
				<option value="<%=ORDER_FREQ_UNIQUE%>"><%=GF_Traducir("Única")%></option>
				<option value="<%=ORDER_FREQ_DAY%>"		<%if SM_OTFrequencyUnit=ORDER_FREQ_DAY		then Response.Write "Selected"%>><%=GF_Traducir("Día/s")%></option>
				<option value="<%=ORDER_FREQ_WEEK%>"	<%if SM_OTFrequencyUnit=ORDER_FREQ_WEEK		then Response.Write "Selected"%>><%=GF_Traducir("Semana/s")%></option>
				<option value="<%=ORDER_FREQ_MONTH%>"	<%if SM_OTFrequencyUnit=ORDER_FREQ_MONTH	then Response.Write "Selected"%>><%=GF_Traducir("Mes/es")%></option>
				<option value="<%=ORDER_FREQ_YEAR%>"	<%if SM_OTFrequencyUnit=ORDER_FREQ_YEAR		then Response.Write "Selected"%>><%=GF_Traducir("Año/s")%></option>
			</select>
        </div>        
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Solicitante")%> </div>
        <div class="col26"> 
			<input name="SM_dsApplicant" type="text" id="SM_dsApplicant" value="<%=SM_dsApplicant%>" style="width:150px">
			<input type="hidden" name="SM_cdApplicant" id="SM_cdApplicant" value="<%=SM_cdApplicant%>">
		</div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Mano Obra")%> </div>
        <div class="col26">
			<input name="SM_dsResponsableCompany" type="text" id="SM_dsResponsableCompany" value="<%=SM_dsResponsableCompany%>" style="width:200px">
			<input type="hidden" name="SM_idResponsableCompany" id="SM_idResponsableCompany" value="<%=SM_idResponsableCompany%>">
		</div>
        
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Part. Pres.")%> </div>
        <div class="col46" style="overflow:visible; height:auto;"> 
				<% 	Set rsObras = obtenerListaObras("", "", "", SM_idDivision,OBRA_ACTIVA)	%>
					<select id="idObra" name="idObra" onChange="actualizarBudgets(this.value,0,0)">
						<option value="0">- <% =GF_TRADUCIR("Seleccione") %>
					<%	while (not rsObras.eof)	%>
							<option value="<% =rsObras("IDOBRA") %>" <% if (rsObras("IDOBRA") = SM_idObra) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %></option>
					<%		rsObras.MoveNext()
						wend 	%>
					</select>
				&nbsp;&nbsp;&nbsp;<span id="secBudgetDiv"></span>
				<input type="hidden" name="SM_cdState" id="SM_cdState" value="<%=SM_cdState%>">
				<input type="hidden" name="SM_cdUser" id="SM_cdUser" value="<%=SM_cdUser%>">
				<input type="hidden" name="SM_moment" id="SM_moment" value="<%=SM_moment%>">
				<input type="hidden" name="SM_observations" id="SM_observations" value="<%=SM_observations%>">
		</div>
	</div>	
</div>

<%if CLNG(SM_idOrder) <> 0 then
	set files = getOTFiles(SM_idOrder)
	if not (files.EoF) then%>
		<div class="tableaside size100">
			<h3> <%=GF_Traducir("Archivos asociados a esta Orden")%> </h3>
				<table class="datagrid" width="90%" align="center">
					<thead>
						<tr>
							<th class="thicon"><%=GF_Traducir("Tipo")%> </th>
							<th><%=GF_Traducir("Nombre del archivo")%></th>
							<th class="thiconac" width="80"> - </th>
					    </tr>
					</thead>
					<tbody>	
					<%while not files.EoF %>
						<tr>
							<td class="thicon">
								<%=getImageByExt(files("EXT"))%>
							</td>	
							<td>
								<%=files("NAME") & "." & files("EXT")%>
							</td>
							<td class="thiconac">
								<a target='_blank' href='comprasOpenArchivo.asp?id=<%=files("ID")%>&secuencia=<%=files("FILENO")%>&type=SM-OT-OPEN'>
									<img width="16" height="16" src="images/download-16.png" title="Descargar Archivo">
								</a>
								<%if isAdminInAny  then%>
									<img src="images/cross-16.png" style="cursor:pointer;" onclick="deleteFile('<%=files("ID")%>','<%=files("FILENO")%>')" title="Eliminar Archivo">
								<%else
									Response.Write "."
								end if%>
							</td>
						</tr>
					<%files.MoveNext
					wend%>
					</tbody>
				</table>		
			</div>	
	<%end if
end if	%>
	
<div class="tableaside size100">

	<h3> <%=GF_Traducir("Descripción del trabajo a realizar")%> </h3>
  
	<div class="tableasidecontent">
		<table id="TASKS_TABLE" border=0 align="center" class="datagrid" width="90%">
            <thead>
                <tr>
					<th class="thicon" colspan="3"> <%=GF_Traducir("Tareas")%> </th>
			    </tr>
		<%if CLNG(SM_idOrder) = 0 then%>
			</thead>
				<tbody>
					<tr>
						<td colspan="3"><%=GF_Traducir("Disponible una vez guardada la cabecera de la OT.")%></td>
					</tr>	
				</tbody>	
		<%else%>
			    <tr>
			        <td class="thicon"> <%=GF_Traducir("Nro")%></td>
			        <td align="center"> <%=GF_Traducir("Descripción")%></td>
			        <td class="thiconac"> <%=GF_Traducir("Hecha")%></td>
			    </tr>
			</thead>
			<tbody>
			<%
			call initTasksOT()
			while readNextTaskOt()
					%>
					<tr>
						<td class="thicon" align="center">
							<%
								if SM_nroTask = 0 then 
									Response.Write "."
								else
									Response.write SM_nroTask
								end if
							%>
							<input type="hidden" id="SM_nroTask<%=SM_ActualTask%>" name="SM_nroTask<%=SM_ActualTask%>" size=4 value="<%=SM_nroTask%>">
						</td>
						<td >
							<input type="text" id="SM_dsTask<%=SM_ActualTask%>" name="SM_dsTask<%=SM_ActualTask%>" size=120 value="<%=SM_dsTask%>">
							<input type="hidden" id="SM_doneTask<%=SM_ActualTask%>" name="SM_doneTask<%=SM_ActualTask%>" size=4 value="<%=SM_doneTask%>">
						</td>						
						<td class="thicon" align="center">
							<% if SM_nroTask <> 0 then %>
								<img src="images/cross-16.png" title="<%=GF_Traducir("Eliminar Tarea")%>" style="cursor:pointer;" onClick="deleteOtTask('<%=SM_idOrder%>','<%=SM_nroTask%>','<%=SM_typeTask%>')">
							<% end if %>
						</td>
					</tr>	
					<%
			wend
			%>
			</tbody>
			<tfoot>
				<tr id="ACTION_ROW">
					<td colspan="4" align="right" >
				        <a class="btnmore" href="javascript:addTask('<%=idOT%>')"><img src="images/plus-16.png"> Agregar Tarea</a>
					</td>
				</tr>	
			</tfoot>                
		<%end if%>                
		</table>
		
		<div class="col66"></div>
		
		<table class="datagrid" id="ITEMS_TABLE" border=0 align="center" width="90%">
			<thead>
                <tr>
					<th class="thicon" colspan="5"> <%=GF_Traducir("Piezas/Repuestos")%> </th>
                </tr>
		<%if CLNG(SM_idOrder) = 0 then%>
			</thead>
			<tbody>				
				<tr>
					<td colspan="5"><%=GF_Traducir("Disponible una vez guardada la cabecera de la OT.")%></td>
				</tr>	
			</tbody>				
		<%else%>
				<tr>
					<td class="thicon" align="center"><%=GF_Traducir("Nro")%></td>
					<td align="center" ><%=GF_Traducir("Descripción")%></td>
					<td class="thicon" align="center"><%=GF_Traducir("Cantidad")%></td>
					<td class="thicon" align="center"><%=GF_Traducir("PM")%></td>
					<td class="thiconac" align="center">-</td>
				</tr>	
		</thead>		
		<tbody>    
					<%
					call initItemsOT()
					while readNextItemOt()
							%>
							<tr>
								<td class="thicon" align="center">
									<%
										if SM_nroItem = 0 then 
											Response.Write "."
										else
											Response.write SM_nroItem
										end if
									%>
									<input type="hidden" id="SM_nroItem<%=SM_ActualItem%>" name="SM_nroItem<%=SM_ActualItem%>" size=4 value="<%=SM_nroItem%>">
								</td>
								<td>
									<input type="hidden" id="SM_idItem<%=SM_ActualItem%>" name="SM_idItem<%=SM_ActualItem%>" size=8 value="<%=SM_idItem%>">
									<input type="text" id="SM_dsItem<%=SM_ActualItem%>" name="SM_dsItem<%=SM_ActualItem%>" size=120 value="<%=SM_dsItem%>">
								</td>
								<td class="thicon" align="center">	
									<input style="text-align:right;" type="text" id="SM_programQuantityItem<%=SM_ActualItem%>" name="SM_programQuantityItem<%=SM_ActualItem%>" size=10 value="<%=SM_programQuantityItem%>"  onkeypress="return controlIngreso(this, event,'N')">
									<input style="text-align:right;" type="hidden" id="SM_realQuantityItem<%=SM_ActualItem%>" name="SM_realQuantityItem<%=SM_ActualItem%>" size=10 value="<%=SM_realQuantityItem%>">
								</td>
								<td class="thicon" align="center">	
									<input type="hidden" id="SM_idPmItem<%=SM_ActualItem%>" name="SM_idPmItem<%=SM_ActualItem%>" size=10 value="<%=SM_idPmItem%>">
								</td>	
								<td class="thicon">	
									
									<% if SM_nroItem <> 0 then %>					
										<img src="images/cross-16.png" title="<%=GF_Traducir("Eliminar Repuesto")%>" style="cursor:pointer;" onClick="deleteOtItem('<%=SM_idOrder%>','<%=SM_nroItem%>')">
									<% end if %>					
								</td>
							</tr>	
							<%
					wend
					%>
			</tbody>    
			<tfoot>
					<tr id="ACTION_ROW_ITEM">
						<td colspan="6" align="right" >
					        <a class="btnmore" href="javascript:addItem('<%=idOT%>')"><img src="images/plus-16.png"> Agregar Repuesto</a>
					    </td>
					</tr>	
			</tfoot>
				<%end if%>
		</table>
	</div>
</div>	

<input type="HIDDEN" name="accion" id="accion" value="<%=accion%>">
<input type="HIDDEN" name="modoEdicion" id="modoEdicion" value="<%=modoEdicion%>">

<input type="HIDDEN" name="SM_ActualTask" id="SM_ActualTask" value="<%=SM_ActualTask%>">
<input type="HIDDEN" name="SM_ActualItem" id="SM_ActualItem" value="<%=SM_ActualItem%>">
</form>
</body>
</html>