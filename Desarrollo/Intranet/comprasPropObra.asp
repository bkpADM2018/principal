<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHTML.asp"-->
<%
Const MOSTRAR_LINEA = 0
Const OCULTAR_LINEA = 1
Const ID_MAXIMO = 1000
Call comprasControlAccesoCM(RES_OBR)
'--------------------------------------------------------------------------------------------
Function controlar(pIdObra, pCdObra, pDsObra, pCdResponsable, pIdDivision, fechaInicio, fechaFin, fechaAjustada, pFileAprobacion, pTipoGasto)
	Dim strSQL, rs, conn, dsDivision, retControlFecha, ret
	ret = RESPUESTA_OK	
	if (pCdObra = "") then
		ret = CODIGO_VACIO
	else		
		if (pIdDivision <> SIN_DIVISION) then
			strSQL="Select * from TBLDATOSOBRAS where CDOBRA='" & Trim(pCdObra) & "'"
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if ((not rs.eof) and (pIdObra = 0)) then 		
				ret = CODIGO_EXISTE			
			else		
				'Se controlan las fechas
				if (fechaInicio = "") then ret=FECHA_INICIO_INCORRECTA
				if (fechaFin = "") then ret=FECHA_FIN_INCORRECTA			
				if (ret = RESPUESTA_OK) then						
					if (GF_CONTROL_PERIODO_2(fechaInicio, fechaFin) <> 0) then ret=PERIODO_ERRONEO
				end if			
				if (fechaAjustada <> "") then											
					if (GF_CONTROL_PERIODO_2(fechaInicio, fechaAjustada) <> 0) then ret=PERIODO_ERRONEO
				end if
				if (pCdResponsable = "") then ret=RESPONSABLE_NO_EXISTE
			end if		
		else
			ret = DIVISION_NO_EXISTE
		end if
	end if
controlar = ret
End Function
'--------------------------------------------------------------------------------------------
Function accionGrabar(ByRef pIdObra, pCdObra, pDsObra, pCdResponsable, pIdDivision, pFechaInicio, pFechaFin, pFechaAjustada, pFileAprobacion, pInversion,pTipoGasto)
	Dim strSQL, rs, conn, rsDat
	if (pFechaAjustada = "") then pFechaAjustada = 0
	
	if (pIdObra = 0) then
		'Es una  nueva
		strSQL="Insert into TBLDATOSOBRAS(CDOBRA, DSOBRA, IDDIVISION, FECHAINICIO, FECHAFIN, FECHAAJUSTADA,  FECHABUDGET, ESINVERSION, CDRESPONSABLE,TIPOGASTO)"
		strSQL= strSQL & " values('" & pCdObra & "', '" & pDsObra & "', " & pIdDivision & ", " & pFechaInicio & ", " & pFechaFin & ", " & pFechaAjustada & ", 0, '" & pInversion & "', '" & pCdResponsable & "','"&pTipoGasto&"')"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		'OBTENGO EL RECIENTE IDOBRA CREADO EN LA TABLA DATOSOBRAS
	    strSQL= "SELECT MAX(IDOBRA) AS IDOBRA FROM TBLDATOSOBRAS"
	    Call executeQueryDB(DBSITE_SQL_INTRA, rsDat, "OPEN", strSQL)
	    if( not rsDat.eof)then pIdObra = rsDat("IDOBRA")
	else
		'Es una modificacion
		strSQL="Update TBLDATOSOBRAS Set CDOBRA='" & pCdObra & "', DSOBRA='" & pDsObra & "', CDRESPONSABLE='" & pCdResponsable & "', IDDIVISION=" & pIdDivision & ", FECHAINICIO=" & pFechaInicio & ", FECHAAJUSTADA=" & pFechaAjustada & ", FECHAFIN=" & pFechaFin & ", ESINVERSION='" & pInversion & "' ,TIPOGASTO='"&pTipoGasto&"'"
		strSQL = strSQL & " where IDOBRA=" & pIdObra
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if		
	accionGrabar = true
End Function

'--------------------------------------------------------------------------------------------         
Function accionConsulta(idObra, ByRef cdObra, ByRef dsObra, ByRef cdResponsable, ByRef idDivision, ByRef fechaInicio, ByRef fechaFin, ByRef fechaAjustada, ByRef fileAprobacion, ByRef fechaBudget, ByRef pInversion, byRef pTipoGasto)
	Dim strSQL, rs, conn

	strSQL = "Select * from TBLDATOSOBRAS where IDOBRA=" & idObra
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		idObra = rs("IDOBRA")
		cdObra = rs("CDOBRA")
		dsObra = Trim(rs("DSOBRA"))
		cdResponsable = rs("CDRESPONSABLE")
		idDivision = rs("IDDIVISION")
		fechaInicio = GF_FN2DTE(rs("FECHAINICIO"))
		fechaFin = GF_FN2DTE(rs("FECHAFIN"))		
		fechaAjustada = GF_FN2DTE(rs("FECHAAJUSTADA"))	
		fechaBudget = rs("FECHABUDGET")
		if (fechaAjustada = "0") then fechaAjustada = ""		
		pInversion = rs("ESINVERSION")		
		pTipoGasto = rs("TIPOGASTO")
	end if		
End Function
'--------------------------------------------------------------------------------------------
Function hayConsumos(idObra)
	Dim strSQL, cn, rs, ret
	
	'Si no hay obra (es nueva) no hay consumos.
	ret=false
	if (idObra > 0) then
		'Si  estoy modificando, asumo que puede haber consumos.
		ret = true
		strSQL = "Select IDCOTIZACION from TBLCTZCABECERA where IDOBRA=" & idObra
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (rs.eof) then		
			strSQL = "Select IDVALE from TBLVALESCABECERA where IDOBRA=" & idObra
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (rs.eof) then ret = false 'Verificado! No hay consumos.
		end if
	end if
	hayConsumos = ret
	
End Function
'--------------------------------------------------------------------------------------------
Function accionGrabarAreaDetalle(pIdObraNueva, pCdObraTemplate, pSoloRepetitivos)
	Dim strSQL,rs,rsBud,rsDat, myObra
	
	'OBTENGO LAS AREAS Y DETALLES DE LA PARTIDA A COPIAR
	strSQL = "SELECT IDAREA,IDDETALLE,DSBUDGET, CDCUENTA, CCOSTOS FROM TBLBUDGETOBRAS WHERE IDOBRA in (Select IDOBRA from TBLDATOSOBRAS where CDOBRA='"& pCdObraTemplate & "')"
	if (pSoloRepetitivos) then
	    strSQL = strSQL	&"  AND (IDAREA < "&ID_MAXIMO&" AND IDDETALLE < "&ID_MAXIMO&")"
    end if	    
    Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	myTipoCambio = getTipoCambio(MONEDA_DOLAR, "")
	while not rs.eof
		'A SU VEZ CADA AREA Y DETALLE LO GUARDO CON LA OBRA NUEVA,  ADEMAS GUARDO LA DESCRIPCION ASEGURANDO QUE CUANDO SE MODIFIQUE
		'UN ITEM NO AFECTE A LA PARTIDA PRESUPUESTARIA.
		strSQL = "INSERT INTO TBLBUDGETOBRAS (IDOBRA,IDAREA,IDDETALLE,DSBUDGET, TIPOCAMBIO, CDCUENTA, CCOSTOS,MOMENTO) "
		strSQL = strSQL & " VALUES("&pIdObraNueva&","&rs("IDAREA")&","&rs("IDDETALLE")&",'"&rs("DSBUDGET")&"'," & myTipoCambio & ", '"&rs("CDCUENTA")&"','"&rs("CCOSTOS")&"', "&session("mmtoSistema")&")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsBud, "EXEC", strSQL)
		rs.MoveNext()
	wend	
End Function
'--------------------------------------------------------------------------------------------
Function puedeModificar(idObra, fechaInicio, idDivision)
	puedeModificar = false
	if (idObra = 0) then 
		puedeModificar = true
	else
		if (not isAuditor(idDivision)) then		
			if (GF_CONTROL_PERIODO_2(GF_FN2DTE(left(session("MmtoDato"),8)), fechaInicio) = 0) then
				puedeModificar = true
			end if
		end if
	end if
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idObra, cdObra, dsObra, idDivision, fileAprobacion
Dim rsDivision, conn,strSQL, path, pathWeb, dsResponsable, cdResponsable
dim colorP, myColor1, myColor2, cont, fechaBudget, esInversion, tipoGasto
dim CdObraBase,DsObraBase, pSoloRepetitivos

idObra = GF_PARAMETROS7("idObra",0,6)
cdObra = UCase(GF_PARAMETROS7("codigo","",6))
dsObra = UCase(GF_PARAMETROS7("descripcion","",6))
idDivision = GF_PARAMETROS7("idDivision",0,6)	
fechaInicio = GF_PARAMETROS7("idate","",6)
fechaFin = GF_PARAMETROS7("fdate","",6)
fechaAjustada = GF_PARAMETROS7("adate","",6)
cdResponsable = GF_PARAMETROS7("cdResponsable","",6)
fileAprobacion = GF_PARAMETROS7("apFile","",6)
tipoGasto = GF_PARAMETROS7("tipoGasto","",6)

'Se prepara el indicador para saber si se graban los items repetitivos (idarea y/o iddetalle <1000) o todos.
pSoloRepetitivos = False
if (GF_PARAMETROS7("soloRepetitivos", 0, 6) = 0) then pSoloRepetitivos = True

CdObraBase = GF_PARAMETROS7("CdObraBase","",6)
DsObraBase = GF_PARAMETROS7("DsObraBase","",6)

if (tipoGasto = OBRA_TIPO_INVERSION) then 
	esInversion = OBRA_INVERSION
else
	esInversion = OBRA_MANTENIMIENTO
end if

accion = GF_PARAMETROS7("accion","",6)

if (idObra <> 0) then
	if (not checkControlObra(idObra)) then
		response.redirect "comprasAccesoDenegado.asp"
	end if
end if

myColor1 = "#d3d3d3"
myColor2 = "#ffffff"
Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then	
	errMsg = controlar(idObra, cdObra, dsObra, cdResponsable, idDivision, fechaInicio, fechaFin, fechaAjustada, fileAprobacion, tipoGasto)
	if (errMsg = RESPUESTA_OK)then		
		Call accionGrabar(idObra, cdObra, dsObra, cdResponsable, idDivision, GF_DTE2FN(fechaInicio), GF_DTE2FN(fechaFin), GF_DTE2FN(fechaAjustada), fileAprobacion, esInversion,tipoGasto)		
		if(CdObraBase <> "")then Call accionGrabarAreaDetalle(idObra, CdObraBase, pSoloRepetitivos)
		accion = ACCION_CERRAR
	else
		setError(errMsg)
	end if
else	
	Call accionConsulta(idObra, cdObra, dsObra, cdResponsable, idDivision, fechaInicio, fechaFin, fechaAjustada, fileAprobacion, fechaBudget, esInversion,tipoGasto)
end if
dsResponsable = getUserDescription(cdResponsable)
esModificable = puedeModificar(idObra, fechaInicio, idDivision)
empezoObra = hayConsumos(idObra)

%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
<link rel="stylesheet" href="css/jquery.fileupload-ui.css"	 type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">


<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<!--Scripts para el Upload-->
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/JQueryUpload2.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery.fileupload.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery.fileupload-ui.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<!-------------------------->
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

<script type="text/javascript">
var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
var refPopUpObra;
var calendar;
var myTipoGasto = "";						
							
function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}
function submitInfo() {		
	document.getElementById("frmSel").submit();
}
function canSubmit() {		
	submitInfo();
}

function seleccionarResponsable(ms) {
	var desc = ms.getSelectedItem();
	if (desc.indexOf('-') != -1) {
		var arr = desc.split('-');
		document.getElementById("cdResponsable").value = arr[0];
		ms.setValue(arr[1]);
	} else {
		if (desc == "") document.getElementById("cdResponsable").value = "";							
	}		
}
	
function SeleccionarCalInicio(cal, date) {
	var str= new String(date);		
	document.getElementById("idateDiv").innerHTML = str;
	document.getElementById("idate").value = str;
	if (cal) cal.hide();	
}

function SeleccionarCalFin(cal, date) {
	var str= new String(date);		
	document.getElementById("fdateDiv").innerHTML = str;
	document.getElementById("fdate").value = str;
	if (cal) cal.hide();	
}

function SeleccionarCalAjustada(cal, date) {
	var str= new String(date);		
	document.getElementById("adateDiv").innerHTML = str;
	document.getElementById("adate").value = str;
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

function createBudget() {
	window.open('comprasBudgetObra.asp?idObra=<% =idObra %>');
	var img = document.getElementById("imgBudget")
	img.src="images/prestamo.gif";
	if (isFirefox) {
		img.setAttribute('onclick', "reloadPage()");
	} else {
		img['onclick'] = new Function("reloadPage()");
	}	
}

function reloadPage() {
	location.href = "comprasPropObra.asp?idObra=<% =idObra %>";
}

function obraOnLoad() {
	var elem = document.getElementById("codigo");
	<%	if (esModificable) then %>
	if (elem.type != "hidden")
		elem.focus();
	else
		document.getElementById("descripcion").focus();
	
	var msResponsable = new MagicSearch("", "divResponsable", 30, 2, "comprasStreamElementos.asp?tipo=personas");
	msResponsable.setToken(";");
	msResponsable.onBlur = seleccionarResponsable;
	msResponsable.setValue('<% =dsResponsable %>');	
	<% end if %>	
	refPopUpObra = getObjPopUp('popupObra');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpObra.hide();
	<% end if %>		

	pngfix();
}	

<%	if (esModificable) then %>
	$(function() {
		$( "#DsObraBase" ).autocomplete({
		minLength: 2,				
		/* SE SEPARO LA FORMA DE SETEAR EL SOURCE, DE ESTA FORMA SE PUEDE MANDAR COMO PARAMETRO EL VALOR DE UN CONTROL(SELECT) DE  
		 * FORMA ESPONTANEA, PUDIENDO FILTRAR LA BUSQUEDA DE LA PARTIDA A COPIAR POR EL TIPO DE GASTO QUE SE SELECCIONE
		 */
		source: function(request,response){
			$.ajax({				
				url: "comprasStreamElementos.asp",
				dataType: "json",
			data: {
				/* ACA VAN DECLARADOS LOS NOMBRES DE LOS PARAMETROS CON SUS RESPECTIVOS VALORES */
				term : request.term,
				Tipo : "JQObras",
				//TipoGasto : document.getElementById("tipogasto").options[document.getElementById("tipogasto").selectedIndex].value
				 },
		    success: function(data) {				
				response(data);
				}
			});	
		},		
		focus: function( event, ui ) {
				$( "#DsObraBase").val(ui.item.dsobra);
				return false;
			},
		select: function( event, ui ) {
				$( "#DsObraBase").val (ui.item.dsobra);
				$( "#CdObraBase").val (ui.item.cdobra );
				return false;
			}		
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
			.data( "item.autocomplete", item )
			.append( "<a>" + item.cdobra + " - <font style='font-size:10;'>" + item.dsobra + "</font></a>" )
			.appendTo( ul );
		};
	});
 <%	end if %>


</script>
</head>
<body onLoad="obraOnLoad()">
<form name="frmSel" id="frmSel" method="post" action="comprasPropObra.asp">
<table width="100%" border="0">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/OBR-32x32.png"> <% =GF_TRADUCIR("Partida Presupuestaria") %></td>
	</tr>
	<tr>
		<td colspan="4"><% call showErrors() %></td>
	</tr>	
	<tr>
		<td></td>
		<td>			
			<table width="100%" border="0">
				<tr>
					<td width="30%" class="reg_header"><% =GF_TRADUCIR("Codigo") %></td>
					<td colspan="2">
					<%	if (not empezoObra) then %>
							<input type="text" id="codigo" name="codigo" maxlength="10" size="10" value="<% =cdObra %>" onblur="codigoOnBlur(this)" onkeypress="return controlSalto(this, event)"></td>
					<%  else %>
							<b><% =cdObra %></b>							
							<input type="hidden" id="codigo" name="codigo" value="<% =cdObra %>">
					<%	end if 	%>
					</td> 
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>					
					<td colspan="2">
					<%	if (not empezoObra) then %>
						<input type="text" id="descripcion" name="descripcion" maxlength="50" size="30" value="<% =dsObra %>" onkeypress="return controlSalto(this, event)"></td>
					<%  else %> 
							<b><% =dsObra %></b> 							
							<input type="hidden" id="descripcion" name="descripcion" value="<% =dsObra %>">
					<%	end if 	%>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Responsable") %></td>
					<td colspan="2">
						<%	if (esModificable) then %>
								<div id="divResponsable"></div>																		
						
						<%  else
								response.write dsResponsable 											 
							end if 	%>
						<input type="hidden" id="cdResponsable" name="cdResponsable" value="<% =cdResponsable %>"/>
					</td>					
				</tr>
				<tr>				
					<td class="reg_header"><% =GF_TRADUCIR("Fecha Inicio") %></td>
					<td width="32px">
						<%	if (esModificable) then %>
						<a href="javascript:MostrarCalendario('imgInicio', SeleccionarCalInicio)"><img id="imgInicio" src="images/compras/calendar-16x16.png"></a>										
						<% end if %>
					</td>
					<td>								
						<%	if (esModificable) then %>												
						<div id="idateDiv" class="labelStyle"><% =fechaInicio %></div>															
						<%  else 
							 response.write fechaInicio 
						end if 	%>
						<input type="hidden" id="idate" name="idate" value="<% =fechaInicio %>"/>										
					</td>					
				<tr>
				</tr>
					<td class="reg_header"><% =GF_TRADUCIR("Fecha Fin") %></td>
					<td>
						<%	if (esModificable) then %>
						<a href="javascript:MostrarCalendario('imgFin', SeleccionarCalFin)"><img id="imgFin" src="images/compras/calendar-16x16.png"></a>					
						<%	end if 	%>
					</td>
					<td>
						<%	if (esModificable) then %>												
						<div id="fdateDiv" class="labelStyle"><% =fechaFin %></div>	
						<%  else 
							 response.write fechaFin
							end if 	%>
						<input type="hidden" id="fdate" name="fdate" value="<% =fechaFin %>" />					
					</td>					
				</tr>	
				</tr>
					<td class="reg_header"><% =GF_TRADUCIR("Fecha Ajustada") %></td>
					<td>						
						<a href="javascript:MostrarCalendario('imgAjustada', SeleccionarCalAjustada)"><img id="imgAjustada" src="images/compras/calendar-16x16.png"></a>											
					</td>
					<td>											
						<input type="hidden" id="adate" name="adate" value="<% =fechaAjustada %>" />					
						<div id="adateDiv" class="labelStyle"><% =fechaAjustada %></div>							
					</td>					
				</tr>	
					<tr>
					<% if (idObra <> 0) then %>					
							<td class="reg_header">
								<% =GF_TRADUCIR("Presupuesto") %>
							</td>
							<td colspan="2">
								
								<%  presupuestoEstimado = calcularCostoEstimadoObra(MONEDA_DOLAR, idObra,0,0)
									response.write getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(presupuestoEstimado, 2) 
									if (puedeModificarBudget(CdResponsable, fechaBudget, idDivision)) then	%>
										<a onclick="javascript:createBudget()"><img style="cursor:pointer" id="imgBudget" src="images/edit-16x16.png"></a>
								<%	else	%>
										
										<%if (puedeReasignarBudget(CdResponsable, idDivision)) then%>
											<a onClick="javascript:window.open('comprasBudgetReasignaciones.asp?idObra=<% =idObra %>') "><img style="cursor:pointer" id="imgBudget<% =idObra %>" src="images/compras/budget_view-16x16.png" title="<% =GF_TRADUCIR("Reasignar Presupuesto") %>"></a>
										<%else%>
											<a onClick="javascript:window.open('comprasBudgetObraPrint.asp?idObra=<% =idObra %>')"><img style="cursor:pointer" id="imgBudget<% =idObra %>" src="images/compras/budget_view-16x16.png" title="<% =GF_TRADUCIR("Ver Detalle Presupuesto") %>"></a>
										<%
										end if
								end if	%>
							</td>	
					<% end if %>				
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("División") %></td>
					<td colspan="2">
						<%	if (esModificable) then 
								strSQL="Select * from TBLDIVISIONES"
								Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
						%>
							<select id="idDivision" name="idDivision">
								<option value="SIN_DIVISION" selected="true">- <% =GF_TRADUCIR("Seleccione") %> -
						
								<%	
								while (not rsDivision.eof) 	
										if (checkPointAcceso(rsDivision("IDDIVISION"))) then 
											if not isAuditor(rsDivision("IDDIVISION")) then %>
												<option value="<% =rsDivision("IDDIVISION") %>" <% if (idDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
								<%			end if
										end if
										rsDivision.MoveNext()
								wend	
								%>								
							</select>							
						<%  else %>
							 <% =getDescripcionDivision(idDivision) %>
							 <input type="hidden" id="idDivision" name="idDivision" value="<% =idDivision %>" />					
						<%	end if 	%>
						
					</td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Tipo Partida") %></td>
					<td>						
						<select id="tipogasto" name="tipogasto">
							<option value="" >- <% =GF_TRADUCIR("Seleccione") %> -
							<%	strSQL = "select id AS codigo, descripcion from tbltipogastos"
								Call executeQueryDB(DBSITE_SQL_INTRA, rsTipoGastos, "OPEN", strSQL)
								while (not rsTipoGastos.eof)										
									%>
									<option value="<% =rsTipoGastos("codigo") %>"<% if (tipogasto = rsTipoGastos("codigo")) then response.write "selected='true'" %>><% =UCASE(rsTipoGastos("descripcion")) %>
									<%
									rsTipoGastos.MoveNext()
								wend
								'response.write GF_OPTIONS(rsTipoGastos,tipoGasto)
								%>
						</select>
					</td>					
				</tr>
				<%	if (esModificable) then %>												
					<tr>
						<td class="reg_header"><% =GF_TRADUCIR("Partida base") %></td>
						<td>
							<input type="hidden" id="CdObraBase" name="CdObraBase" value="<%=CdObraBase%>">
							<input id="DsObraBase" name="DsObraBase"  style="width:185px" value="<%=DsObraBase%>">
						</td>
					</tr>			
					<tr>
						<td class="reg_header"><% =GF_TRADUCIR("Items de Partida base a Copiar") %></td>
						<td>
							<input type="radio" id="soloRepetitivos0" name="soloRepetitivos" value="0" <% if (pSoloRepetitivos) then response.write "checked" %>/>Solo Repetitivos (<1000)
							<input type="radio" id="soloRepetitivos1" name="soloRepetitivos" value="1" <% if (not pSoloRepetitivos) then response.write "checked" %>/>Todos
						</td>
					</tr>			
				<%  end if 	%>
				
				<tr>
					<td align="right" colspan="3">
						<input type="button" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idObra = 0) then response.write "disabled=true" %> onClick="javascript:canSubmit()">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	
</table>

<input type="hidden" name="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" id="idObra" name="idObra" value="<% =idObra %>">
<input type="hidden" name="apFile" id="apFile" value="">
</form>		
</body>
</html>