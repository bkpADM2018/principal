<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
'Call controlAccesoAL("")

dim RPT_FechaDesde, RPT_FechaHasta, verDetalle
dim rsAlmacenes, rsSector

RPT_FechaDesde = GF_FN2DTE(Left(session("MmtoDato"),8))
RPT_FechaHasta = GF_FN2DTE(Left(session("MmtoDato"),8))
%>
<html>
<head>
<title><%=GF_TRADUCIR("Almacen - Vales Emitidos")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
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
</style>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	
	var ch = new channel();		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("../DocumentoTexto-16x16.png", "Imprimir PDF", "GenerarPDF()");
		tb.addButton("../excel3.gif", "Imprimir XLS", "GenerarXLS()");
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		tb.draw();			
		
		var msSolicitante = new MagicSearch("", "divSolicitante", 30, 4, "comprasStreamElementos.asp?tipo=personas");
		msSolicitante.setToken(";");
		msSolicitante.onBlur = seleccionarSolicitante;		
		
		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;		
		
		var msCategoria = new MagicSearch("", "divCategoria", 30, 4, "comprasStreamElementos.asp?tipo=categorias");
		msCategoria.setToken(";");
		msCategoria.onBlur = seleccionarCategoria;		
		
		var msCargo = new MagicSearch("", "divCargo", 30, 4, "comprasStreamElementos.asp?tipo=personas");
		msCargo.setToken(";");
		msCargo.onBlur = seleccionarCargo;		
		
		pngfix();
	}

	function hideLabel() {
		document.getElementById("actionLabel").style.visibility='hidden';
	}
	
	function submitInfo() {
		document.getElementById("actionLabel").style.visibility='visible';
		document.getElementById("frmSel").submit();
		setTimeout('hideLabel()',5000);
	}	

	function GenerarPDF() {
		document.getElementById("actionLabel").innerHTML='<% =GF_TRADUCIR("Generando Reporte en PDF") %>...';
		document.getElementById("frmSel").action='almacenReporteValesEmitidosPrint.asp';
		submitInfo();
	}

	function GenerarXLS() {
		document.getElementById("actionLabel").innerHTML='<% =GF_TRADUCIR("Generando Reporte en Excel") %>...';
		document.getElementById("frmSel").action='almacenReporteValesEmitidosPrintXLS.asp';
		submitInfo();
	}

	function SeleccionarCalEmision(cal, date) {
		//Controlar que la fecha desde no sea mayor a la fecha hasta
		var str= new String(date);		
		var auxDate = document.getElementById("closingdate").value;
		if (auxDate!=''){
			var rtrn = compareDates(str,"dd/MM/yyyy", auxDate,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha desde no puede ser mayor a la fecha hasta!");
				str = auxDate;
			}
		}
		document.getElementById("issuedateDiv").innerHTML = str;
		document.getElementById("issuedate").value = str;
		if (cal) cal.hide();	
	}
	function SeleccionarCalLimite(cal, date) {
		//Controlar que la fecha hasta no sea menor a la fecha desde
		var str= new String(date);	
		var auxDate = document.getElementById("issuedate").value;
		if (auxDate!=''){
			var rtrn = compareDates( auxDate,"dd/MM/yyyy", str,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha hasta no puede ser menor a la fecha desde!");
				str = auxDate;
			}
		}			
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

	function seleccionarCargo(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("cdCargo").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("cdCargo").value = "";							
		}		
	}	
	
	function seleccionarArticulo(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');			
			document.getElementById("idArticulo").value = arr[0];
			var arr2 = arr[1].split('[');
			ms.setValue(arr2[0]);			
		} else {
			if (desc == "") document.getElementById("idArticulo").value = "";							
		}		
	}	
	
	function seleccionarCategoria(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {			
			var arr = desc.split('|');			
			document.getElementById("idCategoria").value = arr[0];
			ms.setValue(arr[2]);
		} else {
			if (desc == "") document.getElementById("idCategoria").value = "";							
		}		
	}	
		
	function readBudgetArea() {
		document.getElementById('idBudgetArea').value=$("#idBudgetDetalle option:selected").attr("alt");
	}


	function actualizarBudgetsCallback(){
		document.getElementById("secBudgetDiv").innerHTML = ch.response(); 
	}

	function actualizarBudgets(idObra, idBudgetArea, idBudgetDetalle){
		var myReadOnly = 0;
		if (idObra==0){
			document.getElementById("secBudgetDiv").innerHTML = "";
			document.getElementById("idSector").value = 0;
			document.getElementById("idSector").disabled = false;
		}	
		else{
			ch.bind("almacenObtenerBudget.asp?idObra=" + idObra + "&idBudgetArea=" + idBudgetArea + "&idBudgetDetalle=" + idBudgetDetalle + "&readOnly=" + myReadOnly + "&accion=<%=ACCION_PROCESAR%>", "actualizarBudgetsCallback()");
			ch.send();
			document.getElementById("idSector").disabled = true;
		}	
	}
	function volver() {	
		location.href = "almacenReportes.asp";
	}
	function bloquearbudgets(idSector) {
		if (idSector==0){
			document.getElementById("idObra").value = 0;
			document.getElementById("idObra").disabled = false;
			document.getElementById("secBudgetDiv").innerHTML = "";
		} else {
			document.getElementById("idObra").disabled = true;
		}
	}
</script>
</head>
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Vales Emitidos") %>
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="almacenReporteValesEmitidosPrint.asp" target="_blank" method="POST">	
<table class="reg_Header" id="TAB1" align="center" width="80%" border="0">				
	<tr>
		<td class="reg_Header_nav" align="left" colspan="6">
			<font class="big"><%=GF_Traducir("Reporte de Vales Emitidos")%></big>
		</td>
	</tr>
	<tr>	
		<!--Almacen-->	
		<td class="reg_Header_navdos" align="left" width="15%">
			<%=GF_TRADUCIR("Almacen")%>
		</td>
		<td align="left" colspan="2">
			<% 	
			Set rsAlmacenes = obtenerListaAlmacenesUsuario()
				if rsAlmacenes.recordCount = 1 then
					response.write rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
					%>
					<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =rsAlmacenes("IDALMACEN") %>">
					<%
				else	
					%>
					<select id="idAlmacen" name="idAlmacen">							
						<%	
						while (not rsAlmacenes.eof)	
							%>
							<option value="<% =rsAlmacenes("IDALMACEN") %>"><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
							<%	
							rsAlmacenes.MoveNext()
							wend 	
						%>
					</select>														
					<%		
				end if
				%>
		</td>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Categoria") %></td>
		<td colspan="2">
			<div id="divCategoria"></div>																		
			<input type="hidden" id="idCategoria" name="idCategoria">
		</td>
	</tr>
	<tr>
		<!--Desde / Hasta-->
		<td class="reg_Header_navdos">
			<% =GF_TRADUCIR("Desde") %>
		</td>
		<td align="center" width="20%">
			<div id="issuedateDiv"><% =RPT_FechaDesde %></div>															
			<input type="hidden" id="issuedate" name="issuedate" value="<% =RPT_FechaDesde %>">
		</td>
		<td align="left" width="15%">
			<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCalEmision)"><img id="imgEmision" src="images/DATE.gif"></a>
		</td>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Articulo") %></td>
		<td colspan="2">
			<div id="articuloItem0"></div>																		
			<input type="hidden" id="idArticulo" name="idArticulo">
		</td>
	</tr>	
	<tr>
		<td class="reg_Header_navdos">
			<% =GF_TRADUCIR("Hasta") %>
		</td>
		<td align="center" width="22%">
			<div id="closingdateDiv"><% =RPT_FechaHasta %></div>	
			<input type="hidden" id="closingdate" name="closingdate" value="<% =RPT_FechaHasta %>">					
		</td>
		<td align="left" width="15%">
			<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/DATE.gif"></a>
		</td>
		<td class="reg_Header_navdos"><% =GF_TRADUCIR("Solicitante") %></td>
		<td colspan="2">
			<div id="divSolicitante"></div>																		
			<input type="hidden" id="cdSolicitante" name="cdSolicitante">
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos"><%= GF_TRADUCIR("Sector") %></td>
		<td colspan="2">
			<%
			Set rsSector = obtenerSectores("")
			%>
			<select id="idSector" name="idSector" onChange="bloquearbudgets(this.value)">
				<option value="0"><% =GF_TRADUCIR("Todos") %>
			<%
				while (not rsSector.eof)	%>
					<option value="<% =rsSector("IDSECTOR") %>"><% =GF_TRADUCIR(rsSector("DSSECTOR")) %></option>
			<%
					rsSector.MoveNext()
				wend
			%>
			</select>
		</td>
		<td class="reg_Header_navdos"><% =GF_TRADUCIR("Cargados por") %></td>
		<td colspan="2">
			<div id="divCargo"></div>																		
			<input type="hidden" id="cdCargo" name="cdCargo">
		</td>
	</tr>
	<tr>
		<td class="reg_Header_navdos"><%= GF_TRADUCIR("Part. Pres.") %></td>
		<td colspan="2">
			<%
			Set rsObras = obtenerListaObras("", "", "", "", "")
			%>
			<select id="idObra" name="idObra" onChange="actualizarBudgets(this.value,0,0)">
				<option value="0"><% =GF_TRADUCIR("Todas") %>
				<%
				while (not rsObras.eof)	%>
					<option value="<% =rsObras("IDOBRA") %>"><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %></option>
				<%
						rsObras.MoveNext()
				wend
				%>
				</select>
				<span id="secBudgetDiv"></span>
		</td>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Ver Detalle") %></td>
		<td colspan="2"><input type="checkbox" id="verDetalle" name="verDetalle" value="SI"></td>
	</tr>
</table>
<div align="center"><div id="actionLabel" class="round_border_bottom TDSUCCESS" style="width:80%;visibility:hidden;"><% =GF_TRADUCIR("Generando Reporte") %>...</div></div><br>
<input type="hidden" id="accion" name="accion" value="">
</form>
</body>
</html>
