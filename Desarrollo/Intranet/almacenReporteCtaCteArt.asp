<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<%

dim RPT_Almacen, RPT_FechaDesde, RPT_FechaHasta, RPT_idArticulo, RPT_cdArticulo, RPT_dsArticulo, accion
dim rsAlmacenes, fechaInit, initDay, initMonth, initYear, flagCall

RPT_idArticulo = GF_Parametros7("idArticulo", 0, 6)
RPT_Almacen = GF_Parametros7("idAlmacen", "", 6)
initDay = day(date)
initMonth = month(date)
initYear = year(date)
call GF_STANDARIZAR_FECHA(initDay, initMonth, initYear)
fechaInit = initDay & "/" & initMonth & "/" & initYear
RPT_FechaDesde = GF_Parametros7("issuedate", "", 6)
if (RPT_FechaDesde = "") then RPT_FechaDesde = fechaInit
RPT_FechaHasta = GF_Parametros7("closingdate", "", 6)
if (RPT_FechaHasta = "") then RPT_FechaHasta = fechaInit
accion = GF_Parametros7("accion", "", 6)

Function datosEnviar()
	Dim zError
	zError = true
	if (RPT_idArticulo = 0) then
		setError(POCOS_ARTICULOS)
		zError = false
	end if
	if (RPT_Almacen = "") then
		setError(ALMACEN_NO_EXISTE)
		zError = false
	end if
	if (RPT_FechaDesde = "") or (RPT_FechaHasta = "")  then
		setError(PERIODO_ERRONEO)
		zError = false
	end if	
	datosEnviar = zError
end Function

flagCall=false
if (accion = ACCION_SUBMITIR) then 
	Call getArticuloFull(RPT_idArticulo,RPT_dsArticulo,RPT_cdArticulo)
	flagCall = datosEnviar()
end if

%>
<html>
<head>
<title><%=GF_TRADUCIR("Almacen - Reporte de Cuenta Corriente Articulos")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	
    var minDate = "01/01/2011";
    
	<% if (flagCall) then %>
		window.open("almacenReporteCtaCteArtPrint.asp?idArticulo=<% =RPT_idArticulo %>&idAlmacen=<% =RPT_Almacen %>&issuedate=<% =RPT_FechaDesde %>&closingdate=<% =RPT_FechaHasta %>");
	<% end if %>
	var ch = new channel();		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		tb.addButton("printer-16x16.png", "Generar", "submitInfo()");
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		tb.draw();

		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;
		msArticulo.setValue('<% =RPT_dsArticulo %>');		
		pngfix();
	}

	function submitInfo() {		
		document.getElementById("accion").value = '<%=ACCION_SUBMITIR%>';
		document.getElementById("frmSel").submit();
	}	
	
	function SeleccionarCalEmision(cal, date) {
		//Controlar que la fecha desde no sea mayor a la fecha hasta
		var str= new String(date);	
		//La fecha minima debe ser el inicio del stock cero!
		var rtrn = compareDates(minDate,"dd/MM/yyyy", str,"dd/MM/yyyy")
		if (rtrn == 1) {
		    str = minDate;
		    alert("La fecha desde no puede ser anterior a " + minDate);
		}
		//La fecha de emision no puede ser mayor que la de cierre.
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

	function seleccionarArticulo(ms) {
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("idArticulo").value = arr[0];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") document.getElementById("idArticulo").value = "";							
		}		
	}

	function volver() {	
		location.href = "almacenReportes.asp";
	}
</script>
</head>
<body onLoad="bodyOnLoad()">	
<div id="toolbar"></div>
<br>		

<form id="frmSel" name="frmSel" action="almacenReporteCtaCteArt.asp" method="POST">
<% Call showErrors() %>
<table class="reg_Header" id="TAB1" align="center" width="80%" border="0">				
	<tr>
		<td class="reg_Header_nav" align="left" colspan="6">
			<font class="big"><%=GF_Traducir("Reporte de Cuenta Corriente del Articulo")%></font>
		</td>
	</tr>	
	<tr>
		<!--Articulo-->	
		<td class="reg_Header_navdos"><% =GF_TRADUCIR("Articulo") %></td>
		<td colspan="2">
			<div id="articuloItem0"></div>																		
			<input type="hidden" id="idArticulo" name="idArticulo" value="<% =RPT_idArticulo%>">
		</td>
		<!--Almacen-->	
		<td class="reg_Header_navdos" align="left" width="15%">
			<%=GF_TRADUCIR("Almacen")%>
		</td>
		<td align="left" colspan="2">
			<% 	
			Set rsAlmacenes = obtenerListaAlmacenesSolicitud()
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
							
							<option <% if (rsAlmacenes("IDALMACEN") = cint(RPT_Almacen) ) then %> selected='selected' <% end if %> value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = RPT_idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
							<%	
							rsAlmacenes.MoveNext()
						wend 	
						%>
					</select>														
					<%		
				end if
			%>
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
	</tr>
	<tr><td></td></tr>
	<tr>
		<td align="left" colspan="6">
			<span style="color:red;"><% =GF_TRADUCIR("* Se incluye el ejercicio completo de la fecha desde y hasta") %></span>
		</td>
	</tr>
</table>
<input type="hidden" id="accion" name="accion" value="">
</form>
</body>
</html>