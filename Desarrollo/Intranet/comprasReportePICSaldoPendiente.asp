<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
'***********************************************************************************************************
'											INICIO PAGINA
'***********************************************************************************************************
Dim g_FechaDesde,g_FechaHasta,g_idDivision,g_fechaDesdeD,g_fechaDesdeM,g_fechaDesdeA,g_fechaHastaD
Dim g_fechaHastaM,g_fechaHastaA,flagCall,g_accion

g_accion = GF_PARAMETROS7("accion", "", 6)
g_idDivision = GF_PARAMETROS7("idDivision", 0, 6)
g_fechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
if g_fechaDesdeD = "" then g_fechaDesdeD = "01"
g_fechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
if g_fechaDesdeM = "" then g_fechaDesdeM = "01"
g_fechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
if g_fechaDesdeA = "" then g_fechaDesdeA = GF_nDigits(Year(Now()),4)
g_fechaDesde = g_fechaDesdeD &"/"& g_fechaDesdeM &"/"& g_fechaDesdeA
g_fechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
if g_fechaHastaD = "" then g_fechaHastaD = GF_nDigits(Day(Now()),2)
g_fechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
if g_fechaHastaM = "" then g_fechaHastaM = GF_nDigits(Month(Now()),2)
g_fechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
if g_fechaHastaA = "" then g_fechaHastaA = GF_nDigits(Year(Now()),4)
g_fechaHasta = g_fechaHastaD &"/"& g_fechaHastaM &"/"& g_fechaHastaA

flagCall = false
if (g_accion = ACCION_SUBMITIR) then	 
	ret = GF_CONTROL_PERIODO(g_fechaDesdeD, g_fechaHastaD, g_fechaDesdeM, g_fechaHastaM, g_fechaDesdeA, g_fechaHastaA)
	Select case (ret)
		case 0
			if (g_idDivision <> 0) Then
				flagCall = true
			else
				Call setError(DIVISION_NO_EXISTE)
			end if
		case 1
			Call setError(FECHA_INICIO_INCORRECTA)
		case 2
			Call setError(FECHA_FIN_INCORRECTA)
		case 3
			Call setError(PERIODO_ERRONEO)
	end select
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Compras - Compras por Articulo</title>
<link rel="stylesheet" type="text/css" href="css/main.css"> 
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">	
	
	function bodyOnLoad() {			
		var tb = new Toolbar('toolbar', 5, "images/");	
		tb.addButtonRETURN("Volver", "irA()");
		tb.draw();
		<%	if (flagCall) then %>
				generateXLS();
		<%	end if %>	
	}
	
	function irA() {
		location.href = "comprasReportes.asp";
	}
	
	function generateXLS(){
		document.getElementById("actionLabel").innerHTML='<% =GF_TRADUCIR("Generando Reporte en Excel") %>...';
		document.getElementById("actionLabel").style.visibility='visible';
		document.getElementById("frmSel").action='comprasReportePICSaldoPendientePrintXLS.asp';
		submitInfo();
	}
	function submitInfo() {
		document.getElementById("frmSel").submit();
		document.getElementById("actionLabel").style.visibility='hidden';
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
	    document.getElementById("fechaDesdeD").value = str.substr(0,2);
	    document.getElementById("fechaDesdeM").value = str.substr(3,2);
	    document.getElementById("fechaDesdeA").value = str.substr(6,4);
		if (cal) cal.hide();
	}	
	function QuitarFechaDesde(){
		document.getElementById("dtFechaDesde").value = "";
	    document.getElementById("fechaDesdeD").value = "";
	    document.getElementById("fechaDesdeM").value = "";
	    document.getElementById("fechaDesdeA").value = "";
	}	
	function SeleccionarCalHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("dtFechaHasta").value = str;	    
	    document.getElementById("fechaHastaD").value = str.substr(0,2);
	    document.getElementById("fechaHastaM").value = str.substr(3,2);
	    document.getElementById("fechaHastaA").value = str.substr(6,4);
		if (cal) cal.hide();	
	}	
	function QuitarFechaHasta(){
		document.getElementById("dtFechaHasta").value = "";
	    document.getElementById("fechaHastaD").value = "";
	    document.getElementById("fechaHastaM").value = "";
	    document.getElementById("fechaHastaA").value = "";	    
	}		
	function volver() {	
		location.href = "comprasReportes.asp";
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">

<div id="toolbar"></div>

<form name="frmSel" id="frmSel">
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> Reporte de Pic con saldo pendiente </h3>
    <div ><% Call showMessages() %></div>
    <div id="searchfilter" class="tableasidecontent">        
		<div class="col66"></div>        
		<div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha desde:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaDesde" id="dtFechaDesde" readonly onclick="javascript:MostrarCalendario('dtFechaDesde', SeleccionarCalDesde)" value="<% =g_fechaDesde %>">
					</td>
				</tr>
				<input type="hidden" id="fechaDesdeD" name="fechaDesdeD" value="<%=g_fechaDesdeD%>">
				<input type="hidden" id="fechaDesdeM" name="fechaDesdeM" value="<%=g_fechaDesdeM%>">
				<input type="hidden" id="fechaDesdeA" name="fechaDesdeA" value="<%=g_fechaDesdeA%>">
			</table>
	    </div>
	    <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaHasta" id="dtFechaHasta" readonly onclick="javascript:MostrarCalendario('dtFechaHasta', SeleccionarCalHasta)" value="<% =g_fechaHasta %>">
					</td>
				</tr>
				<input type="hidden" id="fechaHastaD" name="fechaHastaD" value="<%=g_fechaHastaD%>">
				<input type="hidden" id="fechaHastaM" name="fechaHastaM" value="<%=g_fechaHastaM%>">
				<input type="hidden" id="fechaHastaA" name="fechaHastaA" value="<%=g_fechaHastaA%>">
			</table>
	    </div>	    
        <div class="col16 reg_header_navdos"> Division </div>
        <div class="col16">
		<%	Call executeProcedureDb(DBSITE_SQL_INTRA, rsDiv, "TBLDIVISIONES_GET", "") %>
			<select style="z-index:-1;" name="idDivision" id="idDivision">
		        <option SELECTED value="<% =SIN_DIVISION %>">- <% =GF_TRADUCIR("Seleccione") %> -
				<%	while (not rsDiv.eof)		
						selected = ""
						if (CLng(rsDiv("IDDIVISION")) = CLng(g_idDivision)) then selected = "selected" %>
						<option value="<% =rsDiv("IDDIVISION") %>" <% =selected %>><% =rsDiv("DSDIVISION") %>
				<%		rsDiv.MoveNext()
					wend	%>
			</select>			
        </div>        
        <span style="text-align:center; clear:both; float:left; width:100%"><input type="button" value="Exportar xls" onclick="submitInfo()"></span>
    </div>
</div><!-- END BUSCAR -->
<br>
<div id="actionLabel" class="confirmsj" style="width:80%;visibility:hidden;"></div>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">
</form>
</body>
</html>
