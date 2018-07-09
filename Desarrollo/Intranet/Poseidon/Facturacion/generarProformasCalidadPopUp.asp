<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosFechas.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosTraducir.asp"-->

<%
'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'***************************************************
Dim myHoy, myHasta, tipoDescarga, onScreen, transporte, pto

Call GP_CONFIGURARMOMENTOS()

'						PARAMETROS OBLIGATORIOS
pto = GF_PARAMETROS7("pto", "", 6)
myHoy = Left(session("MmtoDato"), 8)
myHasta = myHoy
%>
<html>
	<head>
		<link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">
		<link rel="stylesheet" href="../../css/main.css" type="text/css"> 
		<script type="text/javascript" src="../../scripts/calendar.js"></script>
		<script type="text/javascript" src="../../scripts/calendar-1.js"></script>
		<script type="text/javascript" src="../../scripts/channel.js"></script>
		<script type ="text/javascript" >
		    var currSegment=0;	 
		    
			function generarMigracion() {				
				var ch = new channel();
				var pto = document.getElementById("pto").value;
				var td = document.getElementById("td").value;
				var fd = document.getElementById("fd").value;
				var fh = document.getElementById("fh").value;
				document.getElementById("msg").innerHTML = "";
				ch.bind("generarProformasCalidad.asp?pto=" + pto + "&fd=" + fd + "&fh=" + fh + "&td=" + td, "generarMigracion_cb()");				
				ch.send();				
			}
			
			function generarMigracion_cb() {
				document.getElementById("msg").innerHTML = "El proceso ha finalziado.";
			}
			
			function MostrarCalendarioDesde(p_objID, funcSel) {
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
				document.getElementById("generarFecha").value = "";
			}	

			function CerrarCal(cal) {
				cal.hide();
			}

			function SeleccionarCalDesde(cal, date) {
				var str= new String(date);		
				document.getElementById("dtFechaDesde").value = str;
				document.getElementById("fd").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			    if (cal) cal.hide();
			}	
			function QuitarFechaDesde(){
				document.getElementById("dtFechaDesde").value = "";			    
			}	
								
			function MostrarCalendarioHasta(p_objID, funcSel) {				
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

			function SeleccionarCalHasta(cal, date) {
				var str= new String(date);		
				document.getElementById("dtFechaHasta").value = str;
				document.getElementById("fh").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			    if (cal) cal.hide();
			}	
			function QuitarFechaHasta(){
				document.getElementById("dtFechaHasta").value = "";
			    
			}	
			function asignarValorComboBox(e,pName){
				document.getElementById(pName).value = e.value;
			}
			
		</script>
	</head>
	<body>
		<div class="tableaside size100"> <!-- BUSCAR -->
			<h3> Migracion de Datos de Acondicionamiento - Facturacion -<% =origen %></h3>
			<div id="msg"></div>
			<div id="searchfilter" class="tableasidecontent">
													   
				<div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Desde") %> </div>
				<div class="col26"> 
					<input type="text" name="dtFechaDesde" id="dtFechaDesde" onClick="javascript:MostrarCalendarioDesde('dtFechaDesde', SeleccionarCalDesde)" value="<% =GF_DateGet("D",myHoy) &"/"& GF_DateGet("M",myHoy) &"/"& GF_DateGet("A",myHoy)%>">
				</div>
				<div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Hasta") %> </div>        
				<div class="col26">
					<input type="text" name="dtFechaHasta" id="dtFechaHasta" onClick="javascript:MostrarCalendarioHasta('dtFechaHasta', SeleccionarCalHasta)" value="<% =GF_DateGet("D",myHasta) &"/"& GF_DateGet("M",myHasta) &"/"& GF_DateGet("A",myHasta)%>">
				</div>                                                                        
				<div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Tipo Descarga") %> </div>
				<div class="col26">            
					<select onchange="asignarValorComboBox(this,'td')" value="<%=cliente%>" >
						<option value="T"> <%=GF_Traducir("TODOS")%></option>
						<option value="<% =FACT_ACOND_DESCARGA_3ROS %>"> <%=GF_Traducir("De Terceros")%></option>
						<option value="<% =FACT_ACOND_DESCARGA_PROPIAS %>"> <%=GF_Traducir("Propias")%></option>				        
					</select>
				</div>
				<span class="btnaction"><input type="button" value="Migrar" onclick="generarMigracion();"></span>
			</div>
		</div><!-- END BUSCAR -->      
	<input type="hidden" name="fd" id="fd" value="<% =myHoy %>" />        
	<input type="hidden" name="fh" id="fh" value="<% =myHasta %>" />        	    
	<input type="hidden" name="pto" id="pto" value="<% =pto %>" />
	<input type="hidden" name="td" id="td" value="T" />	
</body>
</html>