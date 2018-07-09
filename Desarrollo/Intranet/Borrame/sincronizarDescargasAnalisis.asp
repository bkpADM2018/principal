<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<%
'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'****************************************************
Dim myHoy,logMig, myHasta, onScreen, transporte

Call GP_CONFIGURARMOMENTOS()
if (session("Usuario") = "") then session("Usuario") = "SYNC"
'						PARAMETROS OBLIGATORIOS
origen = GF_PARAMETROS7("p", "", 6)
'						PARAMETROS OPCIONAL
myHoy = GF_PARAMETROS7("f", 0, 6)
if (myHoy = 0 ) then myHoy = GF_DTEADD(GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date)),-1,"D")

myHasta = GF_PARAMETROS7("ff", 0, 6)
if (myHasta = 0) then myHasta = myHoy

'Tomo el indicador de tipo de transporte a migrar.
transporte = TIPO_TRANSPORTE_CAMVAG
if (GF_PARAMETROS7("t", 0, 6) <> 0) then transporte = GF_PARAMETROS7("t", 0, 6)

onScreen = TIPO_NEGACION
if (GF_PARAMETROS7("v", "", 6) <> "") then onScreen = GF_PARAMETROS7("v", "", 6)

%>
<html>
	<head>
		<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
		<link rel="stylesheet" href="css/main.css" type="text/css"> 
		<script type="text/javascript" src="scripts/calendar.js"></script>
		<script type="text/javascript" src="scripts/calendar-1.js"></script>
		<script type ="text/javascript" >		    
		    
		    function bodyOnLoad() {			    
				<%	if (onScreen = TIPO_NEGACION) then %>				
				generarMigracion();
				<%	end if %>
			}
			
			function generarMigracion() {
				document.getElementById("frmSincro").submit();
			}
						
			
			function generateSegment_callback(pFechaHoy) {				
				if (pFechaHoy <= document.getElementById("ff").value) {
					document.getElementById("f").value = pFechaHoy;
					generarMigracion();			
				}
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
				
			}	

			function CerrarCal(cal) {
				cal.hide();
			}

			function SeleccionarCalDesde(cal, date) {
				var str= new String(date);		
				document.getElementById("dtFechaDesde").value = str;
				document.getElementById("f").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			    if (cal) cal.hide();
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
				document.getElementById("ff").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			    if (cal) cal.hide();
			}	
			function QuitarFechaHasta(){
				document.getElementById("dtFechaHasta").value = "";
			    
			}	
            function QuitarFechaDesde(){
				document.getElementById("dtFechaDesde").value = "";			    
			}		
			function asignarValorComboBox(e,pName){
				document.getElementById(pName).value = e.value;
			}
			
		</script>
	</head>
	<body onload="bodyOnLoad()">
<% if (onScreen <> TIPO_NEGACION) then  %>    
		<div class="tableaside size100"> <!-- BUSCAR -->
			<h3> Migracion de Datos de Acondicionamiento - Analisis -<% =origen %></h3>        
        <div id="searchfilter" class="tableasidecontent">
                                                   
            <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Desde") %> </div>
            <div class="col16"> 
				<input type="text" name="dtFechaDesde" id="dtFechaDesde" onClick="javascript:MostrarCalendarioDesde('dtFechaDesde', SeleccionarCalDesde)" value="<% =GF_DateGet("D",myHoy) &"/"& GF_DateGet("M",myHoy) &"/"& GF_DateGet("A",myHoy)%>">
			</div>
            <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Hasta") %> </div>        
            <div class="col16">
				<input type="text" name="dtFechaHasta" id="dtFechaHasta" onClick="javascript:MostrarCalendarioHasta('dtFechaHasta', SeleccionarCalHasta)" value="<% =GF_DateGet("D",myHasta) &"/"& GF_DateGet("M",myHasta) &"/"& GF_DateGet("A",myHasta)%>">
	        </div>                                                                                    
            <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Transporte") %> </div>
            <div class="col16"> 
	            <select  name="cdTransporte" id="cdTransporte" onchange="asignarValorComboBox(this,'t')" value="<%=transporte%>" >
	                <option value="<% =TIPO_TRANSPORTE_CAMVAG %>" <% if (transporte = TIPO_TRANSPORTE_CAMVAG) then response.write "selected"%>> <%=GF_Traducir("TODOS")%></option>
			        <option value="<% =TIPO_TRANSPORTE_CAMION %>" <% if (transporte = TIPO_TRANSPORTE_CAMION) then response.write "selected"%>> <%=GF_Traducir("CAMIONES")%></option>
			        <option value="<% =TIPO_TRANSPORTE_VAGON %>" <% if (transporte = TIPO_TRANSPORTE_VAGON) then response.write "selected"%>> <%=GF_Traducir("VAGONES")%></option>				        
	            </select>
            </div>            
            <span class="btnaction"><input type="button" value="Migrar" onclick="generarMigracion();"></span>
        </div>
    </div><!-- END BUSCAR -->  
    <iframe name="ifrmXLS" id="ifrmXLS" width="1px" height="1px" style="visibility:hidden"></iframe>
<%  else %>
    <iframe name="ifrmXLS" id="ifrmXLS" width="1024px" height="794px" ></iframe>
<% end if %>
	<form method="post" action="sincronizarDescargasAnalisisE1.asp" name="frmSincro" id="frmSincro" target="ifrmXLS">
	    <input type="hidden" name="f" id="f" value="<% =myHoy %>" />        
	    <input type="hidden" name="ff" id="ff" value="<% =myHasta %>" />        
	    <input type="hidden" name="t" id="t"  value="<% =transporte %>" />
	    <input type="hidden" name="p" id="p" value="<% =origen %>" />	    
	    <input type="hidden" name="v" id="v" value="<% =onScreen %>" />
	</form>
</body>
</html>