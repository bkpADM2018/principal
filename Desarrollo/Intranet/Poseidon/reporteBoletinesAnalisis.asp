﻿<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<%
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
dim  division,verPagosEfectuados,pto,idcamion,search_radio,params,fecContable
dim accion,nuCartaPorte1,nuCartaPorte2,nuCartaPorte3,fecContableD,fecContableM,fecContableA
dim flagCall,cdProducto,cdVendedor,dsVendedor,cdDestinatario,dsDestinatario,cdCoordinado, dsCoordinado
dim strSQLPro,rsProductos,cdEntregador,dsEntregador, fileCode,dsCalador,cdCalador,sector,fechaHasta,fechaDesde

Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
Call addParam("pto", pto, params)
sector = GF_PARAMETROS7("sector", 0, 6)
accion = GF_PARAMETROS7("accion", "", 6)
cdCoordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
dsCoordinador = GF_PARAMETROS7("dsCoordinador", "", 6)
cdCoordinado = GF_PARAMETROS7("cdCoordinado", "", 6)
dsCoordinado = GF_PARAMETROS7("dsCoordinado", "", 6)
producto = GF_PARAMETROS7("producto", "", 6)
fechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
fechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
fechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
fechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
fechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
fechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
sticker = GF_PARAMETROS7("sticker", "", 6)
cdCalador = GF_PARAMETROS7("cdCalador", "", 6)
dsCalador = GF_PARAMETROS7("dsCalador", "", 6)
certificado = GF_PARAMETROS7("certificado", "", 6)
grado = GF_PARAMETROS7("grado", 0, 6)
fileCode = GF_PARAMETROS7("fileCode", "", 6)
flagCall=false
if (accion = ACCION_SUBMITIR) then
	if ((fechaDesdeM <> "") or (fechaDesdeD <> "")) then fechaDesde = fechaDesdeD &"/"& fechaDesdeM &"/"& fechaDesdeA
	if ((fechaHastaM <> "") or (fechaHastaD <> "")) then fechaHasta = fechaHastaD &"/"& fechaHastaM &"/"& fechaHastaA
	ret = GF_CONTROL_PERIODO(fechaDesdeD, fechaHastaD, fechaDesdeM, fechaHastaM, fechaDesdeA, fechaHastaA)
	Select case (ret)
		case 0			
			flagCall=true
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
<title>Puertos - Reporte Boletines de Analisis</title>

<link rel="stylesheet" type="text/css" href="../css/main.css"> 

<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>


<script type="text/javascript">	

	var maxSegments;
	var currSegment=0;
	var MS_X_DAY = 86400000 //Milisegundos por día.	
	var d = new Date();
	function generateXLS() {
		document.getElementById("fileCode").value = "_" + document.getElementById("usr").value + "_" + currSegment;		
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel").innerHTML = "Inicializando... ";
		calculateSegments();
		document.getElementById("frmSel").action="reporteBoletinesAnalisisPrintE1.asp";
		document.getElementById("frmSel").target="ifrmXLS";
		generateSegment(currSegment)
	}
	
	function restartAttForm(){
		document.getElementById("accion").value = '<%=ACCION_SUBMITIR%>';
		var obj = document.getElementById("frmSel");
		obj.removeAttribute('target');
		obj.removeAttribute('action');
	}	
	
	function bodyOnLoad() {
		autoCompleteCoordinado();
        autoCompleteCalador();
        autoCompleteCoordinador();
		<%	if (flagCall) then %>
				generateXLS();
		<%	end if %>
	}
	function generateSegment(currSegment) {
		document.getElementById("actionLabel").innerHTML = "Recopilando datos...  ( " + (currSegment+1) + " / " + (maxSegments+1) + " )";
		var d = document.getElementById("fechaDesdeD").value;
		var m = document.getElementById("fechaDesdeM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fechaDesdeA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		var d = new Date(fd.getTime() + (MS_X_DAY*currSegment));
		document.getElementById("fecContableDS").value = d.getDate();
		document.getElementById("fecContableMS").value = d.getMonth()+1;	//getMonth() entrega el nro de mes de 0 a 11.
		document.getElementById("fecContableAS").value = d.getFullYear();
		document.getElementById("frmSel").action="reporteBoletinesAnalisisPrintE1.asp";
        document.getElementById("frmSel").target="ifrmXLS";
        document.getElementById("frmSel").submit();
	}
	
	function generateSegment_callback() {				
		if (currSegment < maxSegments) {			
			currSegment += 1; 
			document.getElementById("fileCode").value = "_" + document.getElementById("usr").value + "_" + currSegment;
			generateSegment(currSegment);
		} else {
			document.getElementById("maxSegment").value = currSegment;			
			document.getElementById("accion").value = '<%=ACCION_PROCESAR%>';			
			generateExcel();
		}
	}
	
	function generateExcel() {		
		document.getElementById("actionLabel").innerHTML = "Generando Excel...";
		setTimeout("document.getElementById('actionLabel').style.visibility = 'hidden'", 3000);
		document.getElementById("frmSel").action="reporteBoletinesAnalisisPrintE2.asp";
		document.getElementById("frmSel").target="";
		document.getElementById("frmSel").submit();
		restartAttForm();
	}
		
	function calculateSegments() {
		var d = document.getElementById("fechaDesdeD").value;
		var m = document.getElementById("fechaDesdeM").value-1; //El Month de Date trabaja de 0 a 11
		var y = document.getElementById("fechaDesdeA").value;
		var fd = new Date(y, m, d, 0, 0, 0, 0);		
		d = document.getElementById("fechaHastaD").value;
		m = document.getElementById("fechaHastaM").value-1; //El Month de Date trabaja de 0 a 11
		y = document.getElementById("fechaHastaA").value;
		var fh = new Date(y, m, d, 0, 0, 0, 0);		
		maxSegments = Math.round((fh.getTime() - fd.getTime())/MS_X_DAY)		
	}	
	function GenerarXLS(){
		document.getElementById("frmSel").action="reporteBoletinesAnalisis.asp";
		document.getElementById("frmSel").target="";		
		document.getElementById("frmSel").submit();
	}
	
	function irA() {	
		location.href = "puertosReportes.asp?pto=<%=pto%>&sector=<%=sector%>";
	}	
	
	function asignarProducto(me){
		document.getElementById("producto").value = me.value;
	}
	function asignarGrado(me){
		document.getElementById("grado").value = me.value;
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
	
	function autoCompleteCalador(){
		$( "#dsCalador" ).autocomplete({
			minLength: 2,
			source: "../comprasStreamElementos.asp?tipo=JQPersonas",
			focus: function( event, ui ) {
				$( "#dsCalador").val(ui.item.nombre);
				return false;
			},
			select: function( event, ui ) {
				$( "#dsCalador"    ).val(ui.item.nombre);
				$( "#cdCalador"    ).val(ui.item.cdusuario );
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsCalador").val ("");
					$( "#cdCalador").val ("");
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
	
	function autoCompleteCoordinador(){
		$( "#dsCoordinador" ).autocomplete({
			minLength: 1,
			source: "puertosStreamElementos.asp?tipo=JQEmpresas&pto=<%=pto%>",
			focus: function( event, ui ) {
				$( "#dsCoordinador").val(ui.item.dsempresa);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCoordinador"    ).val (ui.item.dsempresa);
				$( "#cdCoordinador"    ).val (ui.item.cdempresa);
				return false;
			},
			change: function( event, ui ) {				
				if (!ui.item) {
					$( "#dsCoordinador").val("");
					$( "#cdCoordinador").val("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdempresa + " - <font style='font-size:10;'>" + item.dsempresa + "</font></a>" )
				.appendTo( ul );
		};
	}	
	function autoCompleteCoordinado(){
		$( "#dsCoordinado" ).autocomplete({
			minLength: 1,
			source: "puertosStreamElementos.asp?tipo=JQClientes&pto=<%=pto%>",
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
</script>
</head>

<body onLoad="bodyOnLoad()">



<form name="frmSel" id="frmSel">
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> Reporte Boletines de Análisis </h3>
    <div ><% Call showMessages() %></div>
    <div id="searchfilter" class="tableasidecontent">        
		<div class="col66"></div>
        <div class="col16 reg_header_navdos"> Coordinador </div>
        <div class="col16">
            <input type="text"   name="dsCoordinador" id="dsCoordinador" value="<%=dsCoordinador%>" style="width:150px">
			<input type="hidden" name="cdCoordinador" id="cdCoordinador" value="<%=cdCoordinador%>">            
        </div>
		<div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha desde:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaDesde" id="dtFechaDesde" readonly onclick="javascript:MostrarCalendario('dtFechaDesde', SeleccionarCalDesde)" value="<% =fechaDesde %>">
					</td>
				</tr>
				<input type="hidden" id="fechaDesdeD" name="fechaDesdeD" value="<%=fechaDesdeD%>">
				<input type="hidden" id="fechaDesdeM" name="fechaDesdeM" value="<%=fechaDesdeM%>">
				<input type="hidden" id="fechaDesdeA" name="fechaDesdeA" value="<%=fechaDesdeA%>">
			</table>
	    </div>
        <div class="col16 reg_header_navdos"> Prooducto </div>
        <div class="col16">        
            <select id="cmbCdProducto" name="cmbCdProducto" onchange="javascript:asignarProducto(this);">
				<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
				<%	strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
					call GF_BD_Puertos (pto, rsProductos, "OPEN",strSQL)
					while not rsProductos.eof 
						if cint(producto) = cint(rsProductos("CDPRODUCTO")) then
							mySelected = "SELECTED"
						else
							mySelected = ""
						end if	%>
						<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
					<%	rsProductos.movenext
					wend %>
			</select>
			<input type="hidden" id="producto" name="producto" value="<%=g_cdProducto%>">
        </div>
        <div class="col16 reg_header_navdos"> Coordinado </div>
        <div class="col16">
            <input type="text"   name="dsCoordinado" id="dsCoordinado" value="<%=dsCoordinado%>" style="width:150px">
			<input type="hidden" name="cdCoordinado" id="cdCoordinado" value="<%=cdCoordinado%>">
        </div>        
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaHasta" id="dtFechaHasta" readonly onclick="javascript:MostrarCalendario('dtFechaHasta', SeleccionarCalHasta)" value="<% =fechaHasta %>">
					</td>
				</tr>
				<input type="hidden" id="fechaHastaD" name="fechaHastaD" value="<%=fechaHastaD%>">
				<input type="hidden" id="fechaHastaM" name="fechaHastaM" value="<%=fechaHastaM%>">
				<input type="hidden" id="fechaHastaA" name="fechaHastaA" value="<%=fechaHastaA%>">
			</table>
	    </div>
        <div class="col16 reg_header_navdos"> Sticker </div>
        <div class="col16"> <INPUT type='text' name='sticker' id='sticker' VALUE="<%=sticker%>"> </div>        
        <div class="col66"></div>        
        <div class="col16 reg_header_navdos"> Calador </div>        
        <div class="col16">
        	<input type="text"   name="dsCalador" id="dsCalador" value="<%=dsCalador%>" style="width:150px">
			<input type="hidden" name="cdCalador" id="cdCalador" value="<%=cdCalador%>">
        </div>        
        <div class="col16 reg_header_navdos"> Certificado </div>
        <div class="col16"> <INPUT type='text' name='certificado' id='certificado' VALUE="<%=certificado%>"> </div>        
        <div class="col16 reg_header_navdos"> Grado </div>
        <div class="col16">
            <select id="cmbGrado" name="cmbGrado" onchange="javascript:asignarGrado(this);">
				<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
				<option value="1" <%if(grado = 1)then Response.Write "SELECTED" %>><%= GF_TRADUCIR("Grado 1")%></option>				
				<option value="2" <%if(grado = 2)then Response.Write "SELECTED" %>><%= GF_TRADUCIR("Grado 2")%></option>				
				<option value="3" <%if(grado = 3)then Response.Write "SELECTED" %>><%= GF_TRADUCIR("Grado 3")%></option>
				<option value="4" <%if(grado = 4)then Response.Write "SELECTED" %>><%= GF_TRADUCIR("FE")%></option>
			</select>
			<input type="hidden" id="grado" name="grado" value="<%=grado%>">
        </div>
        <span style="text-align:center; clear:both; float:left; width:100%"><input type="submit" value="Exportar xls"></span>
    </div>
</div><!-- END BUSCAR -->
<br>
<div id="actionLabel" class="confirmsj" style="width:80%;visibility:hidden;"></div>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="pto" name="pto" value="<% =pto %>">
<input type="hidden" id="fileCode" name="fileCode" value="">
<input type="hidden" id="fecContableDS" name="fecContableDS">
<input type="hidden" id="fecContableMS" name="fecContableMS">
<input type="hidden" id="fecContableAS" name="fecContableAS">
<input type="hidden" id="maxSegment" name="maxSegment">
<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>">
<input type="hidden" id="sector" name="sector" value="<% =sector %>">
</form>
<iframe name="ifrmXLS" id="ifrmXLS" width="1px" height="1px" style="visibility:hidden"></iframe>
</body>
</html>
