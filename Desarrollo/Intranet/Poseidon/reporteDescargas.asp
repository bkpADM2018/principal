<!--#include file="../Includes/procedimientosUnificador.asp"-->
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
Dim strSQL,rs,flagCall

Call GP_CONFIGURARMOMENTOS()
pto = GF_PARAMETROS7("pto", "", 6)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Puertos - Reporte de Descargas</title>

<link rel="stylesheet" type="text/css" href="../css/main.css"> 
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css" />
<link rel="stylesheet" type="text/css" href="../css/toolbar.css">
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">	

	function bodyOnLoad() {
		autoCompleteCliente();
		autoCompleteCorredor();	
		autoCompleteVendedor();
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
		document.getElementById("fd").value = str.substr(6,4) + '-' + str.substr(3,2) + '-' + str.substr(0,2)	    
		if (cal) cal.hide();
	}	
	function SeleccionarCalHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("dtFechaHasta").value = str;	    
		document.getElementById("fh").value = str.substr(6,4) + '-' + str.substr(3,2) + '-' + str.substr(0,2)	    
	    if (cal) cal.hide();	
	}		

	function autoCompleteVendedor(){
		$( "#dsVendedor" ).autocomplete({
			minLength: 2,
			source: "puertosStreamElementos.asp?tipo=JQVendedores&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsVendedor").val(ui.item.dsvendedor);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsVendedor"    ).val (ui.item.dsvendedor);
				$( "#ven"    ).val (ui.item.cdvendedor);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsVendedor").val ("");
					$( "#ven").val ("");
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
	
	
	function autoCompleteCliente(){
		$( "#dsCoordinado" ).autocomplete({
			minLength: 1,
			source: "puertosStreamElementos.asp?tipo=JQClientes&pto=<%=pto%>",
			focus: function( event, ui ) {
				$( "#dsCoordinado").val(ui.item.dscliente);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCoordinado"    ).val (ui.item.dscliente);
				$( "#cl"    ).val (ui.item.cdcliente);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsCoordinado").val ("");
					$( "#cl").val ("");
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
			source: "puertosStreamElementos.asp?tipo=JQCorredores&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsCorredor").val(ui.item.dscorredor);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCorredor"    ).val (ui.item.dscorredor);
				$( "#cor"    ).val (ui.item.cdcorredor);
				return false;
			},
			change: function( event, ui ) {				
				if (!ui.item) {					
					$( "#dscorredor").val ("");
					$( "#cor").val ("");
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
	
</script>
</head>

<body onLoad="bodyOnLoad()">

<form name="frmSel" id="frmSel" method="post" action="reporteDescargasPrint.asp" target="_blank">
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> Reporte de Descargas </h3>    
    <div id="searchfilter" class="tableasidecontent">        
		<div class="col66"></div>        
		<div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha desde:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaDesde" id="dtFechaDesde" readonly onclick="javascript:MostrarCalendario('dtFechaDesde', SeleccionarCalDesde)" value="<% =GF_FN2DTE(Left(session("MmtoDato"), 8)) %>">
					</td>
				</tr>
				<input type="hidden" id="fd" name="fd" value="<%=GF_FN2DTCONTABLE(Left(session("MmtoDato"), 8))%>">
			</table>
	    </div>
	    <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%> </div>
        <div class="col16">
   			<table>
				<tr>
					<td>
						<input type="text" name="dtFechaHasta" id="dtFechaHasta" readonly onclick="javascript:MostrarCalendario('dtFechaHasta', SeleccionarCalHasta)" value="<% =GF_FN2DTE(Left(session("MmtoDato"), 8)) %>">
					</td>
				</tr>
				<input type="hidden" id="fh" name="fh" value="<%=GF_FN2DTCONTABLE(Left(session("MmtoDato"), 8))%>">
			</table>
	    </div>
	    <div class="col16 reg_header_navdos"> Cliente </div>
        <div class="col16">
            <input type="text"   name="dsCoordinado" id="dsCoordinado" value="" style="width:150px">
			<input type="hidden" name="cl" id="cl" value="">
        </div>        
        <div class="col16 reg_header_navdos"> Corredor </div>        
        <div class="col16">
        	<input type="text"   name="dsCorredor" id="dsCorredor" value="" style="width:150px">
			<input type="hidden" name="cor" id="cor" value="">
        </div>        
        <div class="col16 reg_header_navdos"> Vendedor </div>        
        <div class="col16">
        	<input type="text"   name="dsVendedor" id="dsVendedor" value="<%=g_dsVendedor%>" style="width:150px">
			<input type="hidden" name="ven" id="ven" value="">
        </div>        
        <div class="col16 reg_header_navdos"> Transporte </div>
        <div class="col16">
			<select id="tt" name="tt">
				<option value="<%=TIPO_TRANSPORTE_CAMVAG %>"> Todos </option>
				<option value="<%=TIPO_TRANSPORTE_CAMION %>"> Camiones </option>
				<option value="<%=TIPO_TRANSPORTE_VAGON %>"> Vagones </option>
			</select>
        </div>        
        <div class="col16 reg_header_navdos"> Prooducto </div>
        <div class="col16">        
            <select id="prod" name="prod">
				<option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
				<%	strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM DBO.PRODUCTOS ORDER BY DSPRODUCTO"
					call GF_BD_Puertos (pto, rsProductos, "OPEN",strSQL)
					while not rsProductos.eof 
						if cint(g_cdProducto) = cint(rsProductos("CDPRODUCTO")) then
							mySelected = "SELECTED"
						else
							mySelected = ""
						end if	%>
						<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
					<%	rsProductos.movenext
					wend %>
			</select>
        </div>
        <span class="btnaction"><input type="submit" value="Buscar"></span>
    </div>
</div><!-- END BUSCAR -->
<input type="hidden" name="pto" id="pto" value="<% =pto %>">
</form>
</body>
</html>
