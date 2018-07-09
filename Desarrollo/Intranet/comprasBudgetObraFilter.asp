<!--#include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->

<%
dim pIdObra, pTipoCambio, pFecha,fechaFin
Dim browser,version,classNoHtml5, parcialChk,verFormularioAnual
classNoHtml5 = ""
pIdObra = GF_PARAMETROS7("idobra", 0, 6)
pFecha = GF_PARAMETROS7("fecha", 0, 6)
pTipoCambio = GF_PARAMETROS7("pTipoCambio", 0, 6)
'pTipoCambio = getTipoCambioBudget(pIdObra)
Set nav = Request.ServerVariables("HTTP_USER_AGENT")
fechaFin = GF_FN2DTE(left(Session("MmtoSistema"),8))
parcialChk = ""

verFormularioAnual = True
%>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/main.css" type="text/css">
	<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
	<style type="text/css">
		.frameprint{
			border: 1px solid black;
			width: 76%;
			height:99.9%;
			float: right;
		}
		#filtros{
			float: left;
			width: 23.5%;
		}
		#general{
			border: 1px solid black;
			margin: 0 auto;
			height:90%;
			width: 97%;
		}
		#titulo{
			display:block;
			border: 1px solid black;
			width: 97%;
			margin: 0 auto;
		}
		
	</style>
	<script defer type="text/javascript" src="scripts/pngfix.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
	<script type="text/javascript">
		var calendar;
		function verImpresion(){
			var fecha = document.getElementById("idate").value;
			var moneda = document.getElementById("comboMoneda").value;
			var parcial = "";
			var fac = "";
			var pic = "";
			var vale = "";
			var ctc = "";
						
			if (document.getElementById("bgtParcial").checked)		parcial = 1;
			if (document.getElementById("chkFacturacion").checked)	fac = 0;
			if (document.getElementById("chkPIC").checked)			pic = 1;
			if (document.getElementById("chkVales").checked)			vale = 1;
			
			<% if (verFormularioAnual) then%>
				document.getElementById("framePrint").src = 'comprasBudgetObraInvPrint.asp?idobra=<%=pIdObra%>&hasta='+fecha+'&moneda='+ moneda + '&bgtParcial=' + parcial + "&chkFacturacion=" + fac + "&chkPIC=" + pic + "&chkVales=" + vale;
			<% else %>
				var resumen    = document.getElementById("resumen").checked;
				var trimestre1 = document.getElementById("trimestre1").checked;
				var trimestre2 = document.getElementById("trimestre2").checked;
				var trimestre3 = document.getElementById("trimestre3").checked;
				var trimestre4 = document.getElementById("trimestre4").checked;
				var infoContable = document.getElementById("chkContable").checked;
				document.getElementById("framePrint").src = 'comprasBudgetObraPrint.asp?idobra=<%=pIdObra%>&hasta='+fecha+'&moneda='+moneda + '&bgtParcial=' + parcial + '&trimestre1=' + trimestre1+ '&trimestre2=' + trimestre2+ '&trimestre3=' + trimestre3+ '&trimestre4=' + trimestre4 + '&resumen=' + resumen + "&chkFacturacion=" + fac + "&chkPIC=" + pic + "&chkVales=" + vale + "&chkContable=" + infoContable;
			<% end if %>
			
		}
		
		function SeleccionarCal(cal, date) {
			var str= new String(date);		
			document.getElementById("idateDiv").innerHTML = str;
			document.getElementById("idate").value = str;
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

	</script>
</head>
<body onLoad="verImpresion()">
	<div id='titulo' class='titu_header round_border_top' align='center'><%=GF_TRADUCIR("Impresión de Budget")%></div>
	<div id='general' align='center'>
		<div id='filtros'>
			<p></p>
			<table width='210px'  align='center'>
				<thead>
					<tr width='100%'>						
						<th width="8px"></th>
						<th>Filtros</th>						
					</tr>
				</thead>
				<tbody>					
					<tr class="reg_header_nav">
						<td colspan="2"><%=GF_TRADUCIR("Consumos a la fecha")%>:</td>
					</tr>
					<tr>						
						<td colspan="2">
							<a href="javascript:MostrarCalendario('imgInicio', SeleccionarCal)">								
								<img id="imgInicio" src="images/compras/calendar-16x16.png">
								<span id="idateDiv"><% =fechaFin %></span><input type='hidden' id ='idate' value='<% =fechaFin %>'>								
							</a><br />							
							<input style="cursor:pointer;" type="checkbox" id="bgtParcial" name="bgtParcial">Estimar Budget a la Fecha
						</td>													
					</tr>
					<tr class="reg_header_nav">
						<td colspan="2"><%=GF_TRADUCIR("Incluir")%>:</td>
					</tr>
					<tr>					
						<td colspan="2">
							<input style="cursor:pointer;" type="checkbox" id="chkVales" name="chkVales" checked>Vales
							<input style="cursor:pointer;" type="checkbox" id="chkFacturacion" name="chkFacturacion" checked>Facturación
							<input style="cursor:pointer;" type="checkbox" id="chkPIC" name="chkPIC" checked>PICs
						</td>
					</tr>
					<% if (not verFormularioAnual) then%>
					<tr class="reg_header_nav">
						<td colspan='2'>Trimestres</td>
					</tr>
					<tr>
						<td colspan="2">
							<input type="checkbox" name="resumen"    id="resumen"    value="5" checked> Resumen General
							<input type="checkbox" name="trimestre1" id="trimestre1" value="1" title="Ene, Feb, Mar"> 1er Trimestre
							<input type="checkbox" name="trimestre2" id="trimestre2" value="2" title="Abr, May, Jun"> 2do Trimestre
							<input type="checkbox" name="trimestre3" id="trimestre3" value="3" title="Jul, Ago, Sep"> 3er Trimestre
							<input type="checkbox" name="trimestre4" id="trimestre4" value="4" title="Oct, Nov, Dic"> 4to Trimestre
						</td>
					</tr>
					<tr class="reg_header_nav">
						<td colspan="2"><%=GF_TRADUCIR("Informacion Contable")%></td>
					</tr>
					<tr>
						<td colspan="2">
							<input style="cursor:pointer;" type="checkbox" id="chkContable" name="chkContable">Incluir Cuenta y CC
						</td>
					</tr>											
					<%end if%>
					
					
					<tr class="reg_header_nav">
						<td colspan="2"><%=GF_TRADUCIR("Moneda")%></td>
					</tr>
					<tr>
						<td colspan="2">
							<%=getComboMonedas(MONEDA_DOLAR)%>
						</td>
					</tr>
				</tbody>
				<tfoot>
					<tr>
						<td width='80' colspan="2" align='center'><input class='round_border_bottom_right' type='button' value='Ver' onclick='verImpresion()' id='button'1 name='button'1></td>
					</tr>				
				</tfoot>
			</table>
		</div>
		<div class='frameprint'>			
	        <iframe id='framePrint' name='framePrint' width='100%' height='94%' frameborder=0></iframe>
		</div>
	</div>

</body>
</html>