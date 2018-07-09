<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<%
Dim flagCall,accion,fechaInicio,fechaFin,pTipoReporte,pTipoReporte2

accion	= GF_PARAMETROS7("accion", "", 6)
fechaInicio	= GF_DTE2FN(GF_PARAMETROS7("fechaInicio", "", 6))
fechaFin	= GF_DTE2FN(GF_PARAMETROS7("fechaFin", "", 6))
pTipoReporte= GF_PARAMETROS7("tipoReporte","",6)
if (pTipoReporte = "") then pTipoReporte = REPORTE_CANTIDAD
pTipoReporte2= GF_PARAMETROS7("tipoReporte2","",6)
if (pTipoReporte2 = "") then pTipoReporte2 = REPORTE2_AMBOS


if (fechaFin<fechaInicio) then setError(PERIODO_ERRONEO)
if (GF_PARAMETROS7("fechaInicio", "", 6)="" and accion <> "") then setError(FECHA_INICIO_INCORRECTA)
if (GF_PARAMETROS7("fechaFin", "", 6)="" and accion <> "") then setError(FECHA_FIN_INCORRECTA)

flagCall = false
if (not hayError() and accion = ACCION_PROCESAR) then
	flagCall = true
end if
%>
<HTML>
		<HEAD>
			<LINK rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css"	 type="text/css">
			<LINK rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
			<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
		
			<SCRIPT type="text/javascript" src="scripts/channel.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/toolbar.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/jQueryAutocomplete.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/jQueryComboBox.js"></SCRIPT>
			<SCRIPT type="text/javascript" src="scripts/jQueryPopUp.js"></SCRIPT>
			<SCRIPT>
				ch = new channel();

				function bodyOnLoad()
				{
					$("#ver,#tipos" ).buttonset();

					$(function() {
						$( "#fechaInicio,#fechaFin" ).datepicker({
							showOn: "button",
							buttonImage: "images/almacenes/calendar-16x16.png",
							buttonImageOnly: true,
							dateFormat:"dd/mm/yy"
						});
					});

					var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
					tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
					tb.addButton("Previous-16x16.png", "Volver", "irA('almacenReportes.asp')");		
					tb.draw();	

					<% if (flagCall) then %>
						window.open('<%="almacenReporteArticulosNoConsumidosPrint.asp?fechaInicio="&fechaInicio&"&fechaFin="&fechaFin&"&tipoReporte="&pTipoReporte&"&tipoReporte2="&pTipoReporte2 %>');
					<% end if %>
				}

				function irA(pLink) {
					location.href = pLink;
				}
					
			</SCRIPT>
		</HEAD>
		<BODY onload="bodyOnLoad()">
			<% call GF_TITULO2("kogge64.gif","Articulos No Consumidos") %>
			<div id="toolbar"></div>
			<BR />
			<%=showErrors()%>
			<FORM method="get" action="almacenReporteArticulosNoConsumidos.asp">
				<TABLE align="center" class="reg_header" width="400px">
					<TR >
						<TD colspan="2" class="reg_header_nav ui-corner-top">&nbsp;</TD>
					</TR>
					<TR>
						<TD class="reg_header_navdos" align="right">Fecha inicio&nbsp;</TD>
						<TD><INPUT type="text"  readonly style="width:75px;background-color:rgb(255, 238, 205); border:0px" name="fechaInicio" id="fechaInicio" value="<%=GF_FN2DTE(fechaInicio)%>"></TD>
					</TR>
					<TR>
						<TD class="reg_header_navdos" align="right">Fecha fin&nbsp;</TD>
						<TD><INPUT type="text" readonly style="width:75px;background-color:rgb(255, 238, 205); border:0px" name="fechaFin" id="fechaFin" value="<%=GF_FN2DTE(fechaFin)%>"></TD>
					</TR>
					<TR>
						<TD class="reg_header_navdos" align="right">Listar&nbsp;</TD>
						<TD align="center">
							<DIV id="ver">
								<INPUT type="radio" id="radio1" name="tipoReporte" value="<%=REPORTE_CANTIDAD %>"	<% if (pTipoReporte = REPORTE_CANTIDAD) then response.write "checked='checked'" %>/><LABEL for="radio1">Cantidades</LABEL>
								<INPUT type="radio" id="radio2" name="tipoReporte" value="<%=REPORTE_PESO %>"		<% if (pTipoReporte = REPORTE_PESO) then response.write "checked='checked'" %>/><LABEL for="radio2">Pesos</LABEL>
								<INPUT type="radio" id="radio3" name="tipoReporte" value="<%=REPORTE_DOLAR %>"		<% if (pTipoReporte = REPORTE_DOLAR) then response.write "checked='checked'" %>/><LABEL for="radio3">Dolares</LABEL>
							</DIV>
						</TD>
					</TR>
					<TR>
						<TD class="reg_header_navdos" align="right">Reporte Articulos&nbsp;</TD>
						<TD align="center">
							<DIV id="tipos">
								<INPUT type="radio" id="radio01" name="tipoReporte2" value="<%=REPORTE2_SIN_STOCK %>"	<% if (pTipoReporte2 = REPORTE2_SIN_STOCK) then response.write "checked='checked'" %>/><LABEL for="radio01">Sin Stock</LABEL>
								<INPUT type="radio" id="radio02" name="tipoReporte2" value="<%=REPORTE2_CON_STOCK %>"		<% if (pTipoReporte2 = REPORTE2_CON_STOCK) then response.write "checked='checked'" %>/><LABEL for="radio02">Con Stock</LABEL>
								<INPUT type="radio" id="radio03" name="tipoReporte2" value="<%=REPORTE2_AMBOS %>"		<% if (pTipoReporte2 = REPORTE2_AMBOS) then response.write "checked='checked'" %>/><LABEL for="radio03">Ambos</LABEL>
							</DIV>
						</TD>
					</TR>
					<TR >
						<TD colspan="2" class="reg_header_nav ui-corner-bottom" align="center">
							<INPUT type="submit" value="Enviar Consulta">
						</TD>
					</TR>
				</TABLE>

				<INPUT type="hidden" name="accion" id="accion" value="<%=ACCION_PROCESAR%>">
			</FORM>
			<!--Se usan espacios en blanco para los calendarios entren dentro del body-->
			<!--Sino quedan mal ubicadas -->
			<BR /><BR /><BR /><BR /><BR /><BR /><BR />
		</BODY>
	</HTML>