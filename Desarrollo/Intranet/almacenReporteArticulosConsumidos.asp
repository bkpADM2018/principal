<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosParametros.asp"-->

<%
Const REPORT_PDF = "PDF"
Const REPORT_XLS = "XLS"
'--------------------------------------------------------------------------------------------------
'Funcion responsable por cargar un array con todos los períodos habilitados para pedir la info.
Function setArrayPeriodos()
			
	Dim ret(16)
	
	ret(0) = "Enero"
	ret(1) = "Febrero"
	ret(2) = "Marzo"
	ret(3) = "Abril"
	ret(4) = "Mayo"
	ret(5) = "Junio"
	ret(6) = "Julio"
	ret(7) = "Agosto"
	ret(8) = "Septiembre"
	ret(9) = "Octubre"
	ret(10) = "Noviembre"
	ret(11) = "Diciembre"
	ret(12) = "1er Trimestre"
	ret(13) = "2do Trimestre"
	ret(14) = "3er Trimestre"
	ret(15) = "4to Trimestre"
	ret(16) = "Todo el año"
	 
	setArrayPeriodos = ret
End function 
'--------------------------------------------------------------------------------------------------
'**************************************************************************************************
'*                                                                                                *
'*                                   INICIO DE PAGINA                                             *
'*                                                                                                *
'**************************************************************************************************
Dim pAlmacen,pPeriodo,pAnio,pSolicitante,pUnidad, pDivision, verDetalle
Dim arrPeriodos, flagCall, categoria,  conn, rs,pidArticulo, pVerDetalle, pCategoria, pGenerarReporte

pAlmacen	= GF_PARAMETROS7("almacen", 0, 6)
pDivision	= GF_PARAMETROS7("division", 0, 6)
pPeriodo	= GF_PARAMETROS7("periodo", 0, 6)
pCategoria	= GF_PARAMETROS7("categoria", 0, 6)
pAnio		= GF_PARAMETROS7("anio", 0,6)
pSolicitante= GF_PARAMETROS7("cdSolicitante","",6)
pidArticulo = GF_PARAMETROS7("idArticulo",0,6)
Call getArticuloFull(pidArticulo, pDsArticulo, "")
pVerDetalle = GF_PARAMETROS7("verDetalle","",6)
pCategoria = GF_PARAMETROS7("categoria","",6)
pGenerarReporte = GF_PARAMETROS7("GenerarReporte","",6)
pTipoReporte= GF_PARAMETROS7("tipoReporte","",6)
if (pTipoReporte = "") then pTipoReporte = REPORTE_CANTIDAD
pAccion		= GF_PARAMETROS7("accion","",6)

arrPeriodos = setArrayPeriodos()
flagCall=false
if (pAccion = ACCION_CONFIRMAR) then
	if ((pAlmacen < 0) and (pDivision=0)) then setError(ALMACEN_NO_EXISTE)
	if (pPeriodo < 0 or pAnio = 0 ) then	setError(FECHA_INICIO_INCORRECTA)
		
	if (not hayError()) then flagCall = true	
else
	if (pAccion = ACCION_PROCESAR) then
		strSQL="Select * from TBLALMACENES where IDDIVISION=" & pDivision
		Call executeQueryDB(DBSITE_SQL_INTRA, rsAlmacenes, "OPEN", strSQL)
		if (not rsAlmacenes.eof) then
			ret = "0|Todas;"
			while (not rsAlmacenes.eof)
				ret = ret & rsAlmacenes("IDALMACEN") & "|" & rsAlmacenes("DSALMACEN") & ";"
				rsAlmacenes.MoveNext()
			wend			
			ret = left(ret,Len(ret)-1)			
		else
			ret = "-1|No hay almacenes"
		end if
		Response.Write ret
		Response.end				
	end if
end if

%>


<html>
	<head>
	    <title>Reporte de Artículos Consumidos</title>
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
				<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
        <link rel="stylesheet" href="css/Toolbar.css" type="text/css">
		
		<script type="text/javascript" src="scripts/channel.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
		<script type="text/javascript" src="scripts/jQueryAutocomplete.js"></script>
		<script type="text/javascript" src="scripts/jQueryComboBox.js"></script>
		<script type="text/javascript" src="scripts/Toolbar.js"></script>
		<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
		<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
		<script type="text/javascript" src="scripts/controles.js"></script>
		<script type="text/javascript">	
		    
		    <% if (flagCall) then %>
                <% if (pGenerarReporte = REPORT_PDF) then %>
                     window.open("almacenReporteArticulosConsumidosPrint.asp?almacen=<% =pAlmacen%>&division=<% =pDivision%>&periodo=<% =pPeriodo%>&periodoDS=<% =arrPeriodos(pPeriodo)%>&anio=<% =pAnio%>&cdSolicitante=<% =pSolicitante%>&tipoReporte=<% =pTipoReporte %>&verDetalle=<%=pVerDetalle%>&idArticulo=<%=pidArticulo%>&categoria=<%=pCategoria%>");
                 <%else%>
                     window.open("almacenReporteArticulosConsumidosPrintXLS.asp?almacen=<% =pAlmacen%>&division=<% =pDivision%>&periodo=<% =pPeriodo%>&periodoDS=<% =arrPeriodos(pPeriodo)%>&anio=<% =pAnio%>&cdSolicitante=<% =pSolicitante%>&tipoReporte=<% =pTipoReporte %>&verDetalle=<%=pVerDetalle%>&idArticulo=<%=pidArticulo%>&categoria=<%=pCategoria%>");
			    <% end if %>
            <% end if %>
			var ch = new channel();

			
			$(function() {
					
				$( "#Solicitante" ).autocomplete({
				minLength: 2,
				source: "comprasStreamElementos.asp?tipo=JQPersonas",
				focus: function( event, ui ) {
					$( "#Solicitante" ).val(ui.item.nombre);
					return false;
				},
				select: function( event, ui ) {
					$( "#Solicitante"    ).val (ui.item.nombre);
					$( "#cdSolicitante"  ).val (ui.item.cdusuario );
								
					return false;
				},
				change: function( event, ui ) {
					if (!ui.item) {
						$( "#cdSolicitante"  ).val("");						
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
					.appendTo( ul );
			};
				//$("#division").combobox()
				//$("#almacen" ).combobox();
				//$("#periodo" ).combobox();
				$("#ver" ).buttonset();
			});
						
			function loadAlmacenes_callback() {				
				var rtrn = ch.response();				
				var arr = rtrn.split(";");
				var cmb = document.getElementById("almacen");
				var alm = <% =pAlmacen %>;
				cmb.options.length = 0;				
				for (i in arr) {				
					var vals = arr[i].split("|");
					var option=document.createElement("option");
					option.value=vals[0];
					option.text =vals[1];
					if (alm == vals[0]) option.selected = true;					
					cmb.add(option, null);					
				}							
			}
			
			function loadAlmacenes() {			
				var division = document.getElementById("division").value;
								
				ch.bind("almacenReporteArticulosConsumidos.asp?accion=<% =ACCION_PROCESAR %>&division=" + division, "loadAlmacenes_callback()");
				ch.send();
				
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

			function bodyOnLoad() {			
                tb = new Toolbar('toolbar', 6,'images/almacenes/');
			    tb.addButton("../pdf-16.png", "Imprimir PDF", "generarTipoReporte('<%=REPORT_PDF%>')");
			    tb.addButton("../excel-16.png", "Imprimir XLS", "generarTipoReporte('<%=REPORT_XLS%>')");
			    tb.addButton("Previous-16x16.png", "Volver", "volver()");
			    tb.draw();

			<%	if (pDivision <> 0) then %>
							loadAlmacenes();
				<%	end if %>
				var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
				msArticulo.setToken(";");	
				msArticulo.setValue('<% =pDsArticulo %>');	
				document.getElementById("idArticulo").value = '<% =pidArticulo %>';
				msArticulo.onBlur = seleccionarArticulo;	
			}	
			function volver() {	
			    location.href="almacenReportes.asp";
			}
			
			function generarTipoReporte(tipo){
			    document.getElementById("GenerarReporte").value = tipo;
			    document.getElementById("form1").submit();
			}
			
			
		</script>
		<style>
			.button.ui-button-icon-only
			{
				height:20px;
			}
		</style>
		
	</head>
    <div id="toolbar"></div>
	<body onLoad="bodyOnLoad()">		
		<form id="form1" name="form1">
			<table class="reg_Header " align="center" width="375px" cellspacing="2" cellpadding="2">
				<tr>
					<th colspan="2" class="ui-widget-header ui-corner-top">	&nbsp;</th>					
				</tr>
				<tr>
				    <td colspan="2"><%=showErrors()%></td>
				</tr>
				<tr>
					<td class="reg_Header_navdos">Almacen</td>
					<td>
						<select id="division" name="division" onChange="javascript:loadAlmacenes()">
							<option value="0">- Seleccione -</option>
							<% 	strSQL = "SELECT * FROM TBLDIVISIONES"
								Call executeQueryDB(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", strSQL)
								while (not rsDivisiones.EoF) 			%>
									<option value="<%=rsDivisiones("IDDIVISION")%>" <% if (pDivision=rsDivisiones("IDDIVISION")) then Response.Write "selected" %>><%=rsDivisiones("DSDIVISION")%></option>
							<%		rsDivisiones.MoveNext()
								wend %>
						</select><br><br>
						<select id="almacen" name="almacen">
							<option value="0">No hay almacenes</option>
						</select>
					</td>
				</tr>
							<tr>
					<td class="reg_Header_navdos">Período</td>
					<td>
						<select id="periodo" name="periodo">
						<option value="-1">- Seleccione -</option>
						<%	pos=0
							For each item in arrPeriodos	%>							
							<option value="<%=pos%>" <% if (pos = pPeriodo) then Response.Write "selected" %> ><% =item %></option>
						<%		pos=pos+1
							Next	%>							
						</select>
					</td>
				</tr>
				<tr>
					<td class="reg_header_navdos">Año</td>
					<td>
						<input type="text" id="anio" name="anio" size="4" value="<% if (pAnio <> 0) then Response.Write pAnio %>">
					</td>
				</tr>
				<tr>
					<td class="reg_header_navdos">Solicitante (Opcional)</td>
					<td>
						<span class="ui-widget">
							<input id="Solicitante" name="Solicitante"  style="width:185px" value="<%=GF_PARAMETROS7("Solicitante","",6)%>">
							<input type="hidden" name="cdSolicitante" id="cdSolicitante" value="<%=pSolicitante%>">
						</span>
					</td>
				</tr>
				<tr>
					<td class="reg_header_navdos">Categoria</td>
					<td>	
						<select name="categoria" id="categoria">
							<option value="-1">Todas</option>
								<%
								strSQL = "select idcategoria id,dscategoria ds from tblartcategorias where ESTADO = " & ESTADO_ACTIVO & " order by DSCATEGORIA"
								Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
								while not rs.eof
									%>
									<option value="<%=rs("id")%>"><%=rs("ds")%></option>
									<%
									rs.movenext
								wend%>
						</select>
					</td>
				</tr>
				<tr>
					<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Articulo") %></td>
					<td colspan="2">
						<div id="articuloItem0"></div>																		
						<input type="hidden" id="idArticulo" name="idArticulo">
					</td>						
				</tr>				
				<tr>
					<td class="reg_header_navdos">Listar</td>
					<td>
						<div id="ver">
						<%
							
						%>						
							<input type="radio" id="radio1" name="tipoReporte" value="<%=REPORTE_CANTIDAD %>"	<% if (pTipoReporte = REPORTE_CANTIDAD) then response.write "checked='checked'" %>/><label for="radio1">Cantidades</label>
							<input type="radio" id="radio2" name="tipoReporte" value="<%=REPORTE_PESO %>"		<% if (pTipoReporte = REPORTE_PESO) then response.write "checked='checked'" %>/><label for="radio2">Pesos</label>
							<input type="radio" id="radio3" name="tipoReporte" value="<%=REPORTE_DOLAR %>"		<% if (pTipoReporte = REPORTE_DOLAR) then response.write "checked='checked'" %>/><label for="radio3">Dolares</label>
						</div>
					</td>
				</tr>
				<tr>
					<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Ver") %></td>
					<td colspan="2">					
						<input type="checkbox" id="verDetalle" name="verDetalle" value="SI" <% if (pVerDetalle <> "") then  response.write "checked" %>>
						<%=GF_TRADUCIR("Detalle consumo articulo")%>
					</td>						
				</tr>
				
			</table>
			<input type="hidden" id="accion" name="accion" value="<%=ACCION_CONFIRMAR %>">
            <input type="hidden" id="GenerarReporte" name="GenerarReporte" value="">
		</form>
		</td></tr>
		</table>
	</body>
</html>