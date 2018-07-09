<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<%
Call initTaskAccessInfo(TASK_EJE_PROVISIONS,"")
'******************************************************************************************************************************
'***************************************************  INICIO DE LA PAGINA  ****************************************************
'******************************************************************************************************************************
dim nroLote, fechaLote, estado, paginaActual, mostrar, rsPro, totalRegistros,flagFirmas

nroLote      = GF_PARAMETROS7("nroLote",0,6)
fechaLote    = GF_PARAMETROS7("fechaLote",0,6)
estado       = GF_PARAMETROS7("estado","",6)
fechaSaldo   = GF_PARAMETROS7("fechaSaldo",0,6)
porcentaje   = GF_PARAMETROS7("porcentaje",2,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
mostrar      = GF_PARAMETROS7("registrosPorPagina",0,6)
if (paginaActual = 0) then paginaActual=1
if (mostrar = 0) then mostrar = 10

GP_ConfigurarMomentos


Set sp_ret = executeSP(rsPro, "EJIFL.TBLPROVISIONESCANE_GET_TOTALES_BY_PARAMETERS", nroLote &"||"& fechaLote &"||"& fechaSaldo &"||"& porcentaje &"||"& estado &"||"& paginaActual &"||"& mostrar&"$$totalRegistros")
totalRegistros = sp_ret("totalRegistros")


%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE PROVISIONES - Autorizaci&oacuten de provisiones desde cancelaci&oacuten autom&aacutetica</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/paginar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
	<script type="text/javascript" src="scripts/paginar.js"></script>
	<script type="text/javascript" src="scripts/date.js"></script>
    <script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
    <script type="text/javascript" src="scripts/controles.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
   	<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="scripts/jqueryPopUp.js"></script>
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>	
	<script type="text/javascript">
	    
	    var ch= new channel();

	    function bodyOnload() {
	        var	tb = new Toolbar('toolbar', 6, "images/");
	        tb.addButton("back-16.png", "Volver", "volver()");
	        tb.addButton("toolbar-refresh","<%=GF_Traducir("Recargar")%>", "submitInfo()");
	        tb.draw();
	        var pgn = new Paginacion("paginacion");
	        pgn.paginar(<% =paginaActual %>, <% =totalRegistros %>, <% =mostrar %>, 50, "submitInfo()");        
	    }
	    function volver(){
	        document.location.href = "provisionesIndex.asp"
	    }
	    function submitInfo() {
	        document.getElementById("frmSel").submit();
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
	    function SeleccionarCalLote(cal, date) {
	        var str = new String(date);
	        document.getElementById("fechaLoteDiv").innerHTML = str + "<a href=javascript:resetFechasLote()><img src='images/button_cancel.png'></a>";
	        document.getElementById("fechaLote").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
	        if (cal) cal.hide();
	    }
	    function resetFechasLote() {
	        document.getElementById("fechaLoteDiv").innerHTML = "";
	        document.getElementById("fechaLote").value = "";
	    }
	    function SeleccionarCalSaldo(cal, date) {
	        var str = new String(date);
	        document.getElementById("fechaSaldoDiv").innerHTML = str + "<a href=javascript:resetFechasSaldo()><img src='images/button_cancel.png'></a>";
	        document.getElementById("fechaSaldo").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
	        if (cal) cal.hide();
	    }
	    function resetFechasSaldo() {
	        document.getElementById("fechaSaldoDiv").innerHTML = "";
	        document.getElementById("fechaSaldo").value = "";
	    }
	    function verFirmasRegistradas(p_NroLote, p_FechaLote){
	        var popUpFirmas = new winPopUp('popUpFirmas', "provisionesCancelacionAutomaticaFirmaPopUp.asp?nroLote="+p_NroLote+"&fechaLote="+p_FechaLote, '500', '250', 'Firmas registradas', "");
	    }
	    function verDetalleLote(p_NroLote, p_FechaLote){
	        var popUpFirmas = new winPopUp('popUpFirmas', "provisionesCancelacionAutomaticaPopUp.asp?nroLote="+p_NroLote+"&fechaLote="+p_FechaLote, '800', '500', 'Detalle del lote', "submitInfo()");
	    }
	    function borrarLote(p_NroLote, p_FechaLote){
	        if(confirm("Desea eliminar el lote?")){
	            ch.bind("provisionesCancelacionAutomaticaAjax.asp?nroLote=" + p_NroLote +"&fechaLote="+ p_FechaLote +"&accion=<%=ACCION_BORRAR%>","CallBack_borrarLote()");
	            ch.send();
	        }
	    }
	    function CallBack_borrarLote(){
	        submitInfo();
	    }
	    function verPDF(p_NroLote, p_FechaLote){
	        window.open("provisionesCancelacionAutomaticaPrint.asp?nroLote=" + p_NroLote+"&fechaLote="+p_FechaLote, "Impresion de Provision") 
	    }
	</script>
</head>
<body onload="bodyOnload()">
<div id="toolbar"></div><br>
<form id="frmSel" name="frmSel" action="provisionesAdministrarCancelacionAutomatica.asp" method="post">
    <div class="tableaside size100"> <!-- BUSCAR -->
	    <h3> Filtros </h3>
		<div id="searchfilter" class="tableasidecontent">
		    <div class="col16 reg_header_navdos"> <%=GF_Traducir("Nro.Lote:")%> </div>
		    <div class="col16"><input type="text" size="7" maxlength="9" id="nroLote" name="nroLote" value="<% if (CDbl(nroLote) <> 0) then response.write nroLote end if %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
		        
            <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Lote:")%> </div>
		    <div class="col16"> 
   			    <table>
				    <tr>
					    <td>
					        <a href="javascript:MostrarCalendario('img_fechaLote', SeleccionarCalLote)">
							    <img id="img_fechaLote" src="images/calendar-16.png" title="Seleccionar fecha">
							</a>
						</td>	
						<td>
						    <div id="fechaLoteDiv">
							<%  if CDbl(fechaLote) <> 0 then
								    Response.Write fechaLote
									Response.Write "<a href=javascript:resetFechasLote()><img src='images/button_cancel.png'></a>"
								end if %>	
							</div>
						</td>	
					</tr>	
					<input type="hidden" id="fechaLote" name="fechaLote" value="<% =fechaLote %>" />
				</table>
		    </div>
            <div class="col16 reg_header_navdos"> <%=GF_Traducir("Fecha Saldo:")%> </div>
		    <div class="col16"> 
   			    <table>
				    <tr>
					    <td>
					        <a href="javascript:MostrarCalendario('img_fechaSaldo', SeleccionarCalSaldo)">
							    <img id="img_fechaSaldo" src="images/calendar-16.png" title="Seleccionar fecha">
							</a>
						</td>	
						<td>
						    <div id="fechaSaldoDiv">
							<%  if CDbl(fechaSaldo) <> 0 then
								    Response.Write fechaSaldo
									Response.Write "<a href=javascript:resetFechasSaldo()><img src='images/button_cancel.png'></a>"
								end if %>	
							</div>
						</td>	
					</tr>	
					<input type="hidden" id="fechaSaldo" name="fechaSaldo" value="<% =fechaSaldo %>" />
				</table>
		    </div>
            <div class="col16 reg_header_navdos"> <%=GF_Traducir("Porcentaje:")%> </div>
		    <div class="col16"> 
                <input type="text" id="porcentaje" name="porcentaje" value="<% if (CDbl(porcentaje) <> 0) then response.write porcentaje end if %>" size="7" maxlength="8" onkeypress="return controlIngreso(this, event, 'I')"/>
		    </div>
            <div class="col16 reg_header_navdos"> <%=GF_Traducir("Estado:")%> </div>
            <div class="col16"> 
                <select id="estado" name="estado">
                    <option value=""> <%=GF_TRADUCIR("TODOS") %></option>
                    <option value="<%=PROVISCIONES_ESTADO_GENERADO %>" <% if (Cstr(estado) = PROVISCIONES_ESTADO_GENERADO) then %> selected <% end if %>><%=GF_TRADUCIR("GENERADO") %></option>
                    <option value="<%=PROVISCIONES_ESTADO_PENDIENTE %>" <% if (Cstr(estado) = PROVISCIONES_ESTADO_PENDIENTE) then %> selected <% end if %>><%=GF_TRADUCIR("PENDIENTE") %></option>
                    <option value="<%=PROVISCIONES_ESTADO_AUTORIZADO %>" <% if (Cstr(estado) = PROVISCIONES_ESTADO_AUTORIZADO) then %> selected <% end if %>><%=GF_TRADUCIR("AUTORIZADO") %></option>
                    <option value="<%=PROVISCIONES_ESTADO_APLICADO %>" <% if (Cstr(estado) = PROVISCIONES_ESTADO_APLICADO) then %> selected <% end if %>><%=GF_TRADUCIR("APLICADO") %></option>
                    <option value="<%=PROVISCIONES_ESTADO_ERROR %>" <% if (Cstr(estado) = PROVISCIONES_ESTADO_ERROR) then %> selected <% end if %>><%=GF_TRADUCIR("ERROR") %></option>
                </select>
		    </div>
            <BR>

		    <span class="btnaction"><input type="submit" value="<% =GF_TRADUCIR("Buscar") %>" id=submitir name=submitir></span>
        </div>
    </div>
	<br>
	<table class="datagrid" width="100%" align="center">
	    <thead>
	        <tr>
				<th class="thiconac" rowspan="2" align="center" style="width:6%;"><%= GF_TRADUCIR("Nro.Lote") %></th>
				<th class="thiconac" rowspan="2" align="center" style="width:8%;"><%= GF_TRADUCIR("Fecha Lote") %></th>
                <th class="thiconac" rowspan="2" align="center" style="width:8%;"><%= GF_TRADUCIR("Fecha Saldo") %></th>
                <th class="thiconac" colspan="2" align="center" style="width:24%;"><%= GF_TRADUCIR("Pesos") %></th>
                <th class="thiconac" colspan="2" align="center" style="width:24%;"><%= GF_TRADUCIR("Dolares") %></th>
                <th class="thiconac" rowspan="2" align="center" style="width:6%;"><%= GF_TRADUCIR("Porcentaje") %></th>
                <th class="thiconac" rowspan="2" align="center" style="width:12%;"><%= GF_TRADUCIR("Estado") %></th>
                <th class="thiconac" rowspan="2" align="center" style="width:3%;">.</th>
                <th class="thiconac" rowspan="2" align="center" style="width:3%;">.</th>
                <th class="thiconac" rowspan="2" align="center" style="width:3%;">.</th>
                <th class="thiconac" rowspan="2" align="center" style="width:3%;">.</th>
            </tr>
            <tr>
                <td class="thiconac" align="center" style="width:12%;"><%= GF_TRADUCIR("Gastos a cancelar") %></td>
                <td class="thiconac" align="center" style="width:12%;"><%= GF_TRADUCIR("Gastos aplicados") %></td>
                <td class="thiconac" align="center" style="width:12%;"><%= GF_TRADUCIR("Gastos a cancelar") %></td>
                <td class="thiconac" align="center" style="width:12%;"><%= GF_TRADUCIR("Gastos aplicados") %></td>
            </tr>
		</thead>
		<tbody>	
        <%  if (not rsPro.Eof) then %>
            <%  while (not rsPro.Eof) %>
			    <tr>
			        <td align="center"><%=rsPro("NROLOTE") %></td>
				    <td align="center"><%=GF_FN2DTE(rsPro("FECHALOTE")) %></td>
                    <td align="center"><%=GF_FN2DTE(rsPro("FECHASALDO")) %></td>
                    <td align="right"><%=TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(rsPro("TOTALCANCELACIONPESOS"))*100,2) %></td>
                    <td align="right"><%=TIPO_MONEDA_PESO &" "& GF_EDIT_DECIMALS(Cdbl(rsPro("TOTALAPLICADOCANCELACIONPESOS"))*100,2) %></td>
                    <td align="right"><%=TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(rsPro("TOTALCANCELACIONDOLARES"))*100,2) %></td>
                    <td align="right"><%=TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(rsPro("TOTALAPLICADOCANCELACIONDOLARES"))*100,2) %></td>
                    <td align="center"><%=rsPro("PORCENTAJE") & "%" %></td>
                    <td align="center"><%= getEstadoProvisionesCancelacion(rsPro("ESTADO")) %></td>
                    <td align="center">
                        <a href="javascript:verFirmasRegistradas(<%= rsPro("NROLOTE")%>,<%=rsPro("FECHALOTE") %>);">
                            <img title="Ver firmas regsitradas" style="cursor:pointer;" src="images/perfil-16.png">
                        </a>
                    </td>
				    <td align="center">
                        <a href="javascript:verPDF(<%= rsPro("NROLOTE")%>,<%=rsPro("FECHALOTE") %>);">
                            <img title="Ver PDF" style="cursor:pointer;" src="images/pdf-16.png">
                        </a>
                    </td>
                    <td align="center">
                        <a href="javascript:verDetalleLote(<%= rsPro("NROLOTE")%>,<%=rsPro("FECHALOTE") %>);">
                            <img title="Ver detalle" style="cursor:pointer;" src="images/search-16.png">
                        </a>
                    </td>
                    <td align="center">
                        <a href="javascript:borrarLote(<%= rsPro("NROLOTE")%>,<%=rsPro("FECHALOTE") %>);">
                            <img title="Eliminar" style="cursor:pointer;" src="images/trash-16.png">
                        </a>
                    </td>
                </tr>
            <%  rsPro.MoveNext() %>
        <%   wend %>
        <%  else %>
            <tr>
                <td colspan="13" align="center"><%=GF_TRADUCIR("No se encontraron resultados") %></td>
            </tr>
        <%  end if %>
        </tbody>
        <tfoot>
            <tr>
			    <td colspan="13"><div id="paginacion"></div></td>
			</tr>
		</tfoot>
	</table>
	
</form>
</body>
</html>