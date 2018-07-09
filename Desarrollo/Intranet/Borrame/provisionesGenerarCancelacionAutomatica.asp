<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<%
Function controlarParametrosGeneracionCancelacionAutomatica(p_fechaDesdeDia,p_fechaHastaDia,p_fechaDesdeMes,p_fechaHastaMes,p_fechaDesdeAnio,p_fechaHastaAnio,p_fechaSaldoDia,p_fechaSaldoMes,p_fechaSaldoAnio,p_porcentaje)
    Dim ret,controlFecha
    controlFecha = GF_CONTROL_PERIODO(p_fechaDesdeDia,p_fechaHastaDia,p_fechaDesdeMes,p_fechaHastaMes,p_fechaDesdeAnio,p_fechaHastaAnio)
    Select case (controlFecha)
		case 0
			if (Cdbl(p_porcentaje) >= 0) then
                ret = true
            else    
                Call setError(VALOR_NO_VALIDO)
			end if
		case 1
			Call setError(FECHA_INICIO_INCORRECTA)
		case 2
			Call setError(FECHA_FIN_INCORRECTA)
		case 3
			Call setError(PERIODO_ERRONEO)
	end select    
    controlarParametrosGeneracionCancelacionAutomatica = ret
End Function
'----------------------------------------------------------------------------------------------------
Dim fechaDesde,fechaDesdeDia,fechaDesdeMes,fechaDesdeAnio,fechaHasta,fechaHastaDia,fechaHastaMes, fechaHastaAnio,porcentaje,accion,ret
Dim flagControl,fechaSaldoDia,fechaSaldoMes,fechaSaldoAnio,fechaSaldo

Call initTaskAccessInfo(TASK_EJE_PROVISIONS,"")
GP_CONFIGURARMOMENTOS

fechaDesdeDia = GF_PARAMETROS7("fechaDesdeDia","",6)
if (fechaDesdeDia = "") then fechaDesdeDia = GF_nDigits(Day(Now),2)
fechaDesdeMes = GF_PARAMETROS7("fechaDesdeMes","",6)
if (fechaDesdeMes = "") then fechaDesdeMes = GF_nDigits(Month(Now),2)
fechaDesdeAnio = GF_PARAMETROS7("fechaDesdeAnio","",6)
if (fechaDesdeAnio = "") then fechaDesdeAnio = Year(Now)
fechaDesde = fechaDesdeAnio & fechaDesdeMes & fechaDesdeDia

fechaHastaDia = GF_PARAMETROS7("fechaHastaDia","",6)
if (fechaHastaDia = "") then fechaHastaDia = GF_nDigits(Day(Now),2)
fechaHastaMes = GF_PARAMETROS7("fechaHastaMes","",6)
if (fechaHastaMes = "") then fechaHastaMes = GF_nDigits(Month(Now),2)
fechaHastaAnio = GF_PARAMETROS7("fechaHastaAnio","",6)
if (fechaHastaAnio = "") then fechaHastaAnio = Year(Now)
fechaHasta = fechaHastaAnio & fechaHastaMes & fechaHastaDia

fechaSaldoDia = GF_PARAMETROS7("fechaSaldoDia","",6)
if (fechaSaldoDia = "") then fechaSaldoDia = GF_nDigits(Day(Now),2)
fechaSaldoMes = GF_PARAMETROS7("fechaSaldoMes","",6)
if (fechaSaldoMes = "") then fechaSaldoMes = GF_nDigits(Month(Now),2)
fechaSaldoAnio = GF_PARAMETROS7("fechaSaldoAnio","",6)
if (fechaSaldoAnio = "") then fechaSaldoAnio = Year(Now)
fechaSaldo = fechaSaldoAnio & fechaSaldoMes & fechaSaldoDia

porcentaje = GF_PARAMETROS7("txtPorcentaje",0,6)
accion = GF_PARAMETROS7("accion","",6)

flagControl = false
if (accion = ACCION_CONTROLAR) then flagControl = controlarParametrosGeneracionCancelacionAutomatica(fechaDesdeDia,fechaHastaDia,fechaDesdeMes,fechaHastaMes,fechaDesdeAnio,fechaHastaAnio, fechaSaldoDia,fechaSaldoMes,fechaSaldoAnio, porcentaje)
    
if (accion = ACCION_PROCESAR) then 
    Call executeSP(rsFir, "STORED.PGM_TPE630", Cstr(fechaDesde) &"||"& Cstr(fechaHasta) &"||"& Cstr(fechaSaldo) &"||"& Cstr(porcentaje) &"||"& Cstr(Session("Usuario")))
    Response.End
end if
%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE PROVISIONES - Generar movimientos de cancelacion automatica</title>
	<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="css/main.css" />	
    <link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
    <link rel="stylesheet" href="css/Toolbar.css" type="text/css">
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
    <script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
    <script type="text/javascript" src="scripts/controles.js"></script>
    <script type="text/javascript" src="scripts/channel.js"></script>
    <script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript">
	    
	    var ch= new channel();
	    function bodyOnload() {
	        var	tb = new Toolbar('toolbar', 6, "images/");
	        tb.addButton("back-16.png", "Volver", "volver()");
	        tb.draw();
	        <% if (flagControl) then %>
                generarCancelacionAutomatica();
            <% end if %>
        }
	    function volver(){
	        document.location.href = "provisionesIndex.asp"
	    }
	    function generarCancelacionAutomatica(){
	        document.getElementById("actionLabel").style.visibility = 'visible';
	        document.getElementById("actionLabel").style.textAlign = 'center';
	        document.getElementById("actionLabel").style.fontSize = "16";
	        document.getElementById("actionLabel").innerHTML = '<img src="images/loading_small_black.gif"/>&nbsp;&nbsp;Generando cancelacion automatica...';
	        ch.bind("provisionesGenerarCancelacionAutomatica.asp?fechaDesdeDia=<%=fechaDesdeDia%>&fechaDesdeMes=<%=fechaDesdeMes%>&fechaDesdeAnio=<%=fechaDesdeAnio%>&fechaHastaDia=<%=fechaHastaDia%>&fechaHastaMes=<%=fechaHastaMes%>&fechaHastaAnio=<%=fechaHastaAnio%>&txtSaldo=<%=saldo%>&txtPorcentaje=<%=porcentaje%>&accion=<%=ACCION_PROCESAR%>","CallBack_generarCancelacionAutomatica()");
	        ch.send();
	    }
	    function CallBack_generarCancelacionAutomatica(){
	        var resp = ch.response();
	        document.getElementById("actionLabel").style.visibility = 'visible';
	        document.getElementById("actionLabel").style.textAlign = 'center';
	        document.getElementById("actionLabel").style.fontSize = "16";
	        document.getElementById("actionLabel").innerHTML = "Proceso finalizado";
	    }
	    function CerrarCal(cal) {
	        cal.hide();
	    }
	    function MostrarCalendario(p_objID, funcSel) {
	        var dte = new Date();
	        var elem = document.getElementById(p_objID);
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
	        var str = new String(date);
	        document.getElementById("divFechaHasta").innerHTML = str;
	        document.getElementById("fechaHastaDia").value = str.substr(0, 2);
	        document.getElementById("fechaHastaMes").value = str.substr(3, 2);
	        document.getElementById("fechaHastaAnio").value = str.substr(6, 4);
	        if (cal) cal.hide();
	    }
	    function SeleccionarCalDesde(cal, date) {
	        var str = new String(date);
	        document.getElementById("divFechaDesde").innerHTML = str;
	        document.getElementById("fechaDesdeDia").value = str.substr(0, 2);
	        document.getElementById("fechaDesdeMes").value = str.substr(3, 2);
	        document.getElementById("fechaDesdeAnio").value = str.substr(6, 4);
	        if (cal) cal.hide();
	    }
	    function SeleccionarCalSaldo(cal, date) {
	        var str = new String(date);
	        document.getElementById("divFechaSaldo").innerHTML = str;
	        document.getElementById("fechaSaldoDia").value = str.substr(0, 2);
	        document.getElementById("fechaSaldoMes").value = str.substr(3, 2);
	        document.getElementById("fechaSaldoAnio").value = str.substr(6, 4);
	        if (cal) cal.hide();
	    }
	</script>
</head>
<BODY onload="bodyOnload()">
    <div id="toolbar"></div><br>
    <form id="frmSel" name="frmSel" action="provisionesGenerarCancelacionAutomatica.asp" method="post">
	    <div class="tableasidecontent"><% call showErrors() %></div>
        <div class="col66"></div>
        <div class="tableasidecontent">
            <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha desde:")%></div>
            <div class="col26">
                <table>
				    <tr>
					    <td>
					        <a href="javascript:MostrarCalendario('img_fechaDesde', SeleccionarCalDesde)">
							    <img id="img_fechaDesde" src="images/calendar-16.png" title="Seleccionar fecha desde">
							</a>
						</td>	
						<td>
						    <div id="divFechaDesde">
							<%  Response.Write GF_FN2DTE(fechaDesde) %>
							</div>
                            <input type="hidden" id="fechaDesdeDia" name="fechaDesdeDia" value="<% =fechaDesdeDia %>" />
                            <input type="hidden" id="fechaDesdeMes" name="fechaDesdeMes" value="<% =fechaDesdeMes %>" />
                            <input type="hidden" id="fechaDesdeAnio" name="fechaDesdeAnio" value="<% =fechaDesdeAnio %>" />
						</td>
					</tr>
				</table>
            </div>
            <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%></div>
            <div class="col26"> 
                <table>
				    <tr>
					    <td>
					        <a href="javascript:MostrarCalendario('img_fechaHasta', SeleccionarCalHasta)">
							    <img id="img_fechaHasta" src="images/calendar-16.png" title="Seleccionar fecha Hasta">
							</a>
						</td>	
						<td>
						    <div id="divFechaHasta">
							<%  Response.Write GF_FN2DTE(fechaHasta) %>
							</div>
                            <input type="hidden" id="fechaHastaDia" name="fechaHastaDia" value="<% =fechaHastaDia %>" />
                            <input type="hidden" id="fechaHastaMes" name="fechaHastaMes" value="<% =fechaHastaMes %>" />
                            <input type="hidden" id="fechaHastaAnio" name="fechaHastaAnio" value="<% =fechaHastaAnio %>" />
						</td>
					</tr>
				</table>
            </div>
            <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha saldo:")%></div>
            <div class="col26"> 
                <table>
				    <tr>
					    <td>
					        <a href="javascript:MostrarCalendario('img_fechaSaldo', SeleccionarCalSaldo)">
							    <img id="img_fechaSaldo" src="images/calendar-16.png" title="Seleccionar fecha Saldo">
							</a>
						</td>	
						<td>
						    <div id="divFechaSaldo">
							<%  Response.Write GF_FN2DTE(fechaSaldo) %>
							</div>
                            <input type="hidden" id="fechaSaldoDia" name="fechaSaldoDia" value="<% =fechaSaldoDia %>" />
                            <input type="hidden" id="fechaSaldoMes" name="fechaSaldoMes" value="<% =fechaSaldoMes %>" />
                            <input type="hidden" id="fechaSaldoAnio" name="fechaSaldoAnio" value="<% =fechaSaldoAnio %>" />
						</td>
					</tr>
				</table>

            </div>
            <div class="col26 reg_header_navdos"> <%=GF_Traducir("Porcentaje indicativo para cancelar:")%></div>
            <div class="col26"> 
                <input type="text" id="txtPorcentaje" name="txtPorcentaje" value="<%=porcentaje %>" onkeypress="return controlIngreso(this,event,'N')" size="5" maxlength="4"/>
            </div>
        </div>
        <div class="col66"></div>
        <div class="col66"></div>
        <span class="btnaction">
            <input type="submit" value="<% =GF_TRADUCIR("Generar") %>" id=submitir name=submitir>
        </span>
        
        <div class="col66">&nbsp</div>
        <div id="actionLabel" class="confirmsj" style="width:100%;visibility:hidden;margin-top:10px;"></div>

	    <input type="hidden" id="accion" name="accion" value="<%=ACCION_CONTROLAR %>" />
    </form>

</BODY>
</html>