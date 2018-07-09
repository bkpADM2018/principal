<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<%
REPORT_PDF = "PDF"
REPORT_XLS = "XLS"
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
Dim strSQL,rs,flagCall,g_cdProducto,pto

Call GP_CONFIGURARMOMENTOS()
pto = GF_PARAMETROS7("pto", "", 6)
Call addParam("pto", pto, params)
g_strPuerto = pto
g_accion = GF_PARAMETROS7("accion", "", 6)
g_fechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
if g_fechaDesdeD = "" then g_fechaDesdeD = GF_nDigits(Day(Now()),2)
g_fechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
if g_fechaDesdeM = "" then g_fechaDesdeM = GF_nDigits(Month(Now()),2)
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
g_cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)

flagCall=false
if (g_accion = ACCION_SUBMITIR) then
	ret = GF_CONTROL_PERIODO(g_fechaDesdeD, g_fechaHastaD, g_fechaDesdeM, g_fechaHastaM, g_fechaDesdeA, g_fechaHastaA)
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
<title>Poseidon - Reporte Merma Volatil</title>

<link rel="stylesheet" type="text/css" href="../css/main.css"> 
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css" />
<link rel="stylesheet" type="text/css" href="../css/toolbar.css">
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript">	

	

	function bodyOnLoad() {
	    <%	if (flagCall) then %>
               generateReporte();
		<%end if %>
		}
	function submitInfo(acc){
	    document.getElementById("accion").value = acc;
	    document.getElementById("frmSel").submit();
	}
	function generateReporte(){
	    window.open("reporteMermaVolatilPrintXLS.asp?pto=<%=g_strPuerto%>&fechaDesde=<%=g_fechaDesdeA&g_fechaDesdeM&g_fechaDesdeD %>&fechaHasta=<%=g_fechaHastaA&g_fechaHastaM&g_fechaHastaD %>&cdProducto=<%=g_cdProducto%>");
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
	function SeleccionarCalHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("dtFechaHasta").value = str;	    
	    document.getElementById("fechaHastaD").value = str.substr(0,2);
	    document.getElementById("fechaHastaM").value = str.substr(3,2);
	    document.getElementById("fechaHastaA").value = str.substr(6,4);
		if (cal) cal.hide();	
	}

</script>
</head>

<body onLoad="bodyOnLoad()">
<div id="toolbar"></div>
<form name="frmSel" id="frmSel" method="post" action="reporteMermaVolatilPrint.asp">
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> Reporte Merma Volatil </h3>
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
        <div class="col16 reg_header_navdos"> <%=GF_Traducir("Producto:")%> </div>
        <div class="col16">
         <% Call GF_BD_Puertos(pto, rsProd, "OPEN","SELECT CDPRODUCTO,DSPRODUCTO FROM PRODUCTOS" )  %>
   		    <select id="cdProducto" name="cdProducto">
                <option value="0"><%=GF_TRADUCIR("Seleccione") %></option>
                <% while (not rsProd.Eof) %>
                    <option value="<%=rsProd("CDPRODUCTO") %>" <% If(Cdbl(rsProd("CDPRODUCTO")) = Cdbl(g_cdProducto)) then %> selected <% end if %>><%=rsProd("DSPRODUCTO") %></option>
                <%    rsProd.MoveNext()
                   wend %>
            </select>
	    </div>
        <span class="btnaction"><input type="button" value="Generar XLS" id=cmdSearch name=cmdSearch onclick="submitInfo('<%=ACCION_SUBMITIR%>');"></span>	    
    </div>
</div><!-- END BUSCAR -->
<br>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="pto" name="pto" value="<% =pto %>">
</form>
</body>
</html>
