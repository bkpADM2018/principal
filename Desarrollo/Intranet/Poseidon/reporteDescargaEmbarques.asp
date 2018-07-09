<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosTitulos.asp"-->
<%
'Verifica si el array Productos seleccionados corresponde al producto pasado por parametro  
Function isSelectedProducto(pArrPro,pCdProd)
	isSelectedProducto = false
	For i = LBound(pArrPro) to UBound(pArrPro)
		if (Cdbl(pArrPro(i)) = Cdbl(pCdProd)) then
			isSelectedProducto = true
			exit for
		end if
	Next
End Function
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto, fechaDesdeD,fechaDesdeM,fechaDesdeA,fechaHastaD,fechaHastaM,fechaHastaA,accion,g_cdProducto,g_chkDetalle,arrProductos

g_strPuerto = GF_Parametros7("Pto","",6)
accion = GF_Parametros7("accion","",6)

fechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
if (fechaDesdeD = "") then fechaDesdeD= GF_nDigits(	Day(Now()),2)
fechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
if (fechaDesdeM = "") then fechaDesdeM=GF_nDigits(Month(Now()),2)
fechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
if (fechaDesdeA = "") then fechaDesdeA=GF_nDigits(Year(Now()),4)

fechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
if (fechaHastaD = "") then fechaHastaD=GF_nDigits(Day(Now()),2)
fechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
if (fechaHastaM = "") then fechaHastaM=GF_nDigits(Month(Now()),2)
fechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
if (fechaHastaA = "") then fechaHastaA=GF_nDigits(Year(Now()),4)

fechaDesde = fechaDesdeA & "-" & fechaDesdeM & "-" & fechaDesdeD
fechaHasta = fechaHastaA & "-" & fechaHastaM & "-" & fechaHastaD

g_cdProducto = GF_PARAMETROS7("CdProducto", "", 6)

g_chkDetalle = GF_PARAMETROS7("chkDetalle", "", 6)


flagCall = false
if accion = ACCION_SUBMITIR then
	ret = GF_CONTROL_PERIODO(fechaDesdeD, fechaHastaD, fechaDesdeM, fechaHastaM, fechaDesdeA, fechaHastaA)
	Select case (ret)
		case 0	
			if (g_cdProducto <> "") then
				g_cdProducto = left(g_cdProducto,len(g_cdProducto)-1)
				flagCall=true
			else
				Call setError(PRODUCTO_REQUERIDO)
			end if			
		case 1
			Call setError(FECHA_INICIO_INCORRECTA)
		case 2
			Call setError(FECHA_FIN_INCORRECTA)
		case 3
			Call setError(PERIODO_ERRONEO)
	end select
end if
arrProductos = Split(g_cdProducto,",")


%>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
	<TITLE>Poseidon - Reportes - Descargas y embarques </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />	
	<link rel="stylesheet" href="../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
	<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">	
	<style type="text/css">
	.divListBox {
	    float: left;
	    height: 150px;
	    line-height: 24px;
	    margin: 0 1% 5px 0;
	    width: 23%;
	    font-family: Arial,Helvetica,sans-serif;
	    font-size: 12px;
	    font-weight: normal;
	}
	</style>	
	
<script type="text/javascript" src="../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script language="javascript">	
	
	
	<% if(flagCall)then %>
			window.open("reporteDescargaEmbarquesPrint.asp?pto=<%=g_strPuerto%>&fechaDesde=<%=fechaDesde%>&fechaHasta=<%=fechaHasta%>&cdProducto=<%=g_cdProducto%>&verDetalle=<%=g_chkDetalle%>");
	<% end if%>
	
	function bodyOnLoad() {
	}
	
	function submitInfo(accion)	{
		document.getElementById("accion").value = accion;		
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
	    calendar.setDateFormat("y-mm-dd");
	    calendar.showAtElement(elem);
	}
	function SeleccionarCalDesde(cal, date) {
		var str= new String(date);
		document.getElementById("dtFechaDesde").value = str;
	    document.getElementById("fechaDesdeD").value = str.substr(8,2);
	    document.getElementById("fechaDesdeM").value = str.substr(5,2);
	    document.getElementById("fechaDesdeA").value = str.substr(0,4);
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
	    document.getElementById("fechaHastaD").value = str.substr(8,2);
	    document.getElementById("fechaHastaM").value = str.substr(5,2);
	    document.getElementById("fechaHastaA").value = str.substr(0,4);    
		if (cal) cal.hide();	
	}	
	function QuitarFechaHasta(){
		document.getElementById("dtFechaHasta").value = "";
	    document.getElementById("fechaHastaD").value = "";
	    document.getElementById("fechaHastaM").value = "";
	    document.getElementById("fechaHastaA").value = "";	    
	}	
	function generarPDF(pAccion){
		var strProductos = "";
		 $('#cmbCdProducto option:selected').each(function(){
            strProductos = strProductos + $(this).val() + ",";
        });
        document.getElementById("cdProducto") .value = strProductos;
        submitInfo(pAccion);
	}
	function validarSeleccionados(e){
		if (e.value == 0){
			$('#cmbCdProducto option:selected ').each(function(){
				if ($(this).val() != 0) $(this).attr("selected", "");
			});
		}
		else{
			$('#cmbCdProducto option:selected ').each(function(){
				if ($(this).val() == 0) $(this).attr("selected", "");
			});
		}
	}
</script>
</HEAD>
<BODY onload="bodyOnLoad()">
<form name="frmSel" id="frmSel" action="reporteDescargaEmbarques.asp">
<div class="tableaside size100">
	<h3> Reporte Descarga y Embarques</h3>
	<div ><% Call showMessages() %></div>
    <div id="searchfilter" class="tableasidecontent">
	    <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Desde:")%> </div>
        <div class="col26">
   			<table>
				<tr><td>
				<input type="text" name="dtFechaDesde" id="dtFechaDesde" readonly onclick="javascript:MostrarCalendario('dtFechaDesde', SeleccionarCalDesde)" value="<% =fechaDesde %>" onChange="cambioBusqueda();">
				</td></tr>
				<input type="hidden" id="fechaDesdeD" name="fechaDesdeD" value="<%=fechaDesdeD%>">
				<input type="hidden" id="fechaDesdeM" name="fechaDesdeM" value="<%=fechaDesdeM%>">
				<input type="hidden" id="fechaDesdeA" name="fechaDesdeA" value="<%=fechaDesdeA%>">
			</table>
	    </div>
		<div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Hasta:")%> </div>
        <div class="col26">
   			<table>
				<tr><td><input type="text" name="dtFechaHasta" id="dtFechaHasta" readonly onclick="javascript:MostrarCalendario('dtFechaHasta', SeleccionarCalHasta)" value="<% =fechaHasta %>" onChange="cambioBusqueda();"></td></tr>
				<input type="hidden" id="fechaHastaD" name="fechaHastaD" value="<%=fechaHastaD%>">
				<input type="hidden" id="fechaHastaM" name="fechaHastaM" value="<%=fechaHastaM%>">
				<input type="hidden" id="fechaHastaA" name="fechaHastaA" value="<%=fechaHastaA%>">
			</table>
	    </div>		
	    <div class="col26 reg_header_navdos"> <%=GF_Traducir("Productos:")%> </div>
        <div class="divListBox">
			<table>
				<tr><td>
				<select id="cmbCdProducto" name="cmbCdProducto" size="8px" multiple="multiple" >
					<option value="0" <%if (isSelectedProducto(arrProductos,0)) then Response.Write "SELECTED"%> onclick="validarSeleccionados(this)"><%= GF_TRADUCIR("Todos...")%></option>
				<%	strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM DBO.PRODUCTOS ORDER BY DSPRODUCTO"
					call GF_BD_Puertos (g_strPuerto, rsProductos, "OPEN",strSQL)
					while not rsProductos.eof
						if (isSelectedProducto(arrProductos,rsProductos("CDPRODUCTO"))) then
							mySelected = "SELECTED"
						else
							mySelected = ""
						end if	%>
						<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%> onclick="validarSeleccionados(this)" ><%=rsProductos("DSPRODUCTO")%></option>
				<%		rsProductos.movenext
					wend %>
				</select>
				<input type="hidden" id="cdProducto" name="cdProducto" value="<%=g_cdProducto%>">
				</td></tr>
				<tr><td align="left">
					<%=GF_TRADUCIR("Seleccione los productos que desee")%>
				</td></tr>				
			</table>   			
	    </div>
	    <div class="col26 reg_header_navdos"> <%=GF_Traducir("Ver detalle:")%> </div>
        <div class="col26">
   			<input type="checkbox" id="chkDetalle" name="chkDetalle"  <%if (Ucase(g_chkDetalle) = "ON") then Response.Write " CHECKED " %>>
	    </div>
		<span class="btnaction input"><input type="button" value="Generar PDF" id=cmdSearch name=cmdSearch onclick="generarPDF('<% =ACCION_SUBMITIR %>');"></span>
		
	</div>
</div>

<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">	
<input type="hidden" id="Pto" name="Pto" value="<% =g_strPuerto %>">
</form>
</body>
</html>
