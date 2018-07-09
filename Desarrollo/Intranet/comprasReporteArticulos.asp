<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->


<%
Const REPORT_PDF = "PDF"
Const REPORT_XLS = "XLS"
Function crearSelect(byref rs,pNombre,pValor,pTexto)
	Dim rtrn
	rtrn = "<select class='selects' id='" & pNombre & "' name='" & pNombre & "'>"
	
	while not rs.eof
		rtrn = rtrn & "<option " 
		if ( cstr(rs(pValor))=GF_PARAMETROS7(pNombre,"",6) ) then
			rtrn = rtrn & "selected='selected'"
		end if
		rtrn = rtrn & "value=" & rs(pValor) & ">" & rs(pTexto) & "</option>"
		rs.MoveNext
	wend
	
	rtrn = rtrn & "</select>"
	crearSelect = rtrn
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************

dim RPT_FechaDesde, RPT_FechaHasta, division,picSearch_radioMoneda,verPagosEfectuados

RPT_FechaDesde = GF_FN2DTE(Left(session("MmtoDato"),8))
RPT_FechaHasta = GF_FN2DTE(Left(session("MmtoDato"),8))
picSearch_radioMoneda = GF_PARAMETROS7("radio_TipoMoneda","",6)

%>
<html>
<head>
<title><%=GF_TRADUCIR("Compras - Compras por Articulo")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
</style>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	
	var ch = new channel();		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("../pdf-16.png", "Imprimir PDF", "validarDatos('<%=REPORT_PDF%>')");
		tb.addButton("../excel-16.png", "Imprimir XLS", "validarDatos('<%=REPORT_XLS%>')");
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		tb.draw();			
				
		var msArticulo = new MagicSearch("", "articuloItem0", 30, 4, "comprasStreamElementos.asp?tipo=articulos&linea=0&all=1");
		msArticulo.setToken(";");
		msArticulo.onBlur = seleccionarArticulo;		
					
		pngfix();
	}

	function validarDatos(pTipoReporte) {		
		var auxArt = document.getElementById("idArticulo").value;
		if (auxArt =='' || auxArt == 0){	
			var auxCat = document.getElementById("idCategoria").value;
			if (auxCat == 0){
				alert("Debe ingresar un articulo o seleccionar una categoria!");
			}
			else {
				generarReporte(pTipoReporte)
			}
		} else {
			generarReporte(pTipoReporte)
		}
	}
	function generarReporte(pTipoReporte){
		if (pTipoReporte == '<%=REPORT_PDF%>') {
			document.getElementById("frmSel").action = "comprasReporteArticulosPrint.asp";
			document.getElementById("frmSel").submit();
		}
		else {
			document.getElementById("frmSel").action = "comprasReporteArticulosPrintXLS.asp";
			document.getElementById("frmSel").submit();
		}	
	}
	
	function SeleccionarCalEmision(cal, date) {
		//Controlar que la fecha desde no sea mayor a la fecha hasta
		var str= new String(date);		
		var auxDate = document.getElementById("closingdate").value;
		if (auxDate!=''){
			var rtrn = compareDates(str,"dd/MM/yyyy", auxDate,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha desde no puede ser mayor a la fecha hasta!");
				str = auxDate;
			}
		}
		document.getElementById("issuedateDiv").innerHTML = str;
		document.getElementById("issuedate").value = str;
		if (cal) cal.hide();	
	}
	function SeleccionarCalLimite(cal, date) {
		//Controlar que la fecha hasta no sea menor a la fecha desde
		var str= new String(date);	
		var auxDate = document.getElementById("issuedate").value;
		if (auxDate!=''){
			var rtrn = compareDates( auxDate,"dd/MM/yyyy", str,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha hasta no puede ser menor a la fecha desde!");
				str = auxDate;
			}
		}			
		document.getElementById("closingdateDiv").innerHTML = str;
	    document.getElementById("closingdate").value = str;
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
			
	function volver() {	
		location.href = "comprasReportes.asp";
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Compras por Articulos") %>
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" method="POST" target="_blank">	
<table class="reg_Header" id="TAB1" align="center" width="80%" border="0">				
	<tr>
		<td class="reg_Header_nav" align="left" colspan="6">
			<font class="big"><%=GF_Traducir("Reporte de Compras por Articulos")%></big>
		</td>
	</tr>
	<tr>
		<!--Desde / Hasta-->
		<td class="reg_Header_navdos">
			<% =GF_TRADUCIR("Desde") %>
		</td>
		<td align="center" width="20%">
			<div id="issuedateDiv"><% =RPT_FechaDesde %></div>															
			<input type="hidden" id="issuedate" name="issuedate" value="<% =RPT_FechaDesde %>">
		</td>
		<td align="left" width="15%">
			<a href="javascript:MostrarCalendario('imgEmision', SeleccionarCalEmision)"><img id="imgEmision" src="images/DATE.gif"></a>
		</td>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Articulo") %></td>
		<td colspan="2">
			<div id="articuloItem0"></div>																		
			<input type="hidden" id="idArticulo" name="idArticulo">
		</td>
	</tr>	
	<tr>
		<td class="reg_Header_navdos">
			<% =GF_TRADUCIR("Hasta") %>
		</td>
		<td align="center" width="22%">
			<div id="closingdateDiv"><% =RPT_FechaHasta %></div>	
			<input type="hidden" id="closingdate" name="closingdate" value="<% =RPT_FechaHasta %>">					
		</td>
		<td align="left" width="15%">
			<a href="javascript:MostrarCalendario('imgLimite', SeleccionarCalLimite)"><img id="imgLimite" src="images/DATE.gif"></a>
		</td>	
		
		 <td width="10%" class="reg_Header_navdos"><%=GF_Traducir("Division")%></td>
         <td width="41%"><%
                            strSQL = "select iddivision id,dsdivision ds from tbldivisiones"
                            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
                            response.write crearSelect(rs,"division","id","ds") 
                            %>
		 </td>
	</tr>	
	<tr>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Moneda") %></td>
		<td colspan="2">
			<input type="radio" name="radio_TipoMoneda" id="radio_TipoMoneda" value="$" <%if (picSearch_radioMoneda = TIPO_MONEDA_PESO) then %>checked="checked"<%end if%> / ><% = GF_TRADUCIR("$")%>
			<input type="radio" name="radio_TipoMoneda" id="radio_TipoMoneda" value="US$" <%if (picSearch_radioMoneda = TIPO_MONEDA_DOLAR) then %>checked="checked"<%end if%> /><% = GF_TRADUCIR("US$")%>
			
		</td>	
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Categoria") %></td>
		<td colspan="2">
			<%	Set sp_Div = executeProcedureDb(DBSITE_SQL_INTRA, rsCat, "TBLARTCATEGORIAS_GET_BY_ESTADO", ESTADO_ACTIVO) %>
			<select style="z-index:-1;" name="idCategoria" id="idCategoria">
		        <option SELECTED value="0">- <% =GF_TRADUCIR("Seleccione") %> -
				<%	while (not rsCat.eof)
						selected = ""
						if (CLng(rsCat("IDCATEGORIA")) = CLng(g_idCategoria)) then selected = "selected" %>
						<option value="<% =rsCat("IDCATEGORIA") %>" <% =selected %>><% =rsCat("DSCATEGORIA") %>
				<%		rsCat.MoveNext()
					wend	%>
			</select>
		</td>	
	</tr>
	<tr>
		<td class="reg_Header_navdos" width="15%"><% =GF_TRADUCIR("Ver detalle") %></td>
		<td colspan="2">
			<input type="checkbox" id="verPagosEfectuados" name="verPagosEfectuados"  <%if (Ucase(g_chkDetalle) = "ON") then Response.Write " CHECKED " %>>
		</td>
	</tr>
	
</table>
</form>
</body>
</html>
