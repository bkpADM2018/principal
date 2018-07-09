<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<%

Dim strSQL, conn,rs,nroPagina,rsAlmacenes
Dim accion,categoria,metodo,almacen, fechaBusqueda


'******************************************************
'					INICIO DE LA PAGINA
'******************************************************
	Set	rsAlmacenes = obtenerListaAlmacenesSolicitud()
	if (rsAlmacenes.eof) then response.redirect "comprasAccesoDenegado.asp"
	fechaBusqueda = GF_FN2DTE(left(session("MmtoSistema"), 8))

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>Reporte de Stock</title>
	
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">	
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
	<script type="text/javascript" src="scripts/Toolbar.js"></script>	
	<script type="text/javascript" src="scripts/date.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	
	<script type="text/javascript">		
	var ch = new channel();
	var myDate = formatDate(new Date(),"dd/MM/yyyy");
	var params = "";
		
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("../DocumentoTexto-16x16.png", "Imprimir PDF", "GenerarInfo('PDF')");
		tb.addButton("../excel3.gif", "Imprimir XLS", "GenerarInfo('XLS')");
		tb.addButton("Previous-16x16.png", "Volver", "cerrar()");
		tb.draw();
	}
	function cerrar(){
		location.href='almacenReportes.asp'
	}
	function GenerarInfo(pTipo) {
		document.getElementById("actionLabel").style.visibility = 'visible';		
		var metodo = 0;
		if (document.getElementById("metodo").checked) metodo = 1;
		var categoria = document.getElementById("categoria").value;
		var almacen = document.getElementById("almacen").value;
		var valorizar = document.getElementById("valorizar").checked;
		var incluir = document.getElementById("incluir").checked;
		var fecha = document.getElementById("fecha").value;
		params = "?accion=<% =ACCION_PROCESAR %>&metodo=" + metodo + "&categoria=" + categoria;
		params = params + "&almacen=" + almacen + "&valorizar=" + valorizar + "&incluir=" + incluir;
		params = params + "&fechaBusqueda=" + fecha;		 		
 		document.getElementById('msgproceso').innerHTML = '<% =GF_TRADUCIR("Calculando Stock a la Fecha") %>';
		ch.bind("almacenReporteStockAjax.asp" + params, "printReport('" + pTipo + "')");
		ch.send();
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
	
	
	function SeleccionarCal(cal, date) {
		var str= new String(date);
		if (myDate!=''){
			var rtrn = compareDates(str,"dd/MM/yyyy", myDate,"dd/MM/yyyy")
			if (rtrn == 1){
				alert("La fecha no puede ser mayor a hoy!");
				str = myDate;
			}
		}
		document.getElementById("fechaDiv").innerHTML = str;
		document.getElementById("fecha").value = str;
		if (cal) cal.hide();	
	}

	function printReport(tipo) {
 		if (tipo == 'PDF') {
 			document.getElementById('msgproceso').innerHTML = '<% =GF_TRADUCIR("Generando Reporte en PDF") %>';
			location.href = "almacenReporteStockPrint.asp" + params;
		} else if (tipo == 'XLS') {
			document.getElementById('msgproceso').innerHTML = '<% =GF_TRADUCIR("Generando Reporte en Excel") %>';
			location.href = "almacenReporteStockPrintXLS.asp" + params;			
			document.getElementById("actionLabel").style.visibility = 'hidden';
		}
	}
	function cambiarMsg(pObj){
		if (pObj.checked){
			document.getElementById("msg").innerHTML = "Datos calculados al final de la fecha seleccionada.";
		}
		else{
			document.getElementById("msg").innerHTML = "Datos calculados al inicio de la fecha seleccionada.";		
		}
	
	}
	</script>    

</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Reporte de Stock de Articulos") %>
	<div id="toolbar"></div>
	<br>
	<form name="frm" id="frm" action="almacenReporteStock.asp" method="get">		
		
			<table class="reg_Header" width="70%" align="center" border="0">
				<tr>
				  <td colspan="5" class="reg_Header_nav round_border_top_left round_border_top_right">
						<font class="big"><%=GF_Traducir("Reporte de stock del almacen")%></font>
					</td>
				</tr>
				<tr>
					<td width="10%" class="reg_Header_navdos">
						Categoria
					</td>
					<td width="20%" colspan="3">
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
				 	<td width="25%">
						<input style="cursor:pointer;" type="radio" id="metodo" name="metodo" value="1" checked="checked"> Con Stock
						<input style="cursor:pointer;" type="radio" id="metodo" name="metodo" value="0"> Todos
					</td>	
				</tr>			
		
				<tr>
					<td class="reg_Header_navdos"><% =GF_TRADUCIR("Almacen") %></td>
					<td colspan="3">
							<select name="almacen" id="almacen">
								<%

							while not rsAlmacenes.eof
							%>
								<option value="<%=rsAlmacenes("idAlmacen")%>"><%=rsAlmacenes("dsalmacen")%></option>
							<%
								rsAlmacenes.movenext
							wend%>
						</select>
					</td>
                    <td>
                    	<div id="divValorizar"><input style="cursor:pointer;" type="checkbox" name="valorizar" id="valorizar"> Valorizar</div>
                    </td>
				</tr>
				<tr>
					<td class="reg_Header_navdos"><% =GF_TRADUCIR("Stocks al") %></td>
					<td align="center">
						<div id="fechaDiv"><% =fechaBusqueda %></div>
						<input type="hidden" id="fecha" name="fecha" value="<% =fechaBusqueda %>">
					</td>
					<td align="left"><a href="javascript:MostrarCalendario('img_fecha', SeleccionarCal)"><img id="img_fecha" src="images/date.gif"></a></td>
					<td align="left" colspan="1"><div id="msg"><% =GF_TRADUCIR("Datos calculados al inicio de la fecha seleccionada.") %></div></td>
                    <td>
                    	<div id="divIncluir"><input style="cursor:pointer;" onclick="cambiarMsg(this)" type="checkbox" name="incluir" id="incluir">Incluir dia final</div>
                    </td>					
				</tr>
			</table>
			<div width="70%" align="center"><div id="actionLabel" class="round_border_bottom TDSUCCESS" style="width:70%;visibility:hidden;"><label id="msgproceso"></label>...</div></div><br>
	</form>
</body>
</html>
