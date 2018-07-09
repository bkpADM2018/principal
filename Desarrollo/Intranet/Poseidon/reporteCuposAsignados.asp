<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientospuertos.asp"-->


<%
Const REPORTE_CUPOS_ASIGNADOS  = 1
Const REPORTE_CUPOS_RESOLUCION = 2
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Dim pto,accion,flagCall,cuposEspeciales,cuposFechaDesde,cdProducto,fecha,tipoCupo

Call GP_CONFIGURARMOMENTOS()

pto			    = GF_PARAMETROS7("pto", "", 6)
accion			= GF_PARAMETROS7("accion", "", 6)
cuposEspeciales = GF_PARAMETROS7("chkCuposEspeciales", 0, 6)
cdProducto		= GF_PARAMETROS7("cdProducto", 0, 6)
fecha			= GF_PARAMETROS7("fecha",0,6)
fechaRes		= GF_PARAMETROS7("fechaRes",0,6)
cupo			= GF_PARAMETROS7("cupo",0,6)
media			= GF_PARAMETROS7("media", "", 6)

if fecha = 0 then fecha = CLng(Left(session("MmtoDato"),8)) 
if fechaRes = 0 then fechaRes = CLng(Left(session("MmtoDato"),8)) 
flagCall = false
if (accion = ACCION_CONTROLAR) then	
	Select case cupo		
		case REPORTE_CUPOS_RESOLUCION						
			flagCall = true
		case REPORTE_CUPOS_ASIGNADOS 			
			if (cdProducto > 0) then
				flagCall = true												
			else
				setError(PRODUCTO_REQUERIDO)
			end if		
	end select
	
end if

%>
<HTML>
<HEAD>
<title><%=GF_TRADUCIR("Puertos - Cupos Asignados")%></title>
<link rel="stylesheet" href="../css/ActisaIntra-1.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">

<style type="text/css">

.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}

#toolbarResolucion{
	width: 400px;
	float: left;
}
#toolbarAsignados{
	width: 400px;
	float: right;
}

#showErrorCupos{
	width: 400px;
	float: right;
}
</style>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/script_fechas.js"></script>
<script type="text/javascript" src="../scripts/iwin.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript">	

	
	
    <% if((flagCall)and(cupo = REPORTE_CUPOS_ASIGNADOS ))then %>    
		window.open("reporteCuposAsignadosPrint.asp?pto=<%=pto%>&fecha=<%=fecha%>&cdProducto=<%=cdProducto%>&chkCuposEspeciales=<%=cuposEspeciales%>&media=<% =media %>");
	<% end if %>
	<% if((flagCall)and(cupo = REPORTE_CUPOS_RESOLUCION))then %>			
		window.open("reporteCuposResolucion25Print.asp?pto=<%=pto%>&fecha=<%=fechaRes%>");
	<% end if%>
	
	function bodyOnLoad() {
		tb = new Toolbar('toolbarAsignados', 2,'../images/');		
		tb.addButton("print-16.png", "Imprimir XLS", "submitir(<%=REPORTE_CUPOS_ASIGNADOS%>, '')");
        tb.addButton("see-16.png", "Ver cupos Online", "submitir(<% =REPORTE_CUPOS_ASIGNADOS%>, '<% =TIPO_AFIRMACION %>')");
		tb.draw();
		tb = new Toolbar('toolbarResolucion', 2,'../images/');		
		tb.addButton("DocumentoTexto-16x16.png", "Imprimir PDF", "submitir(<%=REPORTE_CUPOS_RESOLUCION%>, '')");
		tb.draw();
		pngfix();		
	}
    	
	function CerrarCal(cal) {
		cal.hide();		
	}
	
	function SeleccionarCal(cal, date) {
		var str= new String(date);
		var anio = str.substring(6);
		var mes = str.substring(3,5);
		var dia = str.substring(0,2);
		var fechaHoy = new Date();
		var fecha = new Date(date);
		document.getElementById("fechaDesdeDiv").innerHTML = str;
		document.getElementById("fecha").value = parseInt(anio + mes + dia);	
		if (cal) CerrarCal(cal);			
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
	
	function CerrarCal(cal) {
		cal.hide();		
	}
	
	function SeleccionarCalRes(cal, date) {
		var str= new String(date);
		var anio = str.substring(6);
		var mes = str.substring(3,5);
		var dia = str.substring(0,2);
		var fechaHoy = new Date();
		var fecha = new Date(date);
		document.getElementById("fechaDesdeDivRes").innerHTML = str;
		document.getElementById("fechaRes").value = parseInt(anio + mes + dia);	
		if (cal) CerrarCal(cal);			
	}
	
	function MostrarCalendarioRes(p_objID, funcSel) {
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
	
	function submitir(tipo, acc){						
	    document.getElementById('media').value = acc;
		document.getElementById('cupo').value = tipo;		
		document.getElementById('frmSel').submit();
	}
	
	
</script>
</HEAD>
<BODY onLoad="bodyOnLoad()">	
<DIV id="toolbar"></DIV>
<!--***************************************************CUPOS ASIGNADOS******************************************-->
<FORM id="frmSel" name="frmSel" method="POST" action="reporteCuposAsignados.asp">
	
	<BR></BR>			
	<table id="cupos" name="cupos" width="400px" align="center">
		<tr>
			<td>
				<div id="showErrorCupos">				
					<% Call showErrors() %>
				</div>			
				<DIV id="toolbarAsignados"></DIV>					
				
				<TABLE class="reg_Header" id="TAB1" align="center" width="400px" border="0">				
					<TR>
						<TD class="reg_Header_nav" align="left" colspan="3">
							<FONT class="big"><%=GF_Traducir("Reporte de Cupos Asignados")%></BIG>
						</TD>
					</TR>
					<TR>
						<TD class="reg_Header_navdos"  width=50%>
							<% =GF_TRADUCIR("Fecha") %>
						</TD>
						<TD align=right width=30%>
							<DIV id="fechaDesdeDiv"><% =GF_FN2DTE(fecha) %></DIV>
							<input type="hidden" id="fecha" name="fecha" value="<% =fecha %>">
						</TD>
						<TD align=right width=20% >
							<A href="javascript:MostrarCalendario('imgEmision', SeleccionarCal)"><IMG id="imgEmision" src="images/DATE.gif"></A>
						</TD>
					</TR>
					<TR>		
						<TD class="reg_Header_navdos"  width=50%>
							<% =GF_TRADUCIR("Producto") %>
						</TD>
						<TD align=left colspan=2>
						<%	strSQL = "SELECT * FROM PRODUCTOS ORDER BY DSPRODUCTO"
							call GF_BD_Puertos(pto, rsProductos, "OPEN",strSQL)
						%>
							<SELECT name="cdProducto" value="<%=cdProducto%>">
								<OPTION value="0"> <%=GF_Traducir("Seleccionar...")%></OPTION>
						<%		while not rsProductos.eof
									mySelected = ""
									if cint(rsProductos("CDPRODUCTO")) = cint(cdProducto) then mySelected = "SELECTED"	%>				
									<OPTION value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>> <%=rsProductos("DSPRODUCTO")%></OPTION>
						<%			rsProductos.movenext
								wend  %>
							</SELECT>
						</TD>
					</TR>
					<TR>		
						<TD class="reg_Header_navdos"  width=50%>
							<% =GF_TRADUCIR("Cupos Especiales") %>
						</TD>
						<TD align=left colspan=2>				
							<INPUT style="border:none;cursor:pointer;" type="checkbox" id="chkCuposEspeciales" name="chkCuposEspeciales" value=1 checked >&nbsp&nbsp<% =GF_TRADUCIR("Incluir detalle Cupos especiales") %>
						</TD>
					</TR>		
				</TABLE>			
			<BR></BR>	
			</td>
		</tr>
		<tr>
			<td>			
<!--***************************************************CUPOS RESOLUCION******************************************-->			
				<DIV id="toolbarResolucion"></DIV>
				<TABLE class="reg_Header" id="TAB2" align="center" width="400px" border="0">				
					<TR>
						<TD class="reg_Header_nav" align="left" colspan="3">
							<FONT class="big"><%=GF_Traducir("Reporte de Cupos para Resolución 25/13")%></BIG>						
						</TD>
					</TR>
					<TR>
						<TD class="reg_Header_navdos"  width=50%>
							<% =GF_TRADUCIR("Fecha") %>
						</TD>
						<TD align=right width=30%>
							<DIV id="fechaDesdeDivRes"><% =GF_FN2DTE(fechaRes) %></DIV>
							<input type="hidden" id="fechaRes" name="fechaRes" value="<% =fechaRes %>">
						</TD>
						<TD align=right width=20% >
							<A href="javascript:MostrarCalendarioRes('imgEmisionRes', SeleccionarCalRes)"><IMG id="imgEmisionRes" src="images/DATE.gif"></A>
						</TD>
					</TR>								
				</TABLE>	
				<BR>	
			</td>
		</tr>
	<!--**********************************************************************************************************-->		
	</table>
	<INPUT type=hidden name=accion id=accion value=<%=ACCION_CONTROLAR%> >
	<INPUT type=hidden name=pto id=pto value=<%=pto%> >
	<INPUT type=hidden name=cupo id=cupo >
    <INPUT type=hidden name="media" id="media" value="">
</FORM>
</BODY>
</HTML>
