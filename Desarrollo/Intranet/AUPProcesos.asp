<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 22/01/2008                         -->
'<!--               Pagina     : AUPAuditoria.ASP                   -->
'<!--               Descripcion: Listado para Auditoria             -->
'<!--               Modificacion: Henzel Pavlo              -->
'<!--               Fecha      : 22/01/2008                         -->
'<!------------------------------------------------------------------->
%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<title>Intranet ActiSA - Definición de Procesos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="Scripts/iwin.js"></script>  
</head>
<script language="JavaScript">
	function closeWin(){
		var refPopUp = startIWin('listaProcesos');
		refPopUp.hide();
	}
</script>
<body>
<form name="frmMain" method="post">
<table width="100%" border=0 align="center" class="reg_header" cellpadding="2" cellspacing="1">
	<tr>
		<td class=reg_header_nav><%=GF_Traducir("Descripción")%></td>
		<td align="center" class=reg_header_nav><%=GF_Traducir("Nivel")%></td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actualiz. Tipos de Asientos")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Mantenim Tipos Ind. Inflac.")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Act. de Ind. de Ajus. Inflacc.")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Asiento Autom. de Ajuste Infl.")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Distribución de Gastos")%></td>			<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actualiz. y Confirm. Provis.")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Interfases de Subsistemas")%></td>		<td align="center">9</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Mantenim. Asientos Repetitivos")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Asiento Ajuste por Conversión")%></td>	<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Adición de Miembros Faltantes")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Definición de Seguridad")%></td>		<td align="center">9</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Eliminación de Movtos. por Ej.")%></td>	<td align="center">9</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Restauración de Movtos. por Ej.")%></td><td align="center">9</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consolidación de Compañias")%></td>		<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Resguardo de Archivos/Diario")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Resguardo de Archivos/Semanal")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Mantenim. de Matrices de Ref.")%></td>	<td align="center">7</td>
	</tr>


	<tr>
		<td><%=GF_Traducir("Llamar al SQL")%></td>					<td align="center">1</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Ingreso de Transacciones")%></td>		<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actualiz. de Tipos de Cambio")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta a numeración Asientos")%></td>	<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta a asientos")%></td>			<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta al Mayor por Cuenta")%></td>	<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta a Saldos por Cuenta")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta por Centro de Costos")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta a Saldos por Grupos")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta Comparativa de Saldos")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Consulta Form. para Cuadr.")%></td>		<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Análisis de Diferenc.de Cambio")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Emisión del Libro Diario")%></td>		<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Emisión del Libro Mayor")%></td>		<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Emis. Mayor (todos los libros)")%></td>	<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Listado de Saldos por Cuenta")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Listado de Saldos por C.Costos")%></td>	<td align="center">5</td>
	</tr>


	<tr>
		<td><%=GF_Traducir("Listado de Saldos Comparativo")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Listado de Fórmulas")%></td>			<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Listado Diferencias de Cambio")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Balance Gral(todos los libros)")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Apertura de Ejercicio Contable")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Cierre de Ejercicio Contable")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Habilitación de Meses")%></td>			<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Cierre Mensual")%></td>					<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Reapertura de Meses Cerrados")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Act. Plan de Cuentas General")%></td>	<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Control ctas. bal. Cons.")%></td>		<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Act. Plan de Cuentas General")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Control Ctas. Balance sheet")%></td>	<td align="center">4</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Mantenim. de Centro de Costos")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actualización a Compañias")%></td>		<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actualiz. de Tipos de Monedas")%></td>	<td align="center">7</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Actual. de Unidades de Medida")%></td>	<td align="center">7</td>
	</tr>
	
	
	
	<tr>
		<td><%=GF_Traducir("Consulta a Referencias")%></td>			<td align="center">6</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Compensación/Modif. de Refer.")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Eliminación de Movtos.Compens.")%></td>	<td align="center">8</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Gastos vs Budget")%></td>				<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Budget Control")%></td>					<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Gastos por C.C.")%></td>				<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Trabajar con Presupuesto")%></td>		<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Trabajar con conceptos de gas")%></td>	<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Control Presupuestario")%></td>			<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Selección de Compañia")%></td>			<td align="center">5</td>
	</tr>
	<tr>
		<td><%=GF_Traducir("Interfase Tesoreria (Export.)")%></td>	<td align="center">5</td>
	</tr>
</table>	

</form>
</body>
</html>