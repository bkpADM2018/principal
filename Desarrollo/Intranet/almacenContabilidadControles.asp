<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<% Call controlAccesoAL(RES_ACC_AL) %>
<html>
<head>
<title>Almacenes</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<style type="text/css">
.title_sec_section {
	font-size: 12px;
	font-weight: bold;
}
.textoSeccion {
	font-size: 12px;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	function irA(pLink) {
		location.href = pLink;
	}
	
	function fcnCall(seccion) {
		if (seccion == 0){
			location.href = "almacenControlesContables.asp";
		} else if(seccion == 1) {
			location.href = "almacenControlCruzadoContable.asp";
		} else if(seccion == 2) {
			location.href = "almacenControlesContables.asp";
		}
	}
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.addButton("Contabilidad_Folder-16x16.png", "Contabilidad", "irA('almacenContabilidad.asp')");				
		tb.draw();		
		pngfix();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Administración - Consultas Contables") %>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">		
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(0)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/itemsSP_Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Articulos sin Precio") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Permite visualizar por cierre si hubo articulos que se movieron en el mes pero para el cual no se pudo obtener un precio.") %></td>
					</tr>
				</tbody>	
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(1)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/itemsIC_Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Informacion cruzada") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Permite comparar los saldos de las cuentas de inventario del cierre con la suma de todos los movimientos de vales de los articulos.") %></td>
					</tr>
				</tbody>	
			</table>
		</td>
	</tr>
</table>

</body>
</html>