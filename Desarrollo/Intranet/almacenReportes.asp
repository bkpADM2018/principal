<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<% 'Call controlAccesoCM("CMADM") %>
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
	
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irA('almacenTableroDeControl.asp')");
		tb.addButton("PM_folder-16x16.png", "Ped. Materiales", "irA('almacenAdministrarPedidosMateriales.asp')");
		tb.addButton("refer_folder-16x16.png", "Remitos", "irA('almacenAdministrarREM.asp')");
		tb.draw();		
		pngfix();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Reportes de Almacenes") %>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteResumenConsumos.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Resumen de Consumos (Contable)") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga un resumen de consumos por Categoria o Partida Presupuestaria") %></td>
					</tr>
                </tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteArticulosConsumidos.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Artículos Consumidos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga un resumen de consumos de articulos mensual por almacen y usuario") %></td>
					</tr>
                </tbody>
				<tr><td>&nbsp;</td></tr>					
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteValesEmitidos.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Vales Emitidos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga todos los vales emitidos de un almacen ya sea en un periodo de tiempo, referidos a una obra, por articulo, por tipo de movimiento, etc.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>	
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteStock.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Stock de Articulos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga el stock actual de los articulos.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>	
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteCtaCteArt.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Cuenta Corriente X Articulo") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Conozca el detalle de los movimientos de un articulo en su almacen.") %></td>
					</tr>
				</tbody>	
                <tr><td>&nbsp;</td></tr>	
                <tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteArticulosPedidosNoRecibidos.asp?origen=ALMACENES';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Articulos pedidos no recibidos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga todos los articulos que fueron pedidos y aun no se recibieron.") %></td>
					</tr>
                </tbody>
                <tr><td>&nbsp;</td></tr>	
                <tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteRemitosCargados.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Remitos Cargados") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga todos los Remitos Cargados de un almacen especifico.") %></td>
					</tr>
                </tbody>   
                <tr><td>&nbsp;</td></tr>	
                <tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteArticulosNoConsumidos.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Articulos No Consumidos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Obtenga todos los articulos con o sin stock que no fueron consumidos en un periodo dado.") %></td>
					</tr>
                </tbody>                                
                <tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteVariacionPrecio.asp?origen=ALMACENES';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Variacion de Precios") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Lista las variaciones de precios de los artículos.") %></td>
					</tr>
				</tbody>
			</table>
		</td>
	</tr>
</table>

</body>
</html>