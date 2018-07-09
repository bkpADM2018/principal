<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<%
 Call initAccessInfo(RES_ADM_AL) 
 %>
 
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
		location.href = "almacenSecciones.asp?seccion=" + seccion
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
<% call GF_TITULO2("kogge64.gif","Administración Almacenes - Maestro de datos") %>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(0)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/warehouses_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Almacenes") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Administre los almacenes disponibles en su puerto.") %></td>
					</tr>
				</tbody>	
				<tr><td>&nbsp;</td></tr>								
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(1)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/categories_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>
						<td class="title_sec_section"><% =GF_TRADUCIR("Categorias") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Cree las categorias que le permitiran agrupar sus articulos y administrarlos más facilmente.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(2)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/units_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Unidades de Medida") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Conozca cuales son y como se relacionan las unidades en las que se miden los articulos de las almacenes.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(3)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/items_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Articulos") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Administre los articulos que adquiere la empresa para sus obras y mantenimiento.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(4)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/users_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>
						<td class="title_sec_section"><% =GF_TRADUCIR("Responsables") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Administre los usuarios autorizados a participar de la apertura de sobre y a firmar las cotizaciones.") %></td>
					</tr>
				</tbody>
			</table>
		</td>
	</tr>
</table>

</body>
</html>