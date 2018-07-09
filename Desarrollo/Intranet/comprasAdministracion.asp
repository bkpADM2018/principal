<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<% 
Call comprasControlAccesoCM(RES_ADM) 
%>
<html>
<head>
<title>Sistema de Compras</title>
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
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}
	
	function irObras() {
		location.href = "comprasObras.asp";
	}
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	function fcnCall(seccion) {
		location.href = "comprasSecciones.asp?seccion=" + seccion
	}
	function irDirecta(){
		location.href = "comprasAdministrarCotizaciones.asp";
	}
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}	
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 6, "images/compras/");	
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButton("OBR-16X16.png", "Obras", "irObras()");				
		tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");
		tb.addButton("Direct_purchase-16x16.png", "Directa", "irDirecta()");
		tb.draw();		
		pngfix();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% 'call GF_TITULO2("kogge64.gif","Administración") 
%>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(0)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/users_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>
						<td class="title_sec_section"><% =GF_TRADUCIR("Responsables") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Administre los usuarios autorizados a participar de la apertura de sobre y a firmar las cotizaciones.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(1)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/categories_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>
						<td class="title_sec_section"><% =GF_TRADUCIR("Categorias") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree las categorias que le permitiran agrupar sus articulos y administrarlos más facilmente.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(2)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/units_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Unidades de Medida") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Conozca cuales son y como se relacionan las unidades en las que se miden los articulos de las almacenes.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(3)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/items_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Articulos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Administre los articulos que adquiere la empresa para sus obras y mantenimiento.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr>
				<!--
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(4)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/Company_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Empresas") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree o modifique datos de las empresas con las cuales trabaja nuestra empresa.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr> -->
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(5)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/budget_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Presupuestos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion" valign=top><% =GF_TRADUCIR("Administre las diferentes áreas que componen un presupuesto y los detalles de cada una de estas.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasListasCorreo.asp';">
					<tr>
					  <td rowspan="2"><img align="absMiddle" src="images/compras/mail-folder-48x48.png"></td>
					  <td rowspan="2">&nbsp;</td>
					  <td class="title_sec_section"><% =GF_TRADUCIR("Listas de Correos") %></td>
				  </tr>
				  <tr>
					  <td class="textoSeccion"><% =GF_TRADUCIR("Defina listas de correo y asocie usuarios a ellas.") %></td>
				  </tr>
				</tbody>							
			</table>
		</td>
	</tr>
</table>

</body>
</html>