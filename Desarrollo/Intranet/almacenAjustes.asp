<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<% 
'Call controlAccesoCM("CMADM")
%>
<html>
<head>
<title>Almacenes</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
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

<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
	function irA(pLink) {
		location.href = pLink;
	}
	
	function fcnCall(seccion) {
		if(seccion == 2) {
			location.href = "almacenValesTitulo.asp?TC=0&cdVale=<%=CODIGO_VS_AJUSTE_STOCK%>"
		} else if(seccion == 4) {
			location.href = "almacenValesVRS.asp";
		} else if(seccion == 5) {
			location.href = "almacenCambioUnidad.asp";
		} 		
	}
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.addButton("Control_panel_folder-16x16.png", "Tablero", "irA('almacenTableroDeControl.asp')");
		tb.draw();		
		pngfix();
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Administración Almacenes - Consulta y Ajustes al Almacen") %>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">		
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(2)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/Setting_A_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Ajustes de Stock") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Realice los ajustes necesarios para mantener el stock de sus articulos actualizado.") %></td>
					</tr>
				</tbody>	
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(4)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/VRS-Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Reclasificación de Stock") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Reclasifique el stock de sus articulos para mantenerlos actualizados.") %></td>
					</tr>
				</tbody>		
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(5)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/ChangeUnit-Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Cambio de Unidad de Medida") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Actualice las unidades de medida que utiliza para cada articulo.") %></td>
					</tr>
				</tbody>
			</table>
		</td>
	</tr>
</table>

</body>
</html>