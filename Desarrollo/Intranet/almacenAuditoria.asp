<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<% 
'Call comprasControlAccesoCM(RES_AUD)

Const ESTADO_TODOS = -1
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
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Auditoria para Almacenes") 
%>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">		
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasNormasAuditoria.asp'">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/condiciongrl-50.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>
						<td class="title_sec_section"><% =GF_TRADUCIR("Normas de Auditoria") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree las reglas que regulan las compras realizadas por la empresa.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr>			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenAdministrarCtSt.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/search-50.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Control de Stock") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Administre y realice sus controles de stock.") %></td>
					</tr>
                </tbody>
                <tr><td>&nbsp;</td></tr>			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasProveedoresCD.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/add-user-50.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Proveedores Pre-autorizados para Compras Directas") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Maestro de Proveedores que no requieren firmas adicionales ante compras directas que superen los l&iacutemites establecidos por auditor&iacutea.") %></td>
					</tr>
                </tbody>
                <tr><td>&nbsp;</td></tr>			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='proveedores/controlCorrelatividadCbtes.asp';">
                	<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/invoice-50.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Control de Facturaci&oacuten Correlativa") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Verifique el grado de dependencia que existe entre los proveedores y la empresa.") %></td>
					</tr>
                </tbody>
				<tr><td>&nbsp;</td></tr>
			</table>
		</td>
	</tr>
</table>

</body>
</html>