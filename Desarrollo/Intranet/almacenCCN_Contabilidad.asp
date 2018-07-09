<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<% Call initAccessInfo(RES_ACC_AL) %>
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
			location.href = "almacenCCN_ConsultasContables.asp";
		} else if(seccion == 1) {
			location.href = "almacenCCN_CierresContables.asp";
		} else if(seccion == 2) {
			location.href = "almacenCCN_ContabilidadControles.asp";
		}
	}
	function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
		tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		tb.draw();		
		pngfix();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Administración - Contabilidad") %>
<div id="toolbar"></div>
<br>
<table align="center" width="80%" height="100%">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">		
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(1)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/Cierres_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Cierre Contable") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Realice los cierres provisorios o definitivos del stock fisico y el stock contable, valorice sus articulos por mes y emita los pases contables.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>							
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(0)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/ConsultasContables_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Consultas Contables") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Consulte los cierres.") %></td>
					</tr>
				</tbody>	
				
				<!--
				<tr><td>&nbsp;</td></tr>				
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="fcnCall(2)">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/Control_Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Controles Contables") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Permite controlar la informacion visualizada en las consultas contables.") %></td>
					</tr>
				</tbody>
				-->
	
								
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenCCN_ReportesCierresContables.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/almacenes/printer_folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Reporte de Cuentas Contables") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Obtenga las variaciones mes a mes de las cuentas contables de inventario.") %></td>
					</tr>
				</tbody>
                <!--
                <tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenContabilidad.asp'">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/Almacenes/Contabilidad_Folder-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Contabilidad 2011/2012 - A partir de PICs") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Consulte la información histórica de los cierres realizados entre Enero de 2011 y Diciembre de 2012.") %></td>
					</tr>
				</tbody>								
				-->	
			</table>
		</td>
	</tr>
</table>

</body>
</html>