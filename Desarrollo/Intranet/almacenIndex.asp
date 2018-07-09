<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<title>Sistema de Almacenes</title>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
</head>
<body>
<table width="80%" height="100%" align="center">
	<tr valign="top">
		<td></td>
		<td>
			<table align="center" width="80%">		
			<tr><td width="45%">					
			<table>							
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenTableroDeControl.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Tablero de Control de Almacen") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Controle los movimientos de su almacen.") %></td>
					</tr>
				</tbody>								
				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAdministrarPedidosMateriales.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Pedido de Materiales") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Administre pedidos de materiales.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAdministrarREM.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Remitos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Ingrese y consulte los remitos entregados por los proveedores.") %></td>
					</tr>
				</tbody>

				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAdministrarVales.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Consulta de Vales") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Realice las consultas de vales necesarias. Además podrá realizar la anulación de los mismos.") %></td>
					</tr>
				</tbody>
				
				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAdministrarArticulosAlmacen.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Consulta de Stock") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Consulte el stock de los artículos en cualquiera de las almacenes de la empresa.") %></td>
					</tr>
				</tbody>
				
                <tr><td>&nbsp;</td></tr>
                <tbody  style="cursor: pointer" onClick="document.location.href='comprasAutorizaciones.asp?origen=Almacenes';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Autorizaciones Pendientes") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Asista a los controles de stock y autorice los mismos.") %></td>
					</tr>
				</tbody>
			</table>		
		</td>
		<td width="10%">&nbsp;</td>
		<td width="45%" style="vertical-align:top">
			<table>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAdministracion.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Administrar Maestros de Datos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree o modifique datos administrativos del sistema.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenReportes.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Reportes") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Obtenga diferentes tipos de reportes pertinentes a su almacen.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAuditoria.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Auditoria") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Controles y parametros que regulan la operatoria de la empresa.") %></td>
					</tr>
				</tbody>	
                <tr><td>&nbsp;</td></tr>
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenCCN_Contabilidad.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Contabilidad") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Realice los cierres contables, genere los asientos, consulte por cuenta y emita reportes.") %></td>
					</tr>
				</tbody>	
				<tr><td>&nbsp;</td></tr>				
				<tbody  style="cursor: pointer" onClick="document.location.href='almacenAjustes.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Ajustes al Almacen") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Confeccione ajustes de Stock, reclasifique y unifique sus articulos.") %></td>
					</tr>
				</tbody>                
 
			</table>
		</td>
	</tr>
</table>
</body>
</html>
