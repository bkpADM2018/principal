<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<title>Sistema de Compras</title>
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
				<tbody style="cursor: pointer" onClick="document.location.href='comprasAdministrarPedidos.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Compras - Pedidos de Precio") %></td>						
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Acceda a todos los pedidos de precios organizados por la empresa.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr>				
				<tbody style="cursor: pointer" onClick="document.location.href='comprasAdministrarCotizaciones.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Compras - PICs") %></td>
					</tr>			
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Conozca y administre los Pedidos Internos de Compra.") %></td>
					</tr>		
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody style="cursor: pointer" onClick="document.location.href='comprasAdministrarAFEs.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Autorizaciones para Gastos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Controle la emisión de autorizaciones para los gastos planificados por la empresa.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr>				
				<tbody style="cursor: pointer" onClick="document.location.href='comprasAutorizaciones.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Autorizaciones pendientes") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Asista a las aperturas de sobre, controle y apruebe las adjudicaciones y autorice los gastos de cada uno de los pedidos.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>				
				<tbody style="cursor: pointer" onClick="document.location.href='comprasObras.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Administrar Partidas Presupuestarias de Mantenimiento e Inversiones") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree y controle las partidas presupuestarias de mantenimiento e inversiones de todas las divisiones.") %></td>
					</tr>
				</tbody>
			</table>		
		</td>
		<td width="10%">&nbsp;</td>
		<td width="45%" style="vertical-align:top">
			<table>
				<tbody style="cursor: pointer" onClick="document.location.href='comprasCTCAdministrar.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Contratos y Servicios Repetitivos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Conozca y administre los servicios contratados y las obras en ejecución.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody style="cursor: pointer" onClick="document.location.href='comprasPDCAdministrar.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Polizas de Caucion") %></td>
					</tr>
					<tr>  
						<td class="textoSeccion"><% =GF_TRADUCIR("Conozca y administre los pagos de los adelantos.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody style="cursor: pointer" onClick="document.location.href='almacenAuditoria.asp'">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Auditoria") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Controles y parametros que regulan la oparatoria de la empresa.") %></td>
					</tr>
				</tbody>				
				<tr><td>&nbsp;</td></tr>
				<tbody style="cursor: pointer" onClick="document.location.href='comprasAdministracion.asp';">
					<tr>
						<td class="title_sec_section"><% =GF_TRADUCIR("Administrar Maestros de Datos") %></td>
					</tr>
					<tr>
						<td class="textoSeccion"><% =GF_TRADUCIR("Cree o modifique datos administrativos del sistema.") %></td>
					</tr>
                </tbody>
                    <tr><td>&nbsp;</td></tr>
                <tbody style="cursor: pointer" onClick="document.location.href='comprasReportes.asp';">
					<tr>
					  <td class="title_sec_section"><% =GF_TRADUCIR("Reportes") %></td>
				  </tr>
				  <tr>
					  <td class="textoSeccion"><% =GF_TRADUCIR("Obtenga diferentes tipos de reportes.") %></td>
				  </tr>
				</tbody>		
				 <tr><td>&nbsp;</td></tr>
                
			</table>
		</td></tr>
		</table>			
		</td>
	</tr>
</table>
</body>
</html>
