<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Compras</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript">	
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		tb.addButton("Home-16x16.png", "Home", "cerrar()");
		tb.draw();	
	}
	function cerrar(){
		location.href='comprasIndex.asp'
	}
</script>
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
</head>

<body onLoad="bodyOnLoad()">
<% call GF_TITULO2("kogge64.gif","Reportes de Compras") %>
<div id="toolbar"></div>
<br />
<table align="center" width="80%" height="100%">
	<tr valign="top">
    	<td></td>
        <td>
        	<table align="center" width="80%">			
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteArticulosPedidosNoRecibidos.asp?origen=COMPRAS';">
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
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasReporteArticulos.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Compras por Articulos ") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Lista todas las compras realizadas de un determinado articulo en un periodo dado.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasReportePICSaldoPendiente.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("PIC con Saldo pendiente") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Lista todos los PIC con saldo pendiente en un periodo dado.") %></td>
					</tr>
				</tbody>
				<tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='almacenReporteVariacionPrecio.asp?origen=COMPRAS';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Variacion de Precios") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Lista las variaciones de precios de los articulos.") %></td>
					</tr>
				</tbody>
<!--Primero-->
                <!--
                <tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasReportePagoDupdoAProveedores.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Pagos duplicados a proveedores.") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Identificar pagos de facturas realizados en forma duplicada a un mismo proveedor, ya sea a traves de una o de varias Sociedades. ") %></td>
					</tr>
				</tbody>
                -->
<!--Segundo-->       
                <!--        
                <tr><td>&nbsp;</td></tr>
				<tbody onMouseOver="this.style.color= 'GoldenRod'" onMouseOut="this.style.color= 'black'" style="cursor: pointer" onClick="document.location.href='comprasReporteFacturaDupdoDeProveedores.asp';">
					<tr>
						<td rowspan="2" width="10%"><img align="absMiddle" src="images/compras/RPT-48x48.png"></td>
						<td rowspan="2" width="1%">&nbsp;</td>						
						<td class="title_sec_section"><% =GF_TRADUCIR("Facturas duplicadas de proveedores") %></td>
					</tr>
					<tr>
						<td valign="top" class="textoSeccion"><% =GF_TRADUCIR("Detectar los casos de registro de facturas de proveedores duplicadas, donde mediante su numero de factura fisico se presume su registro en mas de una oportunidad. Este an&aacute;lisis puede complementarse con informacion del pago de dichas facturas.") %></td>
					</tr>
				</tbody>
                -->
            </table>
        </td>
    </tr>
</table>
</body>
</html>
