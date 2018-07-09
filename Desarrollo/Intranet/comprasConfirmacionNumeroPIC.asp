<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
Dim CTZN_idCotizacion, CTZN_idPedido, CTZN_FechaBudget
Dim CTZN_idObra
'-----------------------------------------------------------------------------------------------
CTZN_idPedido = GF_PARAMETROS7("idPedido",0,6)
CTZN_idCotizacion = GF_PARAMETROS7("idCotizacion",0,6)
CTZN_idObra = GF_PARAMETROS7("idObra",0,6)
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<style type="text/css">
p {
	font: normal normal 18px Times;	
	text-align: center
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

</script>
</head>
<body>
	<p><% =GF_TRADUCIR("El codigo asignado de referencia es ") %>: </p><p><span style="color:blue;">REF <% =CTZN_idCotizacion %></span></p>
	<div id="avisoNDA" align="center" class="TDBAJAS"></div>
	<table align=center>
		<tr>
			<td>
				<a href="comprasPICPrint.asp?idCotizacionElegida=<%=CTZN_idCotizacion%>" target=_blank> <img src="images/compras/printer-16x16.png"></a>
			</td>
			<td>
				<a href="comprasPICPrint.asp?idCotizacionElegida=<%=CTZN_idCotizacion%>" target=_blank> <%=GF_TRADUCIR("Imprimir Pedido")%></a>
			</td>
		</tr>
	</table>
</body>
</html>