<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosObras.asp"-->		
<!--#include file="Includes/procedimientosSql.asp"-->		
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->		
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Dim cdMoneda
cdMoneda = MONEDA_DOLAR
%>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
	<link rel="stylesheet" href="css/main.css" type="text/css">		
	<script type="text/javascript" src="scripts/channel.js"></script>	
</head>
<body>
	<table width="100%">
		<tr>
			<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/AFE_folder-32x32.png"> <% =GF_TRADUCIR("Administracion de AFE") %></td>
		</tr>
		<tr>
			<td></td>
			<td>			
				<table width="100%" border="0" cellpadding="1" cellspacing="2">				
					<tr><td>&nbsp;</td></tr>
					<tr><td>					
						<!--#include file="comprasListaAFE.asp"-->
					</td></tr>
				</table>
			</td>
		</tr>			
	</table>
</body>
</html>