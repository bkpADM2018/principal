<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<% Dim cdPedido
cdPedido = GF_PARAMETROS7("cdPedido","",6) %>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<style type="text/css">p{font: normal normal 18px Times;text-align: center}</style>
</head>
<body>
	<p><% =GF_TRADUCIR("El codigo asignado al pedido es") %>: </p>
	<p><span style="color:blue;"><% =cdPedido %></span></p>
</body>
</html>