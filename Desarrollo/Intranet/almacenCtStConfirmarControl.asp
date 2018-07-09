<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<% Dim idControl
idcontrol = GF_PARAMETROS7("idcontrol",0,6) %>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<style type="text/css">p{font: normal normal 12px Times;text-align: center}</style>
	<script type="text/javascript">		
		function CerrarPopUp(){
			parent.CerrarVentana();	
			
		}
	</script>
</head>
<body>
	<table width='100%' align='center' id='tbLoading' border=0 class='.pp'>
		<tr><td align='center' >
		<h3><% =GF_TRADUCIR("El Id de Control asignado es") %></h3></td></tr>
		<tr><td align='center' >
		<h3><span style='color:blue;'><%=idControl %></span></h3></td></tr>	
		<tr><td align="center"><input type="button" id="acpetar" onclick="CerrarPopUp()" value="Aceptar"></td></tr>
	</table>
</body>
</html>
