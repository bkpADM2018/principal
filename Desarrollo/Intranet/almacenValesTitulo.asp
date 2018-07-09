<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
dim myCdVale, txtLYND
myCdVale = GF_PARAMETROS7("cdVale","",6)
if myCdVale = "" then 
	txtLYND = "Edición"
else
	txtLYND = getLeyendaCdVale(myCdVale)
end if
%>
<html>
<head>
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<body>
<%
	call GF_TITULO2("kogge64.gif","Almacen - " & txtLYND)
	Select Case myCdVale
		case CODIGO_VS_AJUSTE_VALE:
			server.execute( "almacenValesAJU.asp" )
		case CODIGO_VS_AJUSTE_PEDIDO:
			server.execute( "almacenValesAJP.asp" )
		case CODIGO_VS_AJUSTE_STOCK:
			server.execute( "almacenValesAJS.asp" )
		case CODIGO_VS_AJUSTE_TRANSFERENCIA:
			server.execute( "almacenValesAJT.asp" )
		case else
			server.execute( "almacenVales.asp" )
	End Select
%>
</body>