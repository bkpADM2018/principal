<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Dim idPedido, resp, conn, strSQL, rs, comentario, my_dsSolicitante
idPedido = GF_PARAMETROS7("idPedido",0,6)
call initHeader(idPedido)
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
</head>
<body>	
	<table width="100%">				
		<tr>
			<td><font class="big"><% =GF_TRADUCIR("Comentarios al pedido:") %> <b><% =pct_cdPedido %></b></font></td>
		</tr>
	</table>		
	<hr>
	<% 
	strSQL="Select * from TBLPCTCOMENTARIOS where IDPEDIDO=" & idPedido
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	while (not rs.eof)	
	%>	
	<table class="reg_header" align="center" width="100%" border="1" rules="none">				
		<tr>
			<td valign="top" rowspan="2">
				<img src="images/compras/PCT_comments-32x32.png">
			</td>
			<td>
				<% my_dsSolicitante = getUserDescription(rs("CDUSUARIO"))%>
				<% =my_dsSolicitante & GF_TRADUCIR(" el dia ") & GF_FN2DTE(rs("MOMENTO")) & GF_TRADUCIR(" comento:") %>
			</td>
		</tr>				
		<tr>
			<td>
				<% =GF_TRADUCIR(rs("COMENTARIO")) %>
				<br>
			</td>
		</tr>				
	</table>		
	<br>	
	<%
		rs.MoveNext()
	wend 
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)		
	%>

</body>
</html>