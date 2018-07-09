<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<%
	Dim idpedido
	idPedido     = GF_PARAMETROS7("idPedido",0,6)	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Mensaje</title>
<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
	function crearAfe(){
		parent.location.href = "comprasAFE.asp?idPedido=<% =idPedido %>";
	}
</script>
</head>

<body>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" class="reg_header">
  <tr>
    <td class="reg_header_nav round_border_top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center"><img src="images/compras/warning-16x16.png" width="16" height="16"></td>
        <td align="center" class="reg_header_nav"><%=GF_TRADUCIR("Se requiere la confección de un AFE")%></td>
      </tr>
    </table></td>
  </tr>
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td align="center" valign="middle"><table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="right"><%=GF_TRADUCIR("Para realizarlo haga click aqui")%>: </td>
        <td align="center"><img src="images/compras/AFE-32X32.png" width="35" height="35" border="0" onClick="crearAfe()" style="cursor:pointer" title="<%=GF_TRADUCIR("Crear Afe")%>"></td>
      </tr>
    </table></td>
  </tr>
  <tr><td>&nbsp;</td></tr>
</table>
</body>
</html>
