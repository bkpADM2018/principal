<!--#include file="../ActiSAIntra/Includes/procedimientosMG.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"-->
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table width="100%" heigth="100%">
  <tr valign="center">
    <td height="449"> 
   <table align="center" bordercolor="#00AACA" bgcolor="#FFFFFF" width="60%" heigth="60%" border="4">
  <tr align="center"> 
    <td> <img src="images/index_logo.gif"> 
      <hr color="#00AACA">
	  <font color="#447755" size="4" face=Times>
	  <lu> 
	  <li style="list-style-type:disc;"></li>
      <li style="list-style-type:none;"><% =GF_TRADUCIR("Muchas gracias por registrarse en nuestra empresa") %>.</li>
	  <li style="list-style-type:none;"></li>
      <li style="list-style-type:none;"><% =GF_TRADUCIR("A la brevedad recibira la confirmacion de su ingreso en nuestro sistema") %>.</li>
      <li style="list-style-type:disc;"></li>
	  </lu>
	  </font>
	</td>
  </tr>
  <tr><td align="center"><a href="<% =session("Home") %>" target="_parent"><img src="images/Anterior.gif" alt="<% =GF_TRADUCIR("Volver") %> " border="0"></a></td></tr>
</table>
</td></tr></table>
</body>
</html>
