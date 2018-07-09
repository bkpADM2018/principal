<!--#include file="../ActiSAIntra/Includes/procedimientosMG.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"--> 
<!--#include file="../ActiSAIntra/Includes/procedimientosfechas.asp"-->

<%
   ProcedimientoControl "FIESTA"
   
   Dim P_PIC,P_ID
   
   P_PIC=GF_PARAMETROS("P_PIC","")
   P_ID=GF_PARAMETROS("P_ID","")
%>   
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<meta name="Microsoft Border" content="tl, default">
</head>

<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p>&nbsp;</p>
</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td valign="top" width="1%">

<p>&nbsp;</p>
</td><td valign="top" width="24"></td><!--msnavigation--><td valign="top">
<% =GF_TITULO("FiestaLogo.gif","Fiesta de Fin de Año de ActiSA") %>
<p align="center"><img src="images/Fiesta<% =P_ID %>-<% =P_PIC %>.jpg"> </p>
<!--msnavigation--></td></tr><!--msnavigation--></table></body>
</html>
