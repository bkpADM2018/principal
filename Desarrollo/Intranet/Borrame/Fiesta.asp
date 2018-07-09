<!--#include file="../ActiSAIntra/Includes/procedimientosMG.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"--> 
<!--#include file="../ActiSAIntra/Includes/procedimientosfechas.asp"-->
<%
   'ProcedimientoControl "FIESTA"   
   Dim P_intCantidad,P_ID,P_strLugar,i,j
   Dim P_strExtension, isSmall, P_SMALL, numero, P_strPrefijo

   p_titulo=GF_PARAMETROS("P_TITULO","")
   if (p_titulo = "") then p_titulo = "Fiesta de Fin de Año de ACTI"
   P_intCantidad=CInt(GF_PARAMETROS("P_CANT",""))
   P_ID=GF_PARAMETROS("P_ID","")
   P_strLugar=GF_PARAMETROS("P_LUGAR","")
   P_strExtension = GF_PARAMETROS7("P_EXT","",6)
   if (P_strExtension = "") then P_strExtension="jpg" 		
   isSmall = ""
   P_SMALL = GF_PARAMETROS7("P_SMALL","",6)   
   if (P_SMALL = "") then isSmall = "_small"
   P_strPrefijo = GF_PARAMETROS7("P_PREFIJO","",6)
   if (P_strPrefijo = "") then P_strPrefijo = "Fiesta"   

    
%>   
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<meta name="Microsoft Border" content="tl, default">
</head>
<body><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p>&nbsp;</p>
</td></tr></table><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td valign="top" width="1%">

<p>&nbsp;</p>
</td><td valign="top" width="24"></td><!--msnavigation--><td valign="top">
<% =GF_TITULO("FiestaLogo.gif",p_titulo) %>
<font face="Brush Script MT" size=4><div align="center"><% =P_strLugar%></div></font>

<% for i= 1 to P_intCantidad step 4%>
      <p align="center">
   <% for j= 0 to 3 	
         if ((i+j) <= P_intCantidad) then%>
         	<span style="cursor:hand;" onClick="window.open('fotosdetalle.asp?p_prefijo=<%=P_strPrefijo%>&p_id=<%=p_id%>&p_num=<%=i+j%>&p_ext=<%=p_strExtension%>&p_total=<%=p_intCantidad%>','_blank','height=550,width=700,scrollbars=yes');">
			<img src="imagesFiestas/<% =P_strPrefijo %><% =P_ID %>-<% =i+j %><% =isSmall %>.<% =P_strExtension %>" border="2">
		</span>&nbsp;&nbsp;&nbsp;&nbsp; 
   <%    end if
      Next %>
      </p>
<% next %>   
</td></tr></table></body>
</html>
