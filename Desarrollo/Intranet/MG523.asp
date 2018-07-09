<!--#include file="Includes/procedimientosMG.asp"--> 
<!--#include file="Includes/procedimientosUnificador.asp"--> 
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"--> 

<%
' Seleccionar rol o procedimiento
Dim con,sql,rs
dim asa_link, P_km,P_KC,P_DS,P_KR, My_km,My_KC,My_DS,My_KR, My_Public, My_SrKr, My_msg
dim my_LinkLista, my_Link
P_kr = request.querystring("p_kr") + 0
call GP_ConfigurarMomentos()
IF NOT gf_mgkr( p_kr,p_km,p_kc,p_ds ) then
   my_msg = "No se halló el procedimiento a ejecutar"
END IF  
response.write "vamos" 
MY_Link = GF_DT1( "READ","SPLNPT","","",P_KM, P_KC)
response.write "venimos"
if my_link = "?" then
%>
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
</head>
<body>
  <table align=center width=30% border=2 bordercolor="blue">
  <br><br><br><br><br><br>
    <tr>
	   <td colspan=2 align=center bgcolor="#CCCC99"><font size=4 color="red"><b>Path inexistente</b></font></td>
    </tr>
	<tr>   
	   <td align = "center" bgcolor="white">[<A href="javascript:history.back()"><% =GF_TRADUCIR("Volver") %></a>]</td>
	   <td align = "center" bgcolor="white">[<A href="MG210.asp?P_KR=<%=p_kr%>"><% =GF_TRADUCIR("Establecer Path") %></a>]</td>
	</tr> 
  </table>	
</body>
</html>
<%
else  
'Response.Write MY_Link
Response.Redirect(MY_Link)
end if
%>