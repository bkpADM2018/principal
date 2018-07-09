<!--#include file="../ActiSAIntra/Includes/procedimientosMG.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"--> 
<!--#include file="../ActiSAIntra/Includes/procedimientosfechas.asp"-->

<%
'****************************************************************************************************************************
sub levantarParametros(p_prefijo, p_ID, p_num, p_ext, p_total)
	p_ID  = GF_Parametros7("p_id","",6)
	p_num = GF_Parametros7("p_num","",6)
	p_ext = GF_Parametros7("p_ext","",6)
	p_total = GF_Parametros7("p_total","",6)
	p_prefijo = GF_Parametros7("p_prefijo","",6)
end sub
'****************************************************************************************************************************
dim id, num, ext, totalFotos, prefijo

call levantarParametros(prefijo, id, num, ext, totalFotos)

%>
<html>
<head>
	<title></title>
	<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
	<script language="javascript">
                function corregir() {
                         var i = document.getElementById("imagenActual");
                         if (i.width > i.height) {
                            i.width = 640;
                            i.height = 480;
                         } else {
                            i.width = 480;
                            i.height = 640;
                         }
                }
                
	</script>
</head>

<body onload="javascript:corregir();">
<table align="center" border="0">
	<tr>
		<td>
			<table align="center" border="0">
				<tr>
					<td align="right" width="30">
						<%if (cint(num) <> 1) then%>
							<a href="fotosdetalle.asp?p_prefijo=<%=prefijo%>&p_id=<%=id%>&p_num=<%=num-1%>&p_ext=<%=ext%>&p_total=<%=totalFotos%>" title="<%=GF_Traducir("Anterior")%>"><img src="images/anterior.gif"></a>
						<%else%>
							&nbsp;
						<%end if%>
					</td>
					<td align="center" width="40">
						<b><%=num%>/<%=totalFotos%></b>
					</td>
					<td align="left" width="30">
						<%if (cint(num) < cint(totalFotos)) then%>
							<a href="fotosdetalle.asp?p_prefijo=<%=prefijo%>&p_id=<%=id%>&p_num=<%=num+1%>&p_ext=<%=ext%>&p_total=<%=totalFotos%>" title="<%=GF_Traducir("Siguiente")%>"><img src="images/siguiente.gif"></a>
						<%end if%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center">
			<img id="imagenActual" src="imagesFiestas/<%=prefijo%><%=id%>-<%=num%>.<%=ext%>">
		</td>
	</tr>
</table>
</body>
</html>
