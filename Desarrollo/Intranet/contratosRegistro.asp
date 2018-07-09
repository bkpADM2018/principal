<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
dim producto, sucursal, operacion, numero, cosacha
dim rs, oConn, strSQL, registro
producto= GF_Parametros7("producto",0,6)
sucursal= GF_Parametros7("sucursal",0,6)
operacion= GF_Parametros7("operacion",0,6)
numero= GF_Parametros7("numero",0,6)
cosecha= GF_Parametros7("cosecha",0,6)


strSQL = "SELECT * FROM TOEPFERDB.TBLCONTRATOSCONF WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha 
Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
'Response.Write strSQL 
if not rs.eof then
	registro = rs("REGISTRO")
	registro = replace(registro,vbCrLf,"<br>")
	registro = replace(registro,"<hr>","<br><b>-----------------------------------------Fin mensaje----------------------------------------</b><br><br>")
else
	registro = GF_Traducir("No se encontro registro para este contrato")
end if
%>
<html>
<head>
  <title>TOEPFER INTERNATIONAL - <%GF_Traducir("Contratos")%></title>
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>  
<script>
	function imprimir(obj){
		obj.style.visibility="hidden";
		window.print()	
	}
	function ver(){
		document.getElementById("linkPrint").style.visibility = "visible";
	}
</script>
<body onmouseover="ver()">
<table>
	<tr>
		<td>
			<code><%=registro%></code>
		</td>
	</tr>
	<tr>
		<td align="center">
			<a id="linkPrint" style="cursor:pointer;" onclick="imprimir(this);"><font color="blue">[Imprimir]</font></a>
		</td>
	</tr>
</table>
</body>
</html>
