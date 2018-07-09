<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
dim strSQL, rs, oConn 
strSQL = "SELECT VC.IDVALE FROM TBLVALESCABECERA VC INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE WHERE VD.VLUPESOS IS NULL GROUP BY VC.IDVALE"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		Response.Write "<br>GRABANDO VALE(" & rs("IDVALE") & ")"
		call grabarPreciosVigentesPorArticulo(rs("IDVALE")) 
	rs.movenext
wend		
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
%>