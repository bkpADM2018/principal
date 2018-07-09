<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
dim idDivision, idAlmacenes, fechaCierre, strSQL, rs, con
flagHeader = false
idAlmacenes = GF_Parametros7("idAlmacen", "", 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
'Obtener la lista de los VAles a los cuales hay que valorizar (VMS,AJU,AJT,AJS)
strSQL = "SELECT T.IDVALE FROM (" & _
		 "	SELECT VC.IDVALE, EXISTENCIA-TOTAL_VALUADO AS SALDO, FECHA FROM TBLVALESCABECERA VC " & _
		 "  INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _
		 "	LEFT JOIN " & _
		 "		(SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS TOTAL_VALUADO FROM TBLVALESCONTABLE GROUP BY IDVALE, IDARTICULO) AS VAL " & _
		 "	ON VD.IDVALE=VAL.IDVALE AND VD.IDARTICULO=VAL.IDARTICULO " & _
		 "	WHERE VC.IDALMACEN IN (" & idAlmacenes & ") AND VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.CDVALE IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "','" & CODIGO_VS_AJUSTE_STOCK & "') AND EXISTENCIA<>0 " & _
		 "  AND VC.FECHA BETWEEN '" & FECHA_INICIO_CONTABLE & "' AND '" & fechaCierre & "'" & _
		 "  )T " & _
		 "	WHERE T.SALDO <> 0 OR T.SALDO IS NULL GROUP BY T.IDVALE, T.FECHA ORDER BY FECHA ASC"
'Response.Write "<hr>Seleccion de vales a valuar " & strSQL & "<hr>"	
'Response.End 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

while not rs.eof
		call valorizarValeContable(rs("IDVALE"),fechaCierre)
	rs.movenext
wend

Response.Write "Hecho..."
%>