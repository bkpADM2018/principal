<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
dim idDivision, idAlmacenes, fechaCierre, strSQL, rs, con
flagHeader = false
idAlmacenes = GF_Parametros7("idAlmacen", "", 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)

'Obtener la lista de los VRS que aun no esten valorizados contablemente y valorizarlos
strSQL = "SELECT DISTINCT(T.IDVALE) FROM (" & _
		 "	SELECT VC.IDVALE, ABS(EXISTENCIA)-ABS(TOTAL_VALUADO) AS SALDO FROM TBLVALESCABECERA VC " & _
		 "  INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _
		 "	LEFT JOIN " & _
		 "		(SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS TOTAL_VALUADO FROM TBLVALESCONTABLE GROUP BY IDVALE, IDARTICULO) AS VAL " & _
		 "	ON VD.IDVALE=VAL.IDVALE AND VD.IDARTICULO=VAL.IDARTICULO " & _
		 "	WHERE ESTADO=" & ESTADO_ACTIVO & " AND VC.CDVALE='" & CODIGO_VS_RECLASIFICACION_STOCK & "' AND EXISTENCIA<>0 " & _
		 "  AND VC.IDALMACEN IN (" & idAlmacenes & ") AND VC.FECHA BETWEEN '" & FECHA_INICIO_CONTABLE & "' AND '" & fechaCierre & "'" & _
		 "  )T " & _
		 "	WHERE T.SALDO <> 0 OR T.SALDO IS NULL "
Response.Write "<hr>Seleccion de vales a valuar " & strSQL & "<hr>"	
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		call valorizarValeContable(rs("IDVALE"),fechaCierre)
	rs.movenext
wend	

'Calcular nuevo precio a partir de VRS valuados en este cierre
strSQL = "SELECT AL.IDDIVISION, CO.IDARTICULO, SUM(CO.CANTIDAD) AS TOTAL_UNIDADES, SUM(CO.CANTIDAD*CO.VLUPESOS) AS TOTAL_IMPORTE, SUM(CO.CANTIDAD*CO.VLUDOLARES) AS TOTAL_IMPORTE_D " & _
		 "    FROM TBLVALESCABECERA VC " & _ 
		 "        INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE AND VD.EXISTENCIA > 0 " & _
		 "        INNER JOIN TBLVALESCONTABLE CO ON VC.IDVALE=CO.IDVALE AND CO.IDARTICULO=VD.IDARTICULO " & _
		 "        INNER JOIN TBLALMACENES AL " & _ 
		 "        ON VC.IDALMACEN=AL.IDALMACEN " & _ 
		 "    WHERE FECHACIERRE=" & fechaCierre & " AND VC.IDALMACEN IN (" & idAlmacenes & ") AND VC.CDVALE='" & CODIGO_VS_RECLASIFICACION_STOCK & "'" & _
		 "    GROUP BY CO.IDARTICULO, IDDIVISION " 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Response.Write "<hr>Seleccion de vales valuados " & strSQL & "<hr>"	
while not rs.eof 
	call acumularArticulo(rs("IDDIVISION"), rs("IDARTICULO"), rs("TOTAL_UNIDADES"), rs("TOTAL_IMPORTE"), rs("TOTAL_IMPORTE_D"))
	rs.movenext
wend

Response.Write "Hecho..."
'---------------------------------------------------------------------------------------------------------------
%>