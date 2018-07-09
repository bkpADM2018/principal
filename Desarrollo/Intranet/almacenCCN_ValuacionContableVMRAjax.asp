<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<%
dim idDivision, idAlmacenes, fechaCierre, strSQL, rs, con
flagHeader = false
idAlmacenes = GF_Parametros7("idAlmacen", "", 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)

'Obtener la lista de los VMR que aun no esten valorizados contablemente y valorizarlos
strSQL = "SELECT T.IDVALE FROM (" & _
		 "	SELECT VC.IDVALE, EXISTENCIA-TOTAL_VALUADO AS SALDO, FECHA FROM TBLVALESCABECERA VC " & _
		 "  INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _
		 "  INNER JOIN TBLPMCABECERA ALM_DEST ON ALM_DEST.IDPEDIDO=VC.PARTIDAPENDIENTE " & _ 
         "  INNER JOIN TBLALMACENES ALM1 ON ALM_DEST.IDALMACEN=ALM1.IDALMACEN " & _ 
		 "  INNER JOIN TBLALMACENES ALM2 ON ALM_DEST.IDALMACENDEST=ALM2.IDALMACEN " & _ 
		 "	LEFT JOIN " & _
		 "		(SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS TOTAL_VALUADO FROM TBLVALESCONTABLE GROUP BY IDVALE, IDARTICULO) AS VAL " & _
		 "	ON VD.IDVALE=VAL.IDVALE AND VD.IDARTICULO=VAL.IDARTICULO " & _
		 "	WHERE VC.ESTADO=" & ESTADO_ACTIVO & " AND VC.CDVALE='" & CODIGO_VS_RECEPCION & "' AND ALM1.IDDIVISION<>ALM2.IDDIVISION AND EXISTENCIA<>0 " & _
		 "  AND VC.FECHA BETWEEN '" & FECHA_INICIO_CONTABLE & "' AND '" & fechaCierre & "'" & _		 
		 "  )T " & _
		 "	WHERE T.SALDO <> 0 OR T.SALDO IS NULL GROUP BY T.IDVALE, T.FECHA ORDER BY FECHA ASC"
Response.Write "<hr>Seleccion de VMR a valuar " & strSQL & "<hr>"	
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		call valorizarValeContable(rs("IDVALE"),fechaCierre)
	rs.movenext
wend	

'Calcular nuevo precio a partir de VMR valuados en este cierre
strSQL = "SELECT AL.IDDIVISION, IDARTICULO, SUM(CANTIDAD) AS TOTAL_UNIDADES, SUM(CANTIDAD*VLUPESOS) AS TOTAL_IMPORTE, SUM(CANTIDAD*VLUDOLARES) AS TOTAL_IMPORTE_D " & _
		 " FROM TBLVALESCABECERA VC INNER JOIN TBLVALESCONTABLE CO ON VC.IDVALE=CO.IDVALE AND CANTIDAD>0 " & _
		 " INNER JOIN TBLALMACENES AL ON VC.IDALMACEN=AL.IDALMACEN " & _
		 " WHERE FECHACIERRE=" & fechaCierre & " AND VC.CDVALE='" & CODIGO_VS_RECEPCION & "'" & _
		 " GROUP BY IDARTICULO, IDDIVISION "
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Response.Write "<hr>Seleccion de vales valuados " & strSQL & "<hr>"	
while not rs.eof 
	call acumularArticulo(rs("IDDIVISION"), rs("IDARTICULO"), rs("TOTAL_UNIDADES"), rs("TOTAL_IMPORTE"), rs("TOTAL_IMPORTE_D"))
	rs.movenext
wend
Response.Write "Hecho..."
%>