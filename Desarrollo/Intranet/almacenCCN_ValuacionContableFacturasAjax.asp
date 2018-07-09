<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
dim idDivision, fechaCierre, strSQL, rs, con
flagHeader = false
idDivision = GF_Parametros7("idDivision", 0, 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
'Obtener todas las facturas del mes a fin de registrar la mercaderia contable y recalcular precios del mes
'Se deben excluir todo lo que sea inversion ya estos no registran un aumento en la existencia del inventario, se activa por la obra
'SE ASUME QUE TODOS LOS PICS QUE TENGAN PARTIDA PRESUPUESTARIA NO VAN A INVENTARIO
strSQL = "SELECT IDARTICULO, SUM(CANTIDAD) AS TOTAL_UNIDADES, SUM(IMPORTEP) AS TOTAL_IMPORTE, SUM(IMPORTED) AS TOTAL_IMPORTE_D FROM " & _
		 " ( " & _
		 "	SELECT DF.IDArticulo AS IDARTICULO, CANTIDAD AS CANTIDAD, ImportePesos AS IMPORTEP, ImporteDolares AS IMPORTED, ESINVERSION " & _
		 "    FROM MEP001C DF  " & _ 
		 "        INNER JOIN [Database].[dbo].[MEP001A] CF ON CF.NROINT = DF.NROINT AND CF.ANULADO<>'S' " & _
		 "		  INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=DF.IDARTICULO " & _
		 "		  INNER JOIN TBLARTCATEGORIAS CAT ON CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
		 "	      LEFT JOIN TBLDATOSOBRAS DOB ON DOB.IDOBRA=DF.IDOBRA " & _
		 "         WHERE CF.[date] BETWEEN '" & GF_FN2DTCONTABLE(LEFT(fechaCierre,6) & "01") & "' AND '" & GF_FN2DTCONTABLE(fechaCierre) & "' " & _
		 "             AND CF.codcia=" & getCIADivision(idDivision) & _
		 "             AND ART.BIENUSO='N' " & _
		 "			   AND TIPCBT IN (3,2,1)  AND CF.ESTADO IN ('P','','E') " & _
		 "			   AND CAT.TIPOCATEGORIA='" & TIPO_CAT_BIENES & "' " & _
		 " ) TG " & _
		 " WHERE TG.ESINVERSION IS NULL " & _
         "GROUP BY IDARTICULO"

'Ver que formularios entran
'"			   AND ((TIPCBT = 1 AND LETRA = 'A') OR (TIPCBT IN (3,2))) AND CF.ESTADO IN ('P','','E') " & _

Response.Write "<BR>" & strSQL 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof 
	if cdbl(rs("TOTAL_UNIDADES")) <> 0 then call acumularArticulo(idDivision, rs("IDARTICULO"), rs("TOTAL_UNIDADES"), Cdbl(rs("TOTAL_IMPORTE"))*10000, Cdbl(rs("TOTAL_IMPORTE_D"))*10000)
	rs.movenext
wend

if len(myName) = 0 then	
	Response.Write "Hecho..."
else
	Response.Write "ERROR-" & myName
end if	
'Response.Write "Hecho..."
%>