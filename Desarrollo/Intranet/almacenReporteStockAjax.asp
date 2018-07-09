<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<%
Dim idAlmacen, fechaBusqueda, strSQL, rs, con
Dim rsArt, categoria, incluir

idAlmacen = GF_Parametros7("almacen", 0, 6)
fechaBusqueda = GF_Parametros7("fechaBusqueda", "", 6)
incluir = GF_Parametros7("incluir", "", 6)
categoria	= GF_PARAMETROS7("categoria", 0 , 6)
fechaBusqueda = GF_DTE2FN(fechaBusqueda)
if incluir then
	fechaBusqueda = GF_DTEADD(fechaBusqueda,1,"d")
end if	
'**************************************************************
'		COMIENZO DE LA PAGINA
'**************************************************************
'Inicializa el workfile.
Call dropArticulosTemp()
'Completa el workfile.
Call fillWorkFile()
'Completa datos de stock.
Set rsArt = CalcularStock()
while not rsArt.eof
	call insertStockInfo(rsArt("IDARTICULO"), rsArt("TOTALEXISTENCIA"), rsArt("TOTALSOBRANTE"))
	rsArt.movenext
wend
'--------------------------------------------------------------------------------------
'Se obtiene la fecha del último cierre contable.
Function getFechaCierreAnt()
	dim fecha, FEC_strSQL, FEC_conn, FEC_rs

	'Se obtiene la fecha de ultimo cierre, recordando que los cierres son al final del dia que indican 
	'y yo necesito los datos al inicio de la fecha solicitada, 
	'por eso se debe tomar un cierre estrictamente anterior a la fecha de busqueda.
	FEC_strSQL = " SELECT MAX(FECHACIERRE) FECHACIERRE FROM TBLCIERRESARTICULOS2 "
	FEC_strSQL = FEC_strSQL & " WHERE idalmacen = " & idAlmacen	
	FEC_strSQL = FEC_strSQL & " AND fechacierre < " & fechaBusqueda
	call executeQueryDb(DBSITE_SQL_INTRA, FEC_rs, "OPEN", FEC_strSQL)	
	if ((not FEC_rs.eof) and (not isNull(FEC_rs("FECHACIERRE")))) then
		fecha = FEC_rs("FECHACIERRE")
	else
		fecha = 0
	end if
	call executeQueryDb(DBSITE_SQL_INTRA, FEC_rs, "CLOSE", FEC_strSQL)	
	getFechaCierreAnt = fecha
End Function
'--------------------------------------------------------------------------------------
'Se cargan los datos de los artículos y su valuación en el workfile.
Function fillWorkFile()
	Dim mywhere, auxrs, auxconn, auxstrSQL, INS_strSQL, INS_rs, INS_conn, idDivision

	if (categoria <> -1) then Call mkWhere(mywhere, "ART.IDCATEGORIA", categoria,"=", 1)	
	Call mkWhere(mywhere, "ART.ESTADO", ESTADO_ACTIVO,"=", 1)
	idDivision = getDivisionAlmacen(idAlmacen)

	auxstrSQL = "SELECT " & fechaBusqueda & " FECHACIERRE, " & idAlmacen & ", ART.IDARTICULO, "
	auxstrSQL = auxstrSQL & " 0, 0, vlu.VLUPESOS, vlu.VLUDOLARES, "
	auxstrSQL = auxstrSQL & " '" & session("Usuario") & "' CDUSUARIO, "
	auxstrSQL = auxstrSQL & session("MmtoSistema") & " MOMENTO, 0 STOCKCONTABLE, vlu.MMTOULTIMACOMPRA, vlu.VLUPESOSULTIMACOMPRA , vlu.VLUDOLARESULTIMACOMPRA, vlu.IDPIC "
	auxstrSQL = auxstrSQL & " FROM TBLARTICULOS ART "
	auxstrSQL = auxstrSQL & " LEFT JOIN TBLARTICULOSPRECIOS vlu "
	auxstrSQL = auxstrSQL & " ON       vlu.idarticulo  = art.idarticulo "
	auxstrSQL = auxstrSQL & " AND      vlu.IDDIVISION   = " & idDivision
	auxstrSQL = auxstrSQL & " AND      vlu.MMTOPRECIO = "
									'Subconsulta que permite obtener el momento del último precio registrado antes de la fecha de consulta.
	auxstrSQL = auxstrSQL & "          (SELECT MAX(MMTOPRECIO) "
	auxstrSQL = auxstrSQL & "          	FROM TBLARTICULOSPRECIOS "
	auxstrSQL = auxstrSQL & "          	WHERE IDDIVISION = " & idDivision
	auxstrSQL = auxstrSQL & "          	AND idarticulo = vlu.idarticulo "
										'Para la busqueda se debe poner un momento y no solo una fecha.
										'Se sigue el lineamiento elegido de tomar los valores al inicio de la fecha seleccionada.
	auxstrSQL = auxstrSQL & "			AND mmtoprecio <= " & fechaBusqueda & "000000"	
	auxstrSQL = auxstrSQL & "           )"
	auxstrSQL = auxstrSQL & mywhere
	auxstrSQL = auxstrSQL & " ORDER BY ART.IDARTICULO "
	call executeQueryDb(DBSITE_SQL_INTRA, auxrs, "OPEN", auxstrSQL)
	if (not auxrs.eof) then
		INS_strSQL = "INSERT INTO TBLREPORTESTOCKWF "
		INS_strSQL = INS_strSQL & auxstrSQL
		'Response.Write "<HR>INS:" & INS_strSQL 
		call executeQueryDb(DBSITE_SQL_INTRA, INS_rs, "EXEC", INS_strSQL)
	end if	
End Function
'--------------------------------------------------------------------------------------
'Ontiene el stock de los artículos a la fecha solicitada. OJO! Obtiene el stock de todos los artículos y no solo de los que interesan, esto habría que mejorarlo.
Function CalcularStock()
	Dim fechaCierreAnterior, strSQL, strSQLArt

	fechaCierreAnterior = getFechaCierreAnt()

	strSQLArt = "Select IDARTICULO from TBLREPORTESTOCKWF where CDUSUARIO= '" & session("Usuario") & "' and IDALMACEN=" & idAlmacen

	'Tener en cuenta que los stocks se calculan al inicio de la fecha solicitada (fechaBusqueda).
	'Por esto, se toma como stock inicial el indicado en el cierre ultimo anterior a esta fecha y los movimientos 
	'desde el dia siguiente hasta el dia anterior al indicado.
	'(Las constantes en las Sub-SQLs son para evitar que el UNION elimine posible registros duplicados)
	strSQL ="SELECT IDARTICULO, SUM(TOTALEXISTENCIA) AS TOTALEXISTENCIA, SUM(TOTALSOBRANTE) AS TOTALSOBRANTE FROM " & _
			"( " 
		'Se totalizan los remitos (ingresos).
	strSQL = strSQL & "SELECT 1 as NUMERO, IDARTICULO, SUM(EXISTENCIA) AS TOTALEXISTENCIA, SUM(SOBRANTE) AS TOTALSOBRANTE FROM " & _
			"	TBLREMCABECERA RC INNER JOIN TBLREMDETALLE RD " & _
			"	    ON RC.IDREMITO = RD.IDREMITO " & _
			"	        WHERE RC.IDALMACEN=" & idAlmacen & " AND FECHA > " & fechaCierreAnterior & " AND FECHA < " & fechaBusqueda & _
			"	        AND IDARTICULO IN (" & strSQLArt & ")" & _
			"	            GROUP BY IDARTICULO " & _
			" UNION "
			'Se totalizan los egresos por vales.
			'Debido a que en el vale de reclasificacion de stock figuran tanto los valores que egresan como los que ingresan debo separarlo para agregar la condicion Mayor/Menor 
	strSQL = strSQL & "SELECT 2 as NUMERO, IDARTICULO, SUM(-ABS(EXISTENCIA)) AS TOTALEXISTENCIA, SUM(-ABS(SOBRANTE)) AS TOTALSOBRANTE FROM  " & _
			"    TBLVALESCABECERA VC INNER JOIN TBLVALESDETALLE VD " & _
			"        ON VC.IDVALE = VD.IDVALE " & _
			"            WHERE VC.IDALMACEN=" & idAlmacen & " AND VC.FECHA > " & fechaCierreAnterior & " AND VC.FECHA < " & fechaBusqueda & " AND VC.ESTADO=" & ESTADO_ACTIVO & _			
			"                AND VC.CDVALE IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_TRANSFERENCIA & "','" & CODIGO_VS_PRESTAMO & "')" & _
			"				 AND IDARTICULO IN (" & strSQLArt & ")" & _
			"					GROUP BY IDARTICULO " & _
			" UNION "
			'Se totalizan los ingresos por vales.
			'Debido a que en el vale de reclasificacion de stock figuran tanto los valores que egresan como los que ingresan debo separarlo para agregar la condicion Mayor/Menor 
	strSQL = strSQL & "SELECT 3 as NUMERO, IDARTICULO, SUM(EXISTENCIA) AS TOTALEXISTENCIA, SUM(SOBRANTE) AS TOTALSOBRANTE FROM  " & _
			"    TBLVALESCABECERA VC INNER JOIN TBLVALESDETALLE VD " & _
			"        ON VC.IDVALE = VD.IDVALE " & _
			"            WHERE VC.IDALMACEN=" & idAlmacen & " AND VC.FECHA > " & fechaCierreAnterior & " AND FECHA < " & fechaBusqueda & " AND VC.ESTADO=" & ESTADO_ACTIVO & _
			"                AND VC.CDVALE IN ('" & CODIGO_VS_DEVOLUCION & "','" & CODIGO_VS_ENTRADA & "','" & CODIGO_VS_RECEPCION & "')" & _
			"				 AND IDARTICULO IN (" & strSQLArt & ")" & _
			"					GROUP BY IDARTICULO " & _
			" UNION "
	strSQL = strSQL & "SELECT 4 as NUMERO, IDARTICULO, SUM(EXISTENCIA) AS TOTALEXISTENCIA, SUM(SOBRANTE) AS TOTALSOBRANTE FROM  " & _
			"    TBLVALESCABECERA VC INNER JOIN TBLVALESDETALLE VD " & _
			"        ON VC.IDVALE = VD.IDVALE " & _
			"            WHERE VC.IDALMACEN= " & idAlmacen & " AND VC.FECHA > " & fechaCierreAnterior & " AND FECHA < " & fechaBusqueda & " AND VC.ESTADO=" & ESTADO_ACTIVO & _
			"                AND VC.CDVALE IN ('" & CODIGO_VS_AJUSTE_STOCK & "','" & CODIGO_VS_FIX & "', '" & CODIGO_VS_RECLASIFICACION_STOCK & "') " & _
			"				 AND IDARTICULO IN (" & strSQLArt & ")" & _
			"					GROUP BY IDARTICULO " & _
			" UNION " 
			'Se toma el stock al último cierre.
	strSQL = strSQL & "SELECT 5 as NUMERO, IDARTICULO, EXISTENCIA AS TOTALEXISTENCIA, SOBRANTE AS TOTALSOBRANTE FROM  " & _
			"    TBLCIERRESARTICULOS2 AC " & _
			"		WHERE AC.IDALMACEN=" & idAlmacen & " AND AC.FECHACIERRE = " & fechaCierreAnterior & _
			"			AND IDARTICULO IN (" & strSQLArt & ")" & _
			") TG " & _
			"GROUP BY IDARTICULO"		
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	Set CalcularStock = rs
End Function
'--------------------------------------------------------------------------------------
Function insertStockInfo(pidArticulo, pExistencia, pSobrante)
dim UP_strSQL, UP_rs, UP_con	
	if ((pExistencia <> "0") or (pSobrante <> "0")) then
		UP_strSQL = "UPDATE TBLREPORTESTOCKWF SET EXISTENCIA=" & pExistencia & ", SOBRANTE=" & pSobrante & " WHERE IDARTICULO=" & pidArticulo & " AND FECHACIERRE = " & fechaBusqueda & " AND IDALMACEN=" & idAlmacen & " and CDUSUARIO	= '" & session("Usuario") & "' " 
		'Response.Write "<HR>UPD:" & UP_strSQL 
	end if
	call executeQueryDb(DBSITE_SQL_INTRA, UP_rs, "EXEC", UP_strSQL)
End Function
'--------------------------------------------------------------------------------------
'Se borra el workfile generado por el usuario.
Function dropArticulosTemp()
	Dim DEL_strSQL, DEL_conn, DEL_rs
		DEL_strSQL = "DELETE FROM TBLREPORTESTOCKWF "
		DEL_strSQL = DEL_strSQL & " WHERE	CDUSUARIO	= '" & session("Usuario") & "' " 
		DEL_strSQL = DEL_strSQL & " AND		IDALMACEN	= " & idAlmacen
		call executeQueryDb(DBSITE_SQL_INTRA, DEL_rs, "EXEC", DEL_strSQL)
End Function
'--------------------------------------------------------------------------------------
%>