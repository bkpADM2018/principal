<!--#include file="procedimientosUnificador.asp"-->
<%
'/*	 Constantes	*/
Const SISTEMA_ALMACENES = "ALMACENES" 
const ALMACEN_ADMIN = "A"
const ALMACEN_USUARIO = "U"
const ALMACEN_AUDITOR = "Y"
const ALMACEN_SOLICITANTE_CONTROL = "C"
const ALMACEN_SOLICITANTE = "S"
'Const 


Const TIPO_CUENTA_DEBE = "1"
Const TIPO_CUENTA_HABER = "2"
Const CUENTA_AJUSTE_STOCK = "821001030000"
Const CUENTA_OBRAENCURSO = "121101120000"
Const CUENTA_PROVISIONES = "212201990000"
Const CUENTA_INTERDIVISIONAL_ARROYO = "410001090000"
Const CUENTA_INTERDIVISIONAL_TRANSITO = "410001070000"
Const CUENTA_INTERDIVISIONAL_PIEDRABUENA = "410001100000"
Const CUENTA_PROVISION_AJT = "731309140000"
Const CCOSTO_PROVISION_AJT = "0"

Const NO_MODIFICA_STOCK = 0
Const DISMINUYE_STOCK = 1
Const AUMENTA_STOCK = 2
Const MODIFICA_STOCK = 3 'suma o resta stock, dependiendo del simbolo de la cantidad

Const REPORTE_PESO = "P" 'Igual a MONEDA_PESO
Const REPORTE_DOLAR = "D" ' Igual a MONEDA_DOLAR
Const REPORTE_CANTIDAD = "cantidad"

Const REPORTE2_SIN_STOCK = "S"
Const REPORTE2_CON_STOCK = "C"
Const REPORTE2_AMBOS     = "A"

Const GASTO = "G"
Const PROVISION = "P"
Const REVERSION_PROVISION = "R" 
Const INVENTARIO = "I"
Const MERCADERIA_TRANSITO = "T"
Const REVERSION_MERCADERIA_TRANSITO = "A"
Const FECHA_INICIO_CONTABLE = "20130101" 'Solicitado por contabilidad antes=20110101
'Prefijos para busqueda en tabla de numeraci�n.
Const PREFIX_VNR = "VNR"

'Constantes para el control de stock
Const CTST_RESULTADO_PENDIENTE = 0
Const CTST_SIN_VALE = -1
Const CTST_SELECCION_AUTOMATICA	= "A"
Const CTST_SELECCION_MANUAL		= "M"
Const CTST_ARTICULO_CON_STOCK   = "S"
Const CTST_ARTICULO_SIN_STOCK   = "N"
Const CODIGO_CPTE_CIERRES = "COAL"
Dim oConnAL, errAlmacen, ccAlmacen, ArticulosAleatorios()
'---------------------------------------------------------------------------------------------
Function controlAccesoAL(pRecurso)
	Dim rs, strSQL	
	
	strSQL = "Select * from TBLALMACENESUSUARIO WHERE CDUSUARIO='" & session("usuario") & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.EOF 
			ccAlmacen = rs("IDALMACEN") & "-" & rs("NIVEL") & ", " & ccAlmacen
		rs.MoveNext
	Wend
	if len(ccAlmacen) = 0 then 
		response.redirect "comprasAccesoDenegado.asp"
	else
		ccAlmacen = left(ccAlmacen ,len(ccAlmacen)-2)
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Obtiene la lista de almacenes PARA los cuales el usuario es Pa�olero o adminsitrador.
function obtenerListaAlmacenesUA()
	Dim strSQL, oConn, rs
	strSQL = "Select * from TBLALMACENES where ESTADO=" & ESTADO_ACTIVO & " and IDALMACEN in (Select IDALMACEN from TBLALMACENESUSUARIO WHERE CDUSUARIO='" & session("usuario") & "' and NIVEL NOT IN ('" &  ALMACEN_AUDITOR & "','" &  ALMACEN_SOLICITANTE & "')) order by DSALMACEN"

	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaAlmacenesUA = rs
End function
'---------------------------------------------------------------------------------------------
'Obtiene la lista de almacenes A los cuales el usuario puede hacer pedidos
function obtenerListaAlmacenesSolicitud()
	Dim strSQL, oConn, rs
	strSQL = "Select * from TBLALMACENES where ESTADO=" & ESTADO_ACTIVO & " and IDALMACEN in (Select IDALMACEN from TBLALMACENESUSUARIO WHERE CDUSUARIO='" & session("usuario") & "') order by DSALMACEN"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaAlmacenesSolicitud = rs
End function
'---------------------------------------------------------------------------------------------
'Devuelve lista de almacenes disponibles para el usuario
function obtenerListaAlmacenesUsuario()
	Dim strSQL, oConn, rs
	strSQL = "Select * from TBLALMACENES where ESTADO=" & ESTADO_ACTIVO & " and IDALMACEN in (Select IDALMACEN from TBLALMACENESUSUARIO WHERE CDUSUARIO='" & session("usuario") & "' and NIVEL<>'" & ALMACEN_SOLICITANTE & "') order by DSALMACEN"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaAlmacenesUsuario = rs
End function
'---------------------------------------------------------------------------------------------
function isAuditorAL(pIdAlmacen)
	isAuditorAL = isInListAL(pIdAlmacen, ALMACEN_AUDITOR)
End function
'---------------------------------------------------------------------------------------------
function isAdminAL(pIdAlmacen)
	isAdminAL = isInListAL(pIdAlmacen, ALMACEN_ADMIN)
End function
'---------------------------------------------------------------------------------------------
function isSolicitanteAL(pIdAlmacen)
	isSolicitanteAL = isInListAL(pIdAlmacen, ALMACEN_SOLICITANTE)
End function
'---------------------------------------------------------------------------------------------
function isInListAL(pIdAlmacen, pCargo)
	dim i, mySplit, mySplit2, auxAlm
	mySplit = split(ccAlmacen,",") 
	for i=0 to ubound(mySplit)
		mySplit2 = split(mySplit(i),"-") 
		if cint(mySplit2(0)) = cint(pIdAlmacen) then
			if (mySplit2(1) = pCargo) then 
				isInListAL = true
				exit function
			end if	
		end if
	next
	isInListAL = false
End function
'---------------------------------------------------------------------------------------------
'Devuelve la lista de divisiones en las que tiene permiso de A, Y o U
Function getListaCargosAdmin()

dim rtrn, con, rs, strSQL
	
	strSQL = "Select Distinct C.IDDIVISION from TBLALMACENESUSUARIO A inner join TBLALMACENES B on A.IDALMACEN=B.IDALMACEN inner join TBLDIVISIONES C on B.IDDIVISION=C.IDDIVISION WHERE A.CDUSUARIO='" & session("usuario") & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		rtrn = "," & rs("IDDIVISION") & rtrn
		rs.MoveNext()
	wend
	if (len(rtrn) > 0) then
		getListaCargosAdmin = right(rtrn,len(rtrn)-1)
	else
		getListaCargosAdmin = ""
	end if	
End Function
'--------------------------------------------------------------------------------------
Function getCategoriaFull(idCategoria)
	Dim rs, conn, strSQL, ret
	
	strSQL="Select * from TBLARTCATEGORIAS where IDCATEGORIA=" & idCategoria
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	ret = ""
	if (not rs.eof) then ret = rs("CDCATEGORIA") & " - " & rs("DSCATEGORIA") 
	getCategoriaFull = ret
		
End Function
'--------------------------------------------------------------------------------------
Function getAlmacenesPorDivision(idDivision)
dim rs, conn, strSQL, ret, almacenList
ret = ""
strSQL="Select * from TBLALMACENES where IDDIVISION=" & idDivision
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		ret = rs("IDALMACEN") & "," & ret
	rs.movenext
wend	
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
if len(ret) > 1 then
	'Controlar acceso a la Division
	Response.Write almacenList 
	ret = left(ret,len(ret)-1)
end if	
getAlmacenesPorDivision = ret
end function
'--------------------------------------------------------------------------------------
Function getDivisionAlmacen(pIdAlmacen)
	Dim strSQL, ret, conn,rs
	ret=0
	strSQL="Select * from TBLALMACENES where IDALMACEN=" & pIdAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then ret= rs("IDDIVISION")
	getDivisionAlmacen = ret
End Function
'---------------------------------------------------------------------------------------
'Arma la lista con todos los articulos no recibidos por proveedor.
'Parametros
'	pIdDivision
'	pIdProveedor	(Optativo [pasar 0]) 
'	pListaArticulos	(Optativo [pasar '']) 
'	pIdPedido		(Optativo [pasar ''])
'	pFiltro			(Optativo [pasar ''])
Function obtenerArticulosPedidosNoRecibidos(pIdDivison, pIdProveedor, pListaArticulos, pIdPedido, pFiltro)
	
	Dim strSQL, rs, conn, strSQL1, strSQL2
	
	strSQL1=""
	strSQL2=""
	if (pIdProveedor <> 0) then
		strSQL1 = " inner join TBLREMCABECERA RC on RC.IDREMITO=RP.IDREMITO where RC.IDPROVEEDOR=" & pIdProveedor		
		strSQL2 = " and CAB.IDPROVEEDOR=" & pIdProveedor
	end if	
	if (pIdPedido <> "")     then strSQL2 = strSQL2 & " and CAB.IDCOTIZACION =" & pIdPedido
	if (pListaArticulos <> "") and (pListaArticulos <> "0") then strSQL2 = strSQL2 & " and PEDIDO.IDARTICULO in (" & pListaArticulos & ")"
	if (pFiltro <> "")     then strSQL2 = strSQL2 & pFiltro
	
	strSQL="			Select PEDIDO.IDARTICULO,ART.DSARTICULO,UNI.ABREVIATURA UNIDAD, CAB.IDPROVEEDOR,EMP.nomemp DSPROVEEDOR, sum(PEDIDO.CANTIDAD) CANTIDADP, sum(RECIBIDO.CANTIDAD) CANTIDADR"
	strSQL= strSQL & "	from ("
	strSQL= strSQL & "		Select PIC.* from TBLCTZCABECERA PIC" 
	strSQL= strSQL & "		inner join 	(Select IDCOTIZACION IDPIC from TBLCTZCABECERA"
	strSQL= strSQL & "					EXCEPT" 
	strSQL= strSQL & "					Select IDPIC from TBLREMPIC where IDREMITO=0) NPIC on PIC.IDCOTIZACION=NPIC.IDPIC"
	strSQL= strSQL & "		) CAB inner join  TBLCTZDETALLE PEDIDO on CAB.IDCOTIZACION=PEDIDO.IDCOTIZACION"	
	strSQL= strSQL & "	left join  ("
	strSQL= strSQL & "	    Select RP.IDPIC, RP.IDARTICULO, sum(RP.CANTIDAD) CANTIDAD "
	strSQL= strSQL & "	    from TBLREMPIC RP " & strSQL1	
	strSQL= strSQL & "	    group by RP.IDPIC, RP.IDARTICULO"
	strSQL= strSQL & "	    ) RECIBIDO on PEDIDO.IDCOTIZACION = RECIBIDO.IDPIC and PEDIDO.IDARTICULO=RECIBIDO.IDARTICULO"
	strSQL= strSQL & "  INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO = PEDIDO.IDARTICULO "
	strSQL= strSQL & "  INNER JOIN [Database].[dbo].met001a EMP ON EMP.nroemp = CAB.IDPROVEEDOR"
	strSQL= strSQL & "  INNER JOIN TBLUNIDADES UNI ON UNI.IDUNIDAD = ART.IDUNIDAD"
	strSQL= strSQL & "  INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA = CAT.IDCATEGORIA"
						'Se obtienen datos complementarios de las obras
	strSQL= strSQL & "  LEFT JOIN TBLDATOSOBRAS OBR on CAB.IDOBRA=OBR.IDOBRA"
	strSQL= strSQL & "	where (PEDIDO.CANTIDAD > RECIBIDO.CANTIDAD or RECIBIDO.CANTIDAD is null)"
	'strSQL= strSQL & "		and ART.BIENUSO <>'" & ES_BIEN_DE_USO & "'"		
	strSQL= strSQL & "		and (CAB.ESTADO='" & CTZ_FIRMADA & "' or CAB.ESTADO='" & CTZ_FACTURADA & "') and CAT.TIPOCATEGORIA = '" & TIPO_CAT_BIENES & "' and CAB.IDDIVISION=" & pIdDivison & " AND PEDIDO.cantidad > 0 " & strSQL2
	strSQL= strSQL & "	group by PEDIDO.IDARTICULO, CAB.IDPROVEEDOR,DSARTICULO,EMP.nomemp,UNI.ABREVIATURA"
	strSQL= strSQL & "  order by PEDIDO.IDARTICULO,IDPROVEEDOR"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerArticulosPedidosNoRecibidos = rs	
	
End Function
'---------------------------------------------------------------------------------------------
Function obtenerSectores(pIdSector)
	Dim strSQL, rs
	
	strSQL="Select * from TBLSECTORES"	
	if (pIdSector <> "") then strSQL = strSQL & " where IDSECTOR=" & pIdSector
	strSQL= strSQL & " order by DSSECTOR"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	Set obtenerSectores = rs
End Function
'---------------------------------------------------------------------------------------------
Function validarSector(pIdSector)	
	Dim strSQL, rs, ret
	strSQL="Select * from TBLSECTORES where IDSECTOR=" & pIdSector	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	ret=false
	if (not rs.eof) then ret = true
	validarSector = ret
End Function
'--------------------------------------------------------------------------------------------------------
function getCuentaInterdivisional(pIdDivision)
dim myCdDivision, rtrn
myCdDivision = getDivisionAbreviada(pIdDivision)
if myCdDivision = CODIGO_ARROYO then
	rtrn = CUENTA_INTERDIVISIONAL_ARROYO
elseif myCdDivision = CODIGO_PIEDRABUENA then
	rtrn = CUENTA_INTERDIVISIONAL_PIEDRABUENA		
elseif myCdDivision = CODIGO_TRANSITO then
	rtrn = CUENTA_INTERDIVISIONAL_TRANSITO		
end if
getCuentaInterdivisional = rtrn
end function
'----------------------------------------------------------------------------------------
function getIdCierre2(pAnio, pMes, pDivision, pEstado)
dim strSQL, rs, con, idCierre
idCierre = 1
	strSQL = "SELECT * FROM TBLCIERRESCABECERA2 WHERE ANIO=" & pAnio & " AND MES = " & pMes & " AND IDDIVISION=" & pDivision & " AND ESTADO='" & pEstado & "'"
	'Response.Write "<BR>" & strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		idCierre = rs("IDCIERRE")
	else
		sqlINS = "INSERT INTO TBLCIERRESCABECERA2 (ANIO, MES, IDDIVISION, ESTADO, CDUSUARIO, MOMENTO) VALUES(" & pAnio & "," & pMes & _
   				"," & pDivision & ",'" & pEstado & "','" & session("usuario") & "'," & session("momentodato") & ")"
		'Response.Write "<BR>" & sqlINS
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)

		strSQL = "SELECT MAX(IDCIERRE) as MAX_ID FROM TBLCIERRESCABECERA2"
		'Response.Write "<BR>" & strSQL
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if not rs.eof then idCierre = rs("MAX_ID")
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)

		sqlINS = "INSERT INTO TBLCIERRESFIRMAS2 VALUES(" & idCierre & "," & FIRMA_ROL_RESP_CONTADURIA & ",'" & FIRMA_NO_USER & "'," & session("MmtoDato") & ", '99999999')"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)
		'Response.Write "<BR>" & sqlINS
		sqlINS = "INSERT INTO TBLCIERRESFIRMAS2 VALUES(" & idCierre & "," & FIRMA_ROL_RESP_PUERTO & ",'" & FIRMA_NO_USER & "'," & session("MmtoDato") & ", '99999999')"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)
		'Response.Write "<BR>" & sqlINS	
	
	end if	
getIdCierre2 = idCierre 
end function
'----------------------------------------------------------------------------------------
function getIdCierre(pAnio, pMes, pDivision, pEstado)
dim strSQL, rs, con, idCierre
idCierre = 1
	strSQL = "SELECT * FROM TBLCIERRESCABECERA WHERE ANIO=" & pAnio & " AND MES = " & pMes & " AND IDDIVISION=" & pDivision & " AND ESTADO='" & pEstado & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		idCierre = rs("IDCIERRE")
	else
		sqlINS = "INSERT INTO TBLCIERRESCABECERA (ANIO, MES, IDDIVISION, ESTADO, CDUSUARIO, MOMENTO) VALUES(" & pAnio & "," & pMes & _
   				"," & pDivision & ",'" & pEstado & "','" & session("usuario") & "'," & session("momentodato") & ")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)

		strSQL = "SELECT MAX(IDCIERRE) as MAX_ID FROM TBLCIERRESCABECERA"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if not rs.eof then idCierre = rs("MAX_ID")
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
		sqlINS = "INSERT INTO TBLCIERRESFIRMAS VALUES(" & idCierre & "," & FIRMA_ROL_RESP_CONTADURIA & ",'" & FIRMA_NO_USER & "'," & session("MmtoDato") & ", '99999999')"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)
		sqlINS = "INSERT INTO TBLCIERRESFIRMAS VALUES(" & idCierre & "," & FIRMA_ROL_RESP_PUERTO & ",'" & FIRMA_NO_USER & "'," & session("MmtoDato") & ", '99999999')"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsINS, "EXEC", sqlINS)
	end if	
getIdCierre = idCierre 
end function
'-------------------------------------------------------------------------------------------------'
' Funcion:  cargarCantidadesPedidas
' Parametros:
'			excluyentes: solo 1 de los 2 parametros es necesario para el correcto funcionamiento
'						 de la funcion. Siendo la division predominante ante el almacen
' Devuelve: La funcion devuelve un diccionario. Siendo su key el idarticulo y su value
'			la cantidad de stock en PICs que aun no han sido entregados.
'-------------------------------------------------------------------------------------------------'
Function cargarCantidadesPedidas(pIdAlmacen,pIdDivision)
	Dim strSQL, rs2,conn,rtrn,auxCantidad,oDiccCantidadesPedidas
	Set oDiccCantidadesPedidas  = createObject("Scripting.Dictionary")

	strSQL =          "SELECT a.cantidadPics - a.cantidadRems cantidad, "
	strSQL = strSQL & "       a.idarticulo "
	strSQL = strSQL & "FROM   ( SELECT  SUM(d.cantidad)          cantidadPics, "
	strSQL = strSQL & "                ISnull(SUM(r.cantidad),0) cantidadRems, "
	strSQL = strSQL & "                d.idarticulo "
	strSQL = strSQL & "       FROM     tblctzdetalle d "
	strSQL = strSQL & "                INNER JOIN "
	strSQL = strSQL & "					(Select * from tblctzcabecera where estado NOT IN ('"&CTZ_ANULADA&"')"
	if (pIdDivision = 0) then
		strSQL = strSQL & "						AND    iddivision          = " & getDivisionAlmacen(pIdAlmacen)
	else
		strSQL = strSQL & "						AND    iddivision          = " & pIdDivision
	end if
	strSQL = strSQL & "					) c "
	strSQL = strSQL & "                ON       d.idcotizacion = c.idcotizacion "
	strSQL = strSQL & "                LEFT JOIN (Select IDPIC, IDARTICULO, SUM(CANTIDAD) CANTIDAD from tblrempic group by IDPIC, IDARTICULO) r "
	strSQL = strSQL & "                ON       r.idpic      = c.idcotizacion "
	strSQL = strSQL & "                AND      r.idarticulo = d.idarticulo "
	strSQL = strSQL & "                AND      r.cantidad  <> 0 "	
	strSQL = strSQL & "       WHERE      d.idarticulo IN "
										'articulos con flatante de stock'
	strSQL = strSQL & "                (SELECT  a.idarticulo "
	strSQL = strSQL & "                FROM     (SELECT * "
	strSQL = strSQL & "                         FROM    tblarticulosdatos "
	if (pIdDivision = 0) then
		strSQL = strSQL & "                         WHERE   idalmacen = " & pIdAlmacen
	else
		strSQL = strSQL & "                         WHERE   idalmacen in (" & getAlmacenesPorDivision(pIdDivision) & ") "
	end if
	strSQL = strSQL & "                         AND "
	strSQL = strSQL & "                                 ( "
	strSQL = strSQL & "                                         existencia + sobrante "
	strSQL = strSQL & "                                 ) "
	strSQL = strSQL & "                                              < stockminimo "
	strSQL = strSQL & "                         AND     stockminimo <> 0 "
	strSQL = strSQL & "                         ) "
	strSQL = strSQL & "                         a "
	strSQL = strSQL & "                         INNER JOIN tblarticulos art "
	strSQL = strSQL & "                         ON       a.idarticulo = art.idarticulo "
	strSQL = strSQL & "                 "
	strSQL = strSQL & "                ) "
	strSQL = strSQL & "       GROUP BY d.idarticulo "
	strSQL = strSQL & "       ) "
	strSQL = strSQL & "       a "
	strSQL = strSQL & "WHERE  a.cantidadPics - a.cantidadRems > 0"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)

	while not rs2.EoF
		auxCantidad = cdbl(rs2("cantidad"))

		if (not oDiccCantidadesPedidas.Exists(cdbl(rs2("idArticulo")))) then
			Call oDiccCantidadesPedidas.Add(cdbl(rs2("idArticulo")),auxCantidad)
		else
			oDiccCantidadesPedidas.Item(cdbl(rs2("idArticulo"))) = cdbl(oDiccCantidadesPedidas.Item(cdbl(rs2("idArticulo")))) + auxCantidad
		end if
		rs2.MoveNext
	wend

	Set cargarCantidadesPedidas = oDiccCantidadesPedidas
End Function
'-------------------------------------------------------------------------------------------------------------------
function valorizarValeContable(pIdVale, pFechaCierre)
' Esta funcion podra ser llamada desde cualquier ubicacion y solo sera necesario que se le pase el id de vale
' que se desea valorizar y la fecha a la cual se valoriza.
dim strSQL, rs, rtrn, myPrecioVigente, myPrecioVigenteD, myCantidadValuada, pudoValuar
rtrn = false
'Solo se valorizar�n vales activos, que hayan movido existencia y que tengan unidades que aun no hayan sido valorizadas.
strSQL = "SELECT  T.CDVALE, T.IDARTICULO, T.IDALMACEN, T.EXISTENCIA , CASE WHEN SALDO IS NULL THEN EXISTENCIA ELSE SALDO END AS SALDO_ITEM FROM " & _
		 "  (" & _
		 "	SELECT CDVALE, VD.IDARTICULO, IDALMACEN, EXISTENCIA-TOTAL_VALUADO AS SALDO, EXISTENCIA FROM TBLVALESCABECERA VC " & _
		 "		INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _
		 "		LEFT JOIN " & _
		 "			(SELECT IDVALE, IDARTICULO, SUM(CANTIDAD) AS TOTAL_VALUADO FROM TBLVALESCONTABLE WHERE IDVALE=" & pIdVale & " GROUP BY IDVALE, IDARTICULO) AS VAL " & _
		 "		ON VD.IDVALE=VAL.IDVALE AND VD.IDARTICULO=VAL.IDARTICULO " & _
		 "	  WHERE ESTADO=" & ESTADO_ACTIVO & " AND VD.EXISTENCIA<>0 AND VC.IDVALE = " & pIdVale & _
		 "  )T " & _
		 "	WHERE T.SALDO <> 0 OR T.SALDO IS NULL "
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	select case rs("CDVALE")
		case CODIGO_VS_RECLASIFICACION_STOCK
		'RECLASIFICACIONES
		'Este tipos de vales requiere un proceso un tanto diferente a los demas debido a que el precio del articulo
		'que recibe la mercaderia surge de el o los articulos origenes. por lo tanto no se va a buscar a la tabla
		'de valuaciones directamente.
			dim myCantidadAValuar, myPrecioPromedio, myPrecioPromedioD, myTotalUnidades, myIdProductoDestino, dicPrecios 
			myPrecioPromedio = 0
			myTotalUnidades = 0
			set dicPrecios = Server.CreateObject("Scripting.Dictionary")
			do while not rs.eof 'Puede haber mas de un registro de origen. El negativo es el qeu pierde stock, el positivo es el que lo gana.
					myCantidadAValuar = cdbl(rs("SALDO_ITEM"))
					if myCantidadAValuar<0 then 'Es el articulo qeu pierde stock. Tiene el vlu
						pudoValuar = getArticuloContable(pFechaCierre, rs("IDALMACEN"), rs("IDARTICULO"), abs(myCantidadAValuar), myPrecioVigente, myPrecioVigenteD, myCantidadValuada)
						'Se valuaran tantas unidades como stockdisponible halla
						if pudoValuar then
							'Guardo el precio contable obtenido para guardarlo una vez que el proceso indique que se pudo valuar
							call dicPrecios.Add("P_" & rs("IDARTICULO"),round(myPrecioVigente,0))
							call dicPrecios.Add("D_" & rs("IDARTICULO"),round(myPrecioVigenteD,0))
							call dicPrecios.Add("C_" & rs("IDARTICULO"),myCantidadValuada)
							Response.Write "EN EL DIC($" & dicPrecios("P_" & rs("IDARTICULO")) & ", u$s" & dicPrecios("D_" & rs("IDARTICULO")) & ")"
							'Acumular por si son mas de un articulos los que pierden stock
							myPrecioPromedio = ((cdbl(myPrecioVigente) * cdbl(myCantidadValuada)) + (cdbl(myPrecioPromedio) * cdbl(myTotalUnidades))) / (cdbl(myCantidadValuada) + cdbl(myTotalUnidades))
							myPrecioPromedioD = ((cdbl(myPrecioVigenteD) * cdbl(myCantidadValuada)) + (cdbl(myPrecioPromedioD) * cdbl(myTotalUnidades))) / (cdbl(myCantidadValuada) + cdbl(myTotalUnidades))
							myTotalUnidades = CDbl(myCantidadValuada) + CDbl(myTotalUnidades)
						else
							'No se pudo valuar el vale pues algun origen no tenia precio. Quedara para otro cierre
							'exit do
						end if
					else
						'Es el articulo que recibe la mercaderia. Lo guardo para luego asignarle el precio que corresponda
						myIdProductoDestino = rs("IDARTICULO")
					end if
				rs.movenext
			loop
			if myTotalUnidades <> 0 then
				'En esta etapa recien se quitan las unidades contables a los origenes y se setea el el precio y unidades ganadas 
				'al destino dado que recien aqui nos aseguramos que pudo valuar
				rs.movefirst
				while not rs.eof
					myCantidadAValuar = cdbl(rs("SALDO_ITEM"))
					if myCantidadAValuar<0 then 'Es el articulo qeu pierde stock
						if CDbl(dicPrecios("C_" & rs("IDARTICULO"))) > 0 then 'Solo se setea si se pudo valuar!
							'Setear el precio actual del articulo que pierde la mercaderia
							call setValuacionContable(pFechaCierre, pIdVale, rs("IDARTICULO"), dicPrecios("P_" & rs("IDARTICULO")), dicPrecios("D_" & rs("IDARTICULO")), CDbl(dicPrecios("C_" & rs("IDARTICULO")) )*-1)
							'Se sacan las unidades contables para que otros VRS no las tome en cuenta.
							call quitarUnidadesContables(pFechaCierre,rs("IDALMACEN"),rs("IDARTICULO"),ABS(CDbl(dicPrecios("C_" & rs("IDARTICULO")) )))
						end if
					else
						'Es el que gano el stock establecer el nuevo precio	
						call setValuacionContable(pFechaCierre, pIdVale, myIdProductoDestino, round(myPrecioPromedio,0), round(myPrecioPromedioD,0), myTotalUnidades)
					end if	
					rs.movenext
				wend	
				rtrn = true							
			end if
		'FIN RECLASIFICACIONES			

		case CODIGO_VS_RECEPCION 
			'RECEPCIONES
			'Obtener el di de la transferencia
			call obtenerVMT(pIdVale, myIdVMT, myAlmacenVMT)
			if myIdVMT > 0 then 
				while not rs.eof 
					myCantidadAValuar = cdbl(rs("SALDO_ITEM"))

					pudoValuar = getArticuloContable(pFechaCierre, myAlmacenVMT, rs("IDARTICULO"), abs(myCantidadAValuar), myPrecioVigente, myPrecioVigenteD, myCantidadValuada)
					Response.Write "<br>Cantidad a Valuar VMT(" & myIdVMT & "), VMR(" & pIdVale & "), (" &  myCantidadAValuar & "), pudo(" & pudoValuar & ")"
					if pudoValuar then 
						'ORIGEN
						Call setValuacionContable(pFechaCierre, myIdVMT, rs("IDARTICULO"), myPrecioVigente, myPrecioVigenteD, CDbl(myCantidadValuada))
						Call quitarUnidadesContables(pFechaCierre, myAlmacenVMT, rs("IDARTICULO"), CDbl(myCantidadValuada))
						'DESTINO
						Call setValuacionContable(pFechaCierre, pIdVale, rs("IDARTICULO"), myPrecioVigente, myPrecioVigenteD, CDbl(myCantidadValuada))
					end if	
					rs.movenext
				wend	
			else
				'ERROR - NO HAY VALE DE TRANSFERENCIA!!!
			end if	
			'FIN RECEPCIONES			
		case CODIGO_VS_SALIDA, CODIGO_VS_AJUSTE_VALE, CODIGO_VS_AJUSTE_TRANSFERENCIA, CODIGO_VS_AJUSTE_STOCK, CODIGO_VS_FIX
			'MULTIPLES
				Response.Write "<hr>MULTI INICIO"
				while not rs.eof
						myCantidadAValuar = cdbl(rs("SALDO_ITEM"))
						pudoValuar = getArticuloContable(pFechaCierre, rs("IDALMACEN"), rs("IDARTICULO"), abs(myCantidadAValuar), myPrecioVigente, myPrecioVigenteD, myCantidadValuada)
						if pudoValuar then
							if rs("CDVALE") = CODIGO_VS_AJUSTE_STOCK then myCantidadValuada = CDbl(myCantidadValuada) * -1
							Call quitarUnidadesContables(pFechaCierre,rs("IDALMACEN"),rs("IDARTICULO"),Abs(CDbl(myCantidadValuada)))
							Call setValuacionContable(pFechaCierre, pIdVale, rs("IDARTICULO"), myPrecioVigente, myPrecioVigenteD, CDbl(myCantidadValuada))
						end if
					rs.movenext
				wend	
				Response.Write "<hr>MULTI FIN"
				'FIN MULTIPLES			
		case CODIGO_VS_TRANSFERENCIA
			'TRANSFERENCIAS
			Response.Write "<hr>TRANSFERENCIAS INICIO"
			while not rs.eof
				myCantidadAValuar = cdbl(rs("SALDO_ITEM"))
				pudoValuar = getArticuloContable(pFechaCierre, rs("IDALMACEN"), rs("IDARTICULO"), abs(myCantidadAValuar), myPrecioVigente, myPrecioVigenteD, myCantidadValuada)
				if pudoValuar then
					Call setValuacionContable(pFechaCierre, pIdVale, rs("IDARTICULO"), myPrecioVigente, myPrecioVigenteD, CDbl(myCantidadValuada))
					Call quitarUnidadesContables(pFechaCierre,rs("IDALMACEN"),rs("IDARTICULO"),Abs(CDbl(myCantidadValuada)))
				end if
				rs.movenext
			wend	
			Response.Write "<hr>TRANSFERENCIAS FIN"
			'FIN TRANSFERENCIAS
		case else

		end select
end if
valorizarValeContable = rtrn
end function
'--------------------------------------------------------------------------------------------------------------
function getArticuloContable(pFechaCierre, pIdAlmacen, pIdArticulo, pCantidadAValuar, byref myPrecioVigente, byref myPrecioVigenteD, byref myCantidadValuada)
'Se cargan por referencia el precio contable del articulo y la cantidad que se podria valuar.
dim strSQL, rs, rtrn
rtrn = false
strSQL = "SELECT * FROM TBLARTVALUACION WHERE IDDIVISION=(SELECT IDDIVISION FROM TBLALMACENES WHERE IDALMACEN=" & pIdAlmacen & ")" & _
		 " AND IDARTICULO=" & pIdArticulo & " AND STOCKDISPONIBLE>0 AND FECHACIERRE=" & pFechaCierre 
'Response.Write "<br>get Articulo Contable <BR>" & strSQL & "<BR>"			 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	'Response.Write "1(" & cdbl(rs("STOCKDISPONIBLE")) & "), 2(" & cdbl(pCantidadAValuar) & ")"
	if cdbl(rs("STOCKDISPONIBLE")) >= cdbl(abs(pCantidadAValuar)) then 
		myCantidadValuada = pCantidadAValuar
	else
		myCantidadValuada = rs("STOCKDISPONIBLE")
	end if	
	myPrecioVigente = rs("VLUPESOS")
	myPrecioVigenteD = rs("VLUDOLARES")
	rtrn = true
end if		
getArticuloContable = rtrn  
end function
'--------------------------------------------------------------------------------------------------------------
sub setValuacionContable(pFechaCierre, pIdVale, pIdArticulo, pPrecioPesos, pPrecioDolares, pCantidad)
dim strSQL, rs, mySecuencia
mySecuencia = 1
strSQL = "SELECT * FROM TBLVALESCONTABLE WHERE IDVALE=" & pIdVale & " AND IDARTICULO=" & pIdArticulo
'Response.Write "<br>set Valuacion Contable <BR> " & strSQL & "<BR>"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then	mySecuencia = cint(rs("SECUENCIA")) + 1
strSQL = "INSERT INTO TBLVALESCONTABLE VALUES(" & pFechaCierre & "," & pIdVale & "," & pIdArticulo & "," & pCantidad & "," & pPrecioPesos & "," & pPrecioDolares & "," & mySecuencia & "," & session("momentodato") & ")"
'Response.Write "<br>set Valuacion Contable EXEC <BR>" & strSQL & "<BR>"			 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

end sub
'--------------------------------------------------------------------------------------------------------------
sub quitarUnidadesContables(pFechaCierre, pIdAlmacen, pIdArticulo, pCantidadAQuitar)
'Resta unidades contables de un articulo. Para sumar si o si hay que realizar recalculo de precios
dim strSQL, rs
strSQL = "SELECT * FROM TBLARTVALUACION WHERE IDDIVISION=(SELECT IDDIVISION FROM TBLALMACENES WHERE IDALMACEN=" & pIdAlmacen & ")" & _
		 " AND IDARTICULO=" & pIdArticulo & " AND STOCKDISPONIBLE>0 AND FECHACIERRE=" & pFechaCierre 
'Response.Write "<br>quitarUnidadesContables <BR>" & strSQL & "<BR>"			 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	strSQL = "UPDATE TBLARTVALUACION SET STOCKDISPONIBLE=STOCKDISPONIBLE-" & pCantidadAQuitar & " WHERE IDDIVISION=(SELECT IDDIVISION FROM TBLALMACENES WHERE IDALMACEN=" & pIdAlmacen & ")" & _
		 " AND IDARTICULO=" & pIdArticulo & " AND STOCKDISPONIBLE>=" & pCantidadAQuitar & " AND FECHACIERRE=" & pFechaCierre 
	'Response.Write "<br>quitarUnidadesContables EXEC <BR>" & strSQL & "<BR>"			 
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end if		
end sub
'--------------------------------------------------------------------------------------------------------------
sub obtenerVMT(pIdVMR, byref pIdVMT, byref pAlmacenVMT)
dim strSQL, rs
pIdVMT=0
pAlmacenVMT=0
strSQL = "SELECT IDVALE, IDALMACEN FROM TBLVALESCABECERA VC  " & _
		 "  WHERE PARTIDAPENDIENTE IN (SELECT PARTIDAPENDIENTE FROM TBLVALESCABECERA WHERE IDVALE= " & pIdVMR & ")" & _
		 "		AND CDVALE='" & CODIGO_VS_TRANSFERENCIA & "'"
'Response.Write "<hr>obtenerVMT " & strSQL & "<hr>"	
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then 
	pIdVMT=clng(rs("IDVALE"))
	pAlmacenVMT =Cint(rs("IDALMACEN"))
end if	
end sub
'---------------------------------------------------------------------------------------------------------------
sub acumularArticulo(pIdDivision, pIdArticulo, pUnidades, pImportePesos, pImporteDolares)
'Esta funcion recibe el articulo que se debe acumular en el cierre
dim rs, strSQL, newVluPesos, newVluDolares, newStockDisponible

strSQL = "SELECT * FROM TBLARTVALUACION WHERE IDDIVISION=" & pIdDivision & " AND IDARTICULO=" & pIdArticulo & " AND FECHACIERRE=" & fechaCierre
'Response.Write "<BR>acumularArticulo EXIST " & strSQL 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	Response.Write "<hr> VALORES<br>PRECIO ANTERIOR $" & rs("VLUPESOS") & ", u$s" & rs("VLUDOLARES") & ",STOCK ANTERIOR " & rs("STOCKDISPONIBLE") 
	Response.Write "<br>IMPORTE COMPRADO $" & pImportePesos & ", u$s" & pImporteDolares & ",UNIDADES COMPRADAS " & pUnidades
	if (cdbl(rs("STOCKDISPONIBLE")) + cdbl(pUnidades)) > 0 then
		newVluPesos = round((cDbl(rs("VLUPESOS"))*cDbl(rs("STOCKDISPONIBLE")) + (cDbl(pImportePesos))) /   (cDbl(rs("STOCKDISPONIBLE"))+cDbl(pUnidades)),0)
		newVluDolares = round((cDbl(rs("VLUDOLARES"))*cDbl(rs("STOCKDISPONIBLE")) + (cDbl(pImporteDolares))) /   (cDbl(rs("STOCKDISPONIBLE"))+cDbl(pUnidades)),0)
	else
		newVluPesos = rs("VLUPESOS")
		newVluDolares = rs("VLUDOLARES") 
	end if	
	Response.Write "<br>NUEVO PRECIO $" & newVluPesos & ", u$s" & newVluDolares
	newStockDisponible = cDbl(rs("STOCKDISPONIBLE")) + cDbl(pUnidades)
	strSQL = "UPDATE TBLARTVALUACION SET VLUPESOS=" & newVluPesos & ", VLUDOLARES=" & newVluDolares & ", STOCKDISPONIBLE=" & newStockDisponible & _
			 ", MMTOCALCULO=" & session("momentodato") & ", CDUSUARIO='" & session("usuario") & "' WHERE FECHACIERRE=" & fechaCierre & " AND IDDIVISION=" & pIdDivision & " AND IDARTICULO=" & pIdArticulo 
else
	'Es un articulo que aparecio este mes
	newVluPesos = round((cdbl(pImportePesos) / cdbl(pUnidades)),0)
	newVluDolares = round((cdbl(pImporteDolares) / cdbl(pUnidades)),0)
	strSQL = "INSERT INTO TBLARTVALUACION VALUES(" & fechaCierre & "," & pIdDivision & "," & pIdArticulo & _
			 "," & newVluPesos & "," & newVluDolares & "," & pUnidades & "," & session("momentodato") & ",'" & session("usuario") & "')" 
end if
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
Response.Write "<BR>GRABAR " & strSQL 
end sub
'--------------------------------------------------------------------------------------------
Function getCantidadArticulosRegistrados()
	Dim strSQL ,rs,oConn
	strSQL = 		 " SELECT row_number() OVER (ORDER BY IDARTICULO ASC) as IDALEATORIO ,"
	strSQL = strSQL &" IDARTICULO "
	strSQL = strSQL &" FROM TBLARTICULOS WHERE ESTADO = 1"	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getCantidadArticulosRegistrados = rs	
End function
'----------------------------------------------------------------------------------------
Function DevolverIdcontrolReciente()
	Dim strSQL,rs,oConn,rtrn 
	rtrn= 0
	strSQL = " SELECT MAX(IDCONTROL) AS IDCONT FROM TBLCSTKCABECERA "
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	rtrn = rs("IDCONT")	
	DevolverIdcontrolReciente = rtrn
end function
'-----------------------------------------------------------------------------------------
Function AgregarControlStockCab(cdResponsable,IdAlmacen,v_strTipo,pChkStock,pPrecioMinimo,pCantidadArticulos,pPrecioMaximo)
	Dim strSQL,rs,oConn,rtrn
	rtrn = 0	
	strSQL = " INSERT into TBLCSTKCABECERA (CDRESPONSABLE,IDALMACEN,MOMENTO,TIPO,IDRESULTADO,CDUSUARIO, IDESTADO, CANTIDAD, PRECIOMINIMO,PRECIOMAXIMO, ARTCONSTOCK) " 
	strSQL = strSQL &" VALUES('"&cdResponsable&"',"&IdAlmacen&",'"&session("MmtoSistema")&"','"&v_strTipo&"'," & CTST_RESULTADO_PENDIENTE & ",'"&session("Usuario")&"', " & ESTADO_ACTIVO & ", " & pCantidadArticulos & ", " & pPrecioMinimo & "," & pPrecioMaximo & ", '" & pChkStock & "')"	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = " SELECT MAX(IDCONTROL) AS IDCONT FROM TBLCSTKCABECERA "	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	rtrn = rs("IDCONT")
	AgregarControlStockCab = rtrn
End function 
'-----------------------------------------------------------------------------------------
Function AgregarControlStockDet(IdControl,idArticulo,StockSist,StockFis,pValPeso)
	Dim strSQL,rs,oConn
	strSQL = " INSERT INTO TBLCSTKDETALLE(IDCONTROL,IDARTICULO,STOCKSISTEMA,STOCKFISICO,VLUPESOS) "
	strSQL = strSQL &" VALUES("&IdControl&","&idArticulo&","&StockSist&","&StockFis&","&pValPeso&")"		
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End function 
'-----------------------------------------------------------------------------------------
function getLeyendaTipoValuacion(pCdTipoValuacion)
	select case pCdTipoValuacion
		case GASTO
			getLeyendaTipoValuacion = "GASTO"
		case PROVISION
			getLeyendaTipoValuacion = "PROVISION"
		case REVERSION_PROVISION
			getLeyendaTipoValuacion = "REV PROVISION"
		case INVENTARIO
			getLeyendaTipoValuacion = "INVENTARIO"
		case MERCADERIA_TRANSITO
			getLeyendaTipoValuacion = "MERC TRANSITO"
		case REVERSION_MERCADERIA_TRANSITO
			getLeyendaTipoValuacion = "REV MERC TRANS"
	end select
end function
'----------------------------------------------------------------------------------------
sub actualizarEstadoCierre(pIdCierre, pEstado)
dim strSQL, rs, oConn
strSQL ="UPDATE TBLCIERRESCABECERA2 SET ESTADO='" & pEstado & "' WHERE IDCIERRE = " & pIdCierre
'Response.Write "<HR>" & strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------
' Funci�n:	leerDetallesCtSt
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	- 
' Modifico: CNA - Ajaya Nahuel
' Fecha:	25/02/2013
' Objetivo:	
'			Leer el Detalle de un determinado Control de Stock
' Parametros:
'			pIdControl  		[int] 	  ID CONTROL
'			pIdAlmacen 			[int] 	  ID almacen
'			flagDatosCargados	[Boolean] True = tiene datos Cargados, false = no los tiene 	
'			pSeleccion			[string]  'A' = Automatica, 'M' = Manual
' Devuelve:
'			Record Set
'Observaciones:
'			En caso de que no se hayan cargados los resultados va y busca el stock fisico y el precio de cada articulo del control
'			En caso de que si este cargado los levanta del Detalle del Control
'--------------------------------------------------------------------------------------------
Function leerDetallesCtSt(pIdControl,pIdAlmacen,flagDatosCargados, pSeleccion)
	dim strSQL, rs, myCampo, myStockSist, myStockFis
	if(flagDatosCargados)then
		myCampo = " A.STOCKSISTEMA AS STOCKSISTEMA, A.STOCKFISICO, A.VLUPESOS, "
	else
		myStockSist = "(C.EXISTENCIA + C.SOBRANTE)"		
		myStockFis =  "STOCKFISICO" 
		myCampo = " CASE WHEN "& myStockSist &" IS NULL THEN 0 ELSE " & myStockSist & " END AS STOCKSISTEMA, " & myStockFis 
		myCampo = myCampo & " ,CASE WHEN T3.VLUPESOS IS NULL THEN 0 ELSE T3.VLUPESOS END AS VLUPESOS, "
 	end if	
	strSQL = "	  	SELECT  A.idarticulo, "
	strSQL = strSQL & "     B.dsarticulo, " 
	strSQL = strSQL & "     C.cdinterno,  "
	strSQL = strSQL &		myCampo
	strSQL = strSQL & "     D.abreviatura, "
	strSQL = strSQL & "     E.CDCATEGORIA, "
	strSQL = strSQL & "     E.DSCATEGORIA "
	strSQL = strSQL & " FROM  TBLCSTKDETALLE A "
	strSQL = strSQL & "    INNER JOIN tblarticulos B "
	strSQL = strSQL & "  			ON A.idarticulo = B.idarticulo "	
	strSQL = strSQL & "    INNER JOIN TBLUNIDADES D "
	strSQL = strSQL & "  			ON B.idunidad = D.idunidad "
	strSQL = strSQL & "    INNER JOIN TBLARTCATEGORIAS E "
	strSQL = strSQL & "  			ON B.IDCATEGORIA = E.IDCATEGORIA "	
	if(not flagDatosCargados)then
	    strSQL = strSQL & "	   LEFT JOIN (SELECT T1.idarticulo, "
	    strSQL = strSQL & "                    T1.vlupesos  , "
	    strSQL = strSQL & "                    T2.ultimo	  "
	    strSQL = strSQL & "              FROM  tblarticulosprecios T1 "
	    strSQL = strSQL & "                    INNER JOIN  "
	    strSQL = strSQL & "                            (	SELECT  MAX(MMTOPRECIO) ULTIMO, IDARTICULO "
	    strSQL = strSQL & "						        	FROM    TBLARTICULOSPRECIOS "
	    strSQL = strSQL & "                                 WHERE   IDDIVISION = " & getDivisionAlmacen(pIdAlmacen)
	    strSQL = strSQL & "                              	GROUP BY IDARTICULO "
	    strSQL = strSQL & "                            ) T2 "			
	    strSQL = strSQL & "                      ON  T2.IDARTICULO = T1.IDARTICULO "    
	    strSQL = strSQL & "                      AND T1.MMTOPRECIO = T2.ULTIMO "
	    strSQL = strSQL & "                      AND T1.IDDIVISION = " & getDivisionAlmacen(pIdAlmacen)
	    strSQL = strSQL & "              ) T3 "
	    strSQL = strSQL & "    ON  T3.idarticulo = B.idarticulo "		
	end if
	strSQL = strSQL & "    LEFT JOIN tblarticulosdatos C "
	strSQL = strSQL & "  			ON B.idarticulo = C.idarticulo and C.idalmacen =" & pIdAlmacen
	strSQL = strSQL & " WHERE A.idcontrol = "& pIdControl &" and B.estado = " & ESTADO_ACTIVO
	strSQL = strSQL & " order by E.CDCATEGORIA ,B.idarticulo  "	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	Set leerDetallesCtSt = rs
End Function
'-----------------------------------------------------------------------------------------
' Funci�n:	leerCabeceraCtSt
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	27/02/2013 
' Objetivo:	
'			Leer los datos de la Tabla Cabecera de un determinado Control de Stock
' Parametros:
'			pIdControl  		[int] 	  ID CONTROL
' Devuelve:
'			Record Set
'--------------------------------------------------------------------------------------------
Function leerCabeceraCtSt(pIdControl)
	dim strSQL 
	strSQL = "SELECT * FROM TBLCSTKCABECERA WHERE IDCONTROL = " & pIdControl
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set leerCabeceraCtSt = rs
End Function

%>