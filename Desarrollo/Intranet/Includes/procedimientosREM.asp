<%


dim REM_id, REM_Fecha,REM_idProveedor, REM_cdProveedor, REM_dsProveedor
dim REM_usuario, REM_momento, REM_cdRemito
dim REM_idArticulo, REM_dsArticulo, REM_cantidad
dim rsArticulos, REM_idPIC, REM_Estado
dim REM_CantArticulos, REM_ArticuloActual, REM_idRemito
dim REM_hayCabecera, REM_abreviaturaUnidad, REM_idUnidad, REM_CantOriginal
dim REM_idObra, REM_idAlmacen, REM_nroRemito, REM_cdInterno, REM_Existencia, REM_Sobrante
dim arrArticulosConErrores
dim g_flagExistenciaArticulos 'sirve para controlar, si usuario ingreso al menos un articulo para este remito

arrArticulosConErrores = Array()
'Constante
Const ACCION_REM_COTIZAR = "cotizar"
Const ACCION_REM_APERTURA = "apertura"
Const ACCION_REM_RETIRARSE = "NO_COTIZA"

Const PREFIX_REM = "REM"
Const PREFIX_REM_X = "XEM"

'/* TIPOS DE REMITO */
Const CODIGO_REM_REMITO = "REM"
Const CODIGO_REM_ANULACION = "XEM"

'Inicializacion de datos clave para la accion de la pagina.
REM_idPIC = 0

'---------------------------------------------------------------------------------------------
Function initHeaderREM(p_idRemito)	
	Call clearHeaderREM()	
	if (isFormSubmit()) then
		Call initHeaderREMParams(p_idRemito)
	else 
		if (p_idRemito > 0) then 
			Call initHeaderREMDB(p_idRemito)			
		else		
			Call initHeaderREMNuevo()
		end if
	end if			 	
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderREMNuevo() 
	
	Dim articulos, strSQL, rs
	
	REM_hayCabecera = False					
	
	REM_idRemito = 0
	REM_Fecha = left(GF_VERFECHADATO(),10)	
	REM_idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
	REM_usuario = session("Usuario")
	REM_momento = session("MmtoSistema")
	REM_nroRemito = ""
	REM_cdRemito = GF_PARAMETROS7("cdREM", "", 6)
	REM_idPIC = GF_PARAMETROS7("ref", 0, 6)
	if (REM_idPIC > 0) then
		'Se paso un PIC
		Call initHeaderREMNuevoPIC()	
	end if
	
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderREMNuevoPIC() 
	strSQL="Select * from TBLCTZCABECERA where IDCOTIZACION=" & REM_idPIC	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		REM_idProveedor = rs("IDPROVEEDOR")
		REM_dsProveedor = getDescripcionProveedor(REM_idProveedor)
		REM_Fecha = GF_FN2DTE(Left(session("MmtoDato"), 8))
		REM_hayCabecera = True			
	else
		'El PIC es invalido
		REM_idPIC = 0
	end if			
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderREMParams(p_idRemito)
	dim strSQL, rs, km, kc
	REM_id = p_idRemito	
	REM_Fecha = GF_FN2DTE(Left(session("MmtoDato"), 8))
	REM_idProveedor = GF_PARAMETROS7("idProveedor",0,6)
	REM_dsProveedor = getDescripcionProveedor(REM_idProveedor)
	REM_nroRemito = GF_PARAMETROS7("nroRemito", 0,6)
	REM_idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)	
	REM_idPIC = GF_PARAMETROS7("ref", 0, 6)
	REM_cdRemito = GF_PARAMETROS7("cdREM", "", 6)
	REM_usuario = session("Usuario")
	REM_momento = session("MmtoSistema")
	REM_Estado = GF_PARAMETROS7("estado", 0, 6) 
	REM_hayCabecera = True
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderREMDB(p_idRemito)
	dim strSQL, rs, rs_proveedor, km, kc, tmp, pos
	REM_hayCabecera = False
	strSQL="select * from TBLREMCABECERA where IDREMITO=" & p_idRemito
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		idRemito = rs("IDREMITO")
		REM_idRemito = rs("IDREMITO")
		REM_Fecha = GF_FN2DTE(rs("FECHA"))
		REM_nroRemito = rs("NROREMITO")
		REM_idAlmacen = rs("IDALMACEN")
		REM_idProveedor = rs("IDPROVEEDOR")
		REM_dsProveedor = getDescripcionProveedor(REM_idProveedor)
		REM_cdRemito = rs("CDREMITO")
		REM_usuario = rs("CDUSUARIO")
		REM_momento = rs("MOMENTO")		
		REM_Estado = rs("ESTADO")
		REM_hayCabecera = True
	end if		
End Function
'---------------------------------------------------------------------------------------------
Function initArticulos()
	initArticulos = false	
	if (isFormSubmit()) then
		REM_ArticuloActual=0
		REM_CantArticulos = GF_PARAMETROS7("cantArticulos", 0, 6)
		initArticulos = true
	else
		if ((REM_idPIC > 0) and (REM_idRemito=0)) then
			initArticulos = initArticulosPIC()
		else						
			initArticulos = initArticulosDB()
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
function initArticulosDB()
	dim strSQL, rs, km, kc

	initArticulosDB = false
	if (REM_hayCabecera) then
		strSQL="select rem.* from TBLREMDETALLE rem "
		strSQL = strSQL & " inner join tblarticulos art on art.idarticulo = rem.idarticulo "
		strSQL = strSQL & " inner join tblartcategorias cat on cat.idcategoria = art.idcategoria and cat.tipocategoria <> '" & TIPO_CAT_SERVICIOS & "'"
		strSQL = strSQL & " where rem.IDREMITO=" & REM_idRemito
		call executeQueryDb(DBSITE_SQL_INTRA, rsArticulos, "OPEN", strSQL)
		if (not rsArticulos.eof) then
			initArticulosDB = true
		end if
	end if
end function
'---------------------------------------------------------------------------------------------
Function initArticulosPIC()
	Dim strSQL
	
	initArticulosPIC = false
	if (REM_hayCabecera) then
		'Se traen los articulos del PIC que aun tienen saldo pendiente de recepcion.
		strSQL="			Select PEDIDO.IDARTICULO, PEDIDO.CANTIDAD CANTIDADP, RECIBIDO.CANTIDAD CANTIDADR"		
		strSQL= strSQL & "	from (Select IDCOTIZACION, IDARTICULO, SUM(CANTIDAD) CANTIDAD from TBLCTZDETALLE group by IDCOTIZACION, IDARTICULO) PEDIDO"		
		strSQL= strSQL & "	left join  ("
		strSQL= strSQL & "	    Select RP.IDPIC, RP.IDARTICULO, sum(cantidad) CANTIDAD"
		strSQL= strSQL & "	    from TBLREMPIC RP where RP.IDPIC=" & REM_idPIC
		strSQL= strSQL & "	    group by RP.IDPIC, RP.IDARTICULO"
		strSQL= strSQL & "	    ) RECIBIDO on PEDIDO.IDCOTIZACION = RECIBIDO.IDPIC and PEDIDO.IDARTICULO=RECIBIDO.IDARTICULO"
		strSQL= strSQL & "  inner join tblarticulos art on art.idarticulo = PEDIDO.IDARTICULO "
		strSQL = strSQL & " inner join tblartcategorias cat on cat.idcategoria = art.idcategoria and cat.tipocategoria <> '" & TIPO_CAT_SERVICIOS & "'"
		strSQL= strSQL & "	where PEDIDO.IDCOTIZACION=" & REM_idPIC & " and (PEDIDO.CANTIDAD > RECIBIDO.CANTIDAD or RECIBIDO.CANTIDAD is null) and PEDIDO.CANTIDAD > 0 "
		call executeQueryDb(DBSITE_SQL_INTRA, rsArticulos, "OPEN", strSQL)
		if (not rsArticulos.eof) then
			initArticulosPIC = true
		end if
	end if		
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticulo()
	Call clearArticulo()
	if (isFormSubmit()) then	
		if (REM_cdRemito=CODIGO_REM_ANULACION) then
			readNextArticulo = readNextArticuloAnulacionParams()
		else	
			readNextArticulo = readNextArticuloParams()
		end if	
	else
		if (REM_idPIC > 0) then
			readNextArticulo = readNextArticuloPIC()
		else
			readNextArticulo = readNextArticuloDB()
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloParams()
	dim strSQL, rs, ret
	ret = false
	while ((REM_ArticuloActual < REM_CantArticulos) and (not ret))
		REM_idArticulo = GF_PARAMETROS7("item" & REM_ArticuloActual,"",6)				
		REM_cantidad = GF_PARAMETROS7("amount" & REM_ArticuloActual,3,6)	
		REM_Existencia = GF_PARAMETROS7("amount" & REM_ArticuloActual,3,6)		
		REM_Sobrante = GF_PARAMETROS7("amount_S" & REM_ArticuloActual,3,6)		
		if ((REM_idArticulo <> "") and (REM_cantidad > 0)) then
			g_flagExistenciaArticulos = true 'hay al menos un articulo ingresado
			Call readArticuloDatosAdicionales()			
			ret = true
		end if
		REM_ArticuloActual = REM_ArticuloActual + 1
	wend
	readNextArticuloParams = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloAnulacionParams()
	dim strSQL, rs, ret
	ret = false
	while ((REM_ArticuloActual < REM_CantArticulos) and (not ret))
		REM_idArticulo = GF_PARAMETROS7("item" & REM_ArticuloActual,"",6)				
		REM_cantidad = GF_PARAMETROS7("amount" & REM_ArticuloActual,3,6)	
		REM_Existencia = GF_PARAMETROS7("amount" & REM_ArticuloActual,3,6)		
		REM_Sobrante = GF_PARAMETROS7("amountS" & REM_ArticuloActual,3,6)		
		REM_CantOriginal = GF_PARAMETROS7("original" & REM_ArticuloActual,3,6)		 
			g_flagExistenciaArticulos = true 'hay al menos un articulo ingresado
			Call readArticuloDatosAdicionales()			
			ret = true
		REM_ArticuloActual = REM_ArticuloActual + 1
	wend
	readNextArticuloAnulacionParams = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloDB()
	dim strSQL, rs, km
	
	readNextArticuloDB = false	
	if (not rsArticulos.eof) then
		REM_idArticulo = rsArticulos("IDARTICULO")		
		Call readArticuloDatosAdicionales()
		REM_cantidad = rsArticulos("CANTIDAD")
		REM_CantOriginal = REM_cantidad
		REM_Existencia = rsArticulos("EXISTENCIA")
		REM_Sobrante = rsArticulos("SOBRANTE")
		rsArticulos.MoveNext()
		readNextArticuloDB = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloAnulacionDB()
	dim strSQL, rs, km
	readNextArticuloAnulacionDB = false	
	if (not rsArticulos.eof) then
		REM_idArticulo = rsArticulos("IDARTICULO")		
		Call readArticuloDatosAdicionales()
		REM_cantidad = rsArticulos("CANTIDAD")
		REM_CantOriginal = REM_cantidad
		REM_Existencia = rsArticulos("EXISTENCIA")
		REM_Sobrante = rsArticulos("SOBRANTE")
		rsArticulos.MoveNext()
		readNextArticuloAnulacionDB = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloPIC()
	readNextArticuloPIC = false
	if (not rsArticulos.eof) then		
		REM_idArticulo = rsArticulos("IDARTICULO")
		REM_cantidad = CDbl(rsArticulos("CANTIDADP"))
		if (rsArticulos("CANTIDADR") <> "") then REM_cantidad = REM_cantidad - CDbl(rsArticulos("CANTIDADR"))
		Call readArticuloDatosAdicionales()
		rsArticulos.MoveNext()
		readNextArticuloPIC = True
	end if
End Function
'---------------------------------------------------------------------------------------------
Function readArticuloDatosAdicionales()
		Dim strSQL, rs, conn
		strSQL="select A.*, B.CDINTERNO from TBLARTICULOS A left join TBLARTICULOSDATOS B on A.IDARTICULO=B.IDARTICULO and B.IDALMACEN=" & REM_idAlmacen & " where A.IDARTICULO=" & REM_idArticulo
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			REM_dsArticulo = rs("DSARTICULO")
			REM_abreviaturaUnidad = getAbreviaturaUnidad(rs("IDUNIDAD"))
			REM_cdInterno = rs("CDINTERNO")
		end if
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos del pedido de cotización.
Function controlarRemito()
	Dim tmp, cantProv, provs, nrmName, listaArticulos, dicArticulos, rs
	listaArticulos = "0"
	Set dicArticulos = Server.CreateObject("Scripting.Dictionary")
	Set provs = Server.CreateObject("Scripting.Dictionary")
	controlarRemito = false		
	Call initHeaderREM(idRemito)	
	REM_CantArticulos = GF_PARAMETROS7("cantArticulos", 0, 6)	
	if (controlarHeaderREM()) then
		'Se controlan los proveedores
		tmp = true		
		if (initArticulos()) then
			g_flagExistenciaArticulos = false
			while ((readNextArticulo()) and (tmp))
				tmp = controlarArticuloREM()
				if (tmp) then
					listaArticulos = listaArticulos & "," & REM_idArticulo										
					Call dicArticulos.Add(CLng(REM_idArticulo), REM_cantidad)
				end if
			wend						
			if not g_flagExistenciaArticulos then 
				setError(POCOS_ARTICULOS)				
				tmp = false
			end if
		end if					
		'Valido que las cantidades no excedan los articulos pedidos.		
		if (tmp) then tmp = controlarREMvsPIC(listaArticulos, dicArticulos)
				
		controlarRemito = tmp		
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos del pedido de cotización.
Function controlarRemitoAnulacion()
	
	Dim tmp, cantProv, provs, nrmName, listaArticulos, dicArticulos, rs, index
	
	listaArticulos = "0"
	Set dicArticulos = Server.CreateObject("Scripting.Dictionary")
	controlarRemitoAnulacion = false		
	Call initHeaderREM(idRemito)	
	REM_CantArticulos = GF_PARAMETROS7("cantArticulos", 0, 6)	
	tmp = true
	if (initArticulos()) then
		index = 0
		while ((readNextArticulo()) and (tmp))
			index = index + 1
			listaArticulos = listaArticulos & "," & REM_idArticulo
			Call dicArticulos.Add(CLng(REM_idArticulo), REM_cantidad)
		wend
	end if
	'Se valida que se puede quitar el stock de la base
	tmp = puedeQuitarStock(idRemito, dicArticulos)
	'Verifico si hay suficientes unidades que aun no se haya pagado.
	if (tmp) then tmp = hayUnidadesSinPago(idRemito)	
	controlarRemitoAnulacion = tmp
End Function
'---------------------------------------------------------------------------------------------
Function hayUnidadesSinPago(idRemito)
    Dim strSQL, rs, ret
    
'   Un remito solo se anula compelto, para eso se deben poder devolver todas sus unidades.
'   Cada Remito esta relacionado a varios PICs, se debe cumplir la condicion de que todos los PICs tengan unidades suficientes para permitir devolver.
'   La primera condicion es que las unidades recibidas para ese PIC sean mayores a la unidades ya facturadas. Recibido-Facturado = Saldo permitido a Devolver
'   La segunda condicion es que la cantidad que toma el remito de ese PIC (Relacionado) sea menor o igual al Saldo Permitido de Devolver. O sea 
'	Saldo permitido a Devolver >= Relacionado
'	Recibido-Facturado >= Relacionado
'	------------------------------------------
'	| (Recibido-Facturado)-Relacionado >= 0  |
'	------------------------------------------
'   Esta relacion debe cumplirse para todos los articulos del PIC

    strSQL= "Select MIN((REC.CANTIDAD-FAC.FACTURADO)-FAC.RELACIONADO) SALDO from " & _
            "(Select RP.IDPIC, RP.IDARTICULO, RP.IDAREA, RP.IDDETALLE, RP.CANTIDAD RELACIONADO, case when FACTURADO is Null then 0 else FACTURADO end FACTURADO " & _
            "from TBLREMPIC RP " & _
            "left join TBLCTZDETALLE DET on RP.IDPIC=DET.IDCOTIZACION and RP.IDARTICULO=DET.IDARTICULO and RP.IDAREA=DET.IDAREA and RP.IDDETALLE=DET.IDDETALLE " & _
            "where IDREMITO=" & idRemito & ") FAC inner join " & _
            "(Select  IDPIC, IDARTICULO, IDAREA, IDDETALLE, SUM(CANTIDAD) CANTIDAD " & _
            "from TBLREMPIC where IDPIC in (Select Distinct IDPIC from TBLREMPIC where IDREMITO=" & idRemito & ")" & _
            "group by IDPIC, IDARTICULO, IDAREA, IDDETALLE) REC " & _
            "on FAC.IDPIC=REC.IDPIC and FAC.IDARTICULO=REC.IDARTICULO and FAC.IDAREA=REC.IDAREA and FAC.IDDETALLE=REC.IDDETALLE"
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    ret = false
    if (not rs.eof) then
        if (CDbl(rs("SALDO")) >= 0) then ret = true
    end if
    if (not ret) then Call setError(REM_ANULA_ARTICULO_PAGO)
    hayUnidadesSinPago = ret
End Function
'---------------------------------------------------------------------------------------------
'Se controla que las cantidades ingresadas no superen lo pedido.
Function controlarREMvsPIC(pListaArticulos, pDicArticulos)
	Dim strSQL, rs, ret, index, saldo, idArticulo	
	if (REM_idPIC > 0) then		
		initArticulosPIC()
		Set rs = rsArticulos		
	else
		Set rs = obtenerArticulosPedidosNoRecibidos(getDivisionAlmacen(REM_idAlmacen), REM_idProveedor, pListaArticulos, "", "")
	end if
	ret = true
	index=0	
	while ((not rs.eof) and (ret))
		index=index+1			
		idArticulo = CLng(rs("IDARTICULO"))
		if (pDicArticulos.Exists(idArticulo)) then				
			saldo=0
			if (rs("CANTIDADP") <> "") then saldo = CDbl(rs("CANTIDADP"))
			if (rs("CANTIDADR") <> "") then saldo = saldo - CDbl(rs("CANTIDADR"))			
			if (round(pDicArticulos(idArticulo), 2) > round(saldo, 2)) then 
				ret = false
				addArticulosConErrores arrArticulosConErrores, rs("IDARTICULO")
			end if
		end if
		rs.MoveNext()
	wend
	if (index = 0) then ret =false		
	if (not ret) then setError(CANTIDAD_MENOR_SALDO)	
	controlarREMvsPIC = ret
End Function
'-----------------------------------------------------------------------------------------
function puedeQuitarStock(pIdRemito, pDicArticulos)
dim rs, oConn, strSQL, rtrn, index
rtrn = true
strSQL = "SELECT DET.EXISTENCIA EXISTENCIA_REM, DET.SOBRANTE SOBRANTE_REM, ART.IDARTICULO, ART.DSARTICULO, DAT.EXISTENCIA EXISTENCIA_ALMA, DAT.SOBRANTE SOBRANTE_ALMA FROM TBLREMCABECERA CAB INNER JOIN TBLREMDETALLE DET ON CAB.IDREMITO=DET.IDREMITO INNER JOIN TBLARTICULOSDATOS DAT ON DET.IDARTICULO=DAT.IDARTICULO AND CAB.IDALMACEN=DAT.IDALMACEN INNER JOIN TBLARTICULOS ART ON DET.IDARTICULO=ART.IDARTICULO WHERE CAB.IDREMITO= " & pIdRemito
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if rs.eof then
	rtrn = false
	call setError(ARTICULO_NO_EXISTE)
else
	while not rs.eof
			idArticulo = CLng(rs("IDARTICULO"))
			index = index + 1
			if (pDicArticulos.Exists(idArticulo)) then	
				myCantidad = pDicArticulos(idArticulo)
				'Response.Write "CAN(" & myCantidad & ")EXI(" & rs("EXISTENCIA_ALMA") & ")SOB(" & rs("SOBRANTE_ALMA") & ")"
				if (cdbl(myCantidad) > cdbl(rs("EXISTENCIA_ALMA"))+cdbl(rs("SOBRANTE_ALMA"))) then
					rtrn = false
					call setError(STOCK_ACTUAL_NO_CUBRE)
					addArticulosConErrores arrArticulosConErrores, rs("IDARTICULO")
				end if
			end if
		rs.movenext
	wend
end if
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
'Response.Write "index(" & index & ") pDicArticulos.count(" & pDicArticulos.count & ")"
if CLNG(index) <> CLNG(pDicArticulos.count) then
	rtrn = false
	call setError(ARTICULO_NO_EXISTE)
end if
puedeQuitarStock = rtrn
end function
'---------------------------------------------------------------------------------------------
'Controla los datos de la cabecera cargada.
Function controlarHeaderREM()
	Dim rs, strSQL
	controlarHeaderREM = false
	if (REM_idAlmacen <> 0) then
        if (REM_idProveedor <> 0) then
	        if (REM_nroRemito = 0) then
		        setError(REMITO_NO_EXISTE)
	        else
		        strSQL="select count(*) as CantRemitos from TBLREMCABECERA where NROREMITO=" & REM_nroRemito & " AND CDREMITO='" & REM_cdRemito & "' AND IDPROVEEDOR=" & REM_idProveedor & " and ESTADO<>" & ESTADO_BAJA
		        call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		        if (rs("CantRemitos") > 0) then
			        setError(NRO_REMITO_REPETIDO)
		        else
			        controlarHeaderREM = true
		        end if			
	        end if
        else
	        setError(PROVEEDOR_NO_EXISTE)
        end if
    else
        setError(ALMACEN_NO_EXISTE)
    end if
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloREM() 	
	Dim ret
	ret = false
	if CDbl(REM_cantidad) <= 0 then 
		setError(CANTIDAD_NO_EXISTE)
	else
		ret = true
	end if
	if (not ret) then addArticulosConErrores arrArticulosConErrores, REM_idArticulo
	controlarArticuloREM = ret
End Function
'---------------------------------------------------------------------------------------------
'Devuelve la cantidad de articulos que tiene un pedido
Function getCantidadArticulos(p_idRemito)
	Dim rs, strSQL, rtrn
	rtrn = 0
	strSQL="select count(*) as Cantidad from TBLREMDETALLE where IDREMITO=" & p_idRemito
	'Response.Write strsql
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		if not isnull(rs("Cantidad")) then rtrn = rs("Cantidad")
	end if
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getCantidadArticulos = rtrn
End Function
'---------------------------------------------------------------------------------------------
Function grabarFormulario() 
	REM_Estado = ESTADO_ACTIVO
	REM_id = grabarHeaderREM()
	Call grabarArticulosYStock()
	call ActualizarPrecios(REM_id, CODIGO_REM_REMITO)
	grabarFormulario = REM_id
End Function
'---------------------------------------------------------------------------------------------
Function grabarFormularioAnulacion(pIdRemitoAnterior) 
	call cambiarEstadoRemito(pIdRemitoAnterior, ESTADO_BAJA)
	REM_cdRemito = CODIGO_REM_ANULACION
	REM_Estado = ESTADO_ANULACION
	REM_id = grabarHeaderREMInsert()
	Call grabarArticulosYStockAnulacion(pIdRemitoAnterior)
	call ActualizarPrecios(REM_id, CODIGO_REM_ANULACION)
	grabarFormularioAnulacion = REM_id
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderREM()
	if (REM_id = 0) then		
		REM_id = grabarHeaderREMInsert()		
		grabarHeaderREM = REM_id
	end if		
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderREMInsert()
	Dim strSQL, rs, dte, idPedido
	strSQL= "Insert into TBLREMCABECERA(NROREMITO, IDALMACEN, FECHA, IDPROVEEDOR, CDUSUARIO, MOMENTO, ESTADO, CDREMITO) values(" 
	strSQL = strSQL & REM_nroRemito & ", " & REM_idAlmacen & ", " & GF_DTE2FN(REM_Fecha) & ", " & REM_idProveedor 
	strSQL = strSQL & ", '" & session("Usuario") & "', '" & session("MmtoSistema") & "', " & REM_Estado & ", '" & REM_cdRemito & "')"
	'Response.Write strSQL
	'response.end
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = "Select MAX(IDREMITO) IDREMITO from TBLREMCABECERA where IDALMACEN=" & REM_idAlmacen
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	REM_id = rs("IDREMITO")
	grabarHeaderREMInsert = REM_id
End Function
'---------------------------------------------------------------------------------------------
Function LeerArticulosPendientes(pListaArticulos, pIdPIC)
	
	Dim strSQL, rs, conn, strSQL1, strSQL2
	
	
	if (pIdPIC <> 0) then
		strSQL1 = " and RP.IDPIC = " & pIdPIC
		strSQL2 = " and PEDIDO.IDCOTIZACION = " & pIdPIC
	end if
	
	strSQL="			Select PEDIDO.IDARTICULO, PEDIDO.IDCOTIZACION IDPIC, PEDIDO.IDAREA, PEDIDO.IDDETALLE, PEDIDO.CANTIDAD CANTIDADP, SUM(RECIBIDO.CANTIDAD) CANTIDADR"
	strSQL= strSQL & "	from ("
	strSQL= strSQL & "		Select PIC.* from TBLCTZCABECERA PIC" 
	strSQL= strSQL & "		inner join 	(Select IDCOTIZACION IDPIC from TBLCTZCABECERA"
	strSQL= strSQL & "					EXCEPT" 
	strSQL= strSQL & "					Select IDPIC from TBLREMPIC where IDREMITO=0) NPIC on PIC.IDCOTIZACION=NPIC.IDPIC"
	strSQL= strSQL & "		) CAB inner join  "
	strSQL= strSQL & "			(select sum(cantidad) cantidad,idcotizacion,idarticulo, idarea, iddetalle from TBLCTZDETALLE group by idcotizacion,idarticulo, idarea, iddetalle )  "
	strSQL= strSQL & "      PEDIDO on CAB.IDCOTIZACION=PEDIDO.IDCOTIZACION"	
	strSQL= strSQL & "	left join  ("
	strSQL= strSQL & "	    Select RP.IDPIC, RP.IDARTICULO, RP.IDAREA, RP.IDDETALLE, sum(RP.CANTIDAD) CANTIDAD "
	strSQL= strSQL & "	    from TBLREMPIC RP"
	strSQL= strSQL & "	    inner join TBLREMCABECERA RC on RC.IDREMITO=RP.IDREMITO"
	strSQL= strSQL & "	    where RC.IDPROVEEDOR=" & REM_idProveedor & " and RP.IDARTICULO in (" & pListaArticulos & ")" & strSQL1
	strSQL= strSQL & "	    group by RP.IDPIC, RP.IDARTICULO, RP.IDAREA, RP.IDDETALLE"
	strSQL= strSQL & "	    ) RECIBIDO on PEDIDO.IDCOTIZACION = RECIBIDO.IDPIC and PEDIDO.IDARTICULO=RECIBIDO.IDARTICULO"
	strSQL= strSQL & "  INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO = PEDIDO.IDARTICULO"
	strSQL= strSQL & "  INNER JOIN TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA = CAT.IDCATEGORIA"
						'Se obtienen datos complementarios de las obras
	strSQL= strSQL & "  LEFT JOIN TBLDATOSOBRAS OBR on CAB.IDOBRA=OBR.IDOBRA"
	strSQL= strSQL & "	where (PEDIDO.CANTIDAD > RECIBIDO.CANTIDAD or RECIBIDO.CANTIDAD is null)"
	strSQL= strSQL & "		and CAB.IDDIVISION=" & getDivisionAlmacen(REM_idAlmacen) & " and CAT.TIPOCATEGORIA= '" & TIPO_CAT_BIENES & "' and CAB.IDPROVEEDOR=" & REM_idProveedor
	'strSQL= strSQL & "		and ART.BIENUSO <>'" & ES_BIEN_DE_USO & "'"	
	strSQL= strSQL & "		and (CAB.ESTADO='" & CTZ_FIRMADA & "' or CAB.ESTADO='" & CTZ_FACTURADA & "') and PEDIDO.IDARTICULO in (" & pListaArticulos & ")" & strSQL2
	strSQL= strSQL & "	group by PEDIDO.IDARTICULO, PEDIDO.IDCOTIZACION, PEDIDO.IDAREA, PEDIDO.IDDETALLE, PEDIDO.CANTIDAD"	
	strSQL= strSQL & "	order by PEDIDO.IDARTICULO, PEDIDO.IDCOTIZACION"	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set LeerArticulosPendientes = rs
	
End Function
'-------------------------------------------------------------------------------------------------
'Determina los PICs que seran relacionados a este remito. 
'Devuelve un diccionario con los datos del PIC a relacionar.
Function determinarPICsAfectados(pListaArticulos, pDicArticulos)
	Dim rs, sinAsignar, ret, saldo, idArticulo, index,datos,nuevosValores
	
	Set ret = Server.CreateObject("Scripting.Dictionary")
	'1º leo los pics con articulos pendientes (solo de los articulos del presesnte remito)
	Set rs = LeerArticulosPendientes(pListaArticulos, REM_idPIC)			
	if (not rs.eof) then					
		'2º asigno el nuevo remito a los pics que correspondan.
		index = 0
		while (not rs.eof)			
			idArticulo = CLng(rs("IDARTICULO"))
			sinAsignar = 0				
			'Si el articulo esta pedido tomo el saldo del PIC leido
			if (pDicArticulos.Exists(idArticulo)) then sinAsignar = pDicArticulos(idArticulo)			
			if (sinAsignar > 0) then
				'Hay unidades sin asignar a un PIC.
				saldo = CDbl(rs("CANTIDADP"))
				if (rs("CANTIDADR") <> "") then 
					saldo = CDbl(rs("CANTIDADP"))-CDbl(rs("CANTIDADR"))					
				end if
				pDicArticulos(idArticulo) = 0
				if (sinAsignar > saldo) then pDicArticulos(idArticulo) = sinAsignar - saldo					
				'if (not ret.Exists(cdbl(rs("IDPIC")))) then 												
				Call ret.Add(index, rs("IDPIC") & STRING_DELIMITER & idArticulo & STRING_DELIMITER & (sinAsignar - pDicArticulos(idArticulo)) & STRING_DELIMITER & rs("IDAREA") & STRING_DELIMITER & rs("IDDETALLE"))	
				index = index + 1
				'end if
			end if				
			rs.MoveNext
		wend
	end if
	Set determinarPICsAfectados = ret
	
End Function
'---------------------------------------------------------------------------------------------
'Se graba la relacion entre el remito y sus PICs origen.
Function grabarArticulosYStockAnulacion(pIdRemitoAnterior)
	Dim strSQL, rs, listaArticulos, dicArticulos

	listaArticulos = "0"
	Set dicArticulos = Server.CreateObject("Scripting.Dictionary")

	'Grabo los articulos
	Call initArticulos()
	while (readNextArticulo())
		'call loadCantidadesREM(pIdRemitoAnterior, REM_idArticulo, myExistencia, mySobrante)
		if (cDbl(REM_Existencia) > 0) or (cDbl(REM_Sobrante) > 0) then
			'Grabo el articulo de remito
			strSQL= "Insert into TBLREMDETALLE(IDREMITO, IDARTICULO, CANTIDAD, EXISTENCIA, SOBRANTE) values(" & REM_id & ", " & REM_idArticulo & ", " & REM_Cantidad & ",0,0)"
			call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			listaArticulos = listaArticulos & "," & REM_idArticulo
			Call dicArticulos.Add(CLng(REM_idArticulo), REM_Cantidad)
		end if
	wend
	'Se guarda la relacion con los PICs
	Call grabarRelacionREMPICAnulacion(pIdRemitoAnterior, listaArticulos, dicArticulos)
End Function
'---------------------------------------------------------------------------------------------
'Se graba la relacion entre el remito y sus PICs origen.
Function grabarArticulosYStock()
	Dim strSQL, rs, listaArticulos, dicArticulos
	
	listaArticulos = "0"
	Set dicArticulos = Server.CreateObject("Scripting.Dictionary")
	
	'Grabo los articulos
	Call initArticulos()
	while (readNextArticulo())		
		'Grabo el articulo de remito
		strSQL= "Insert into TBLREMDETALLE(IDREMITO, IDARTICULO, CANTIDAD, EXISTENCIA, SOBRANTE) values(" & REM_id & ", " & REM_idArticulo & ", " & REM_Cantidad & ", 0, 0)"
		'response.write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		listaArticulos = listaArticulos & "," & REM_idArticulo
		Call dicArticulos.Add(CLng(REM_idArticulo), CDbl(REM_Cantidad))
	wend
	'Se guarda la relacion con los PICs
	Call grabarRelacionREMPIC(listaArticulos, dicArticulos)
End Function
'---------------------------------------------------------------------------------------------
Function grabarRelacionREMPICAnulacion(pIdRemitoAnterior, pListaArticulos, pDicArticulos)
	dim strSQL, rs, dic, k, datos, rsIns, cantidad

	strSQL = "SELECT * FROM TBLREMPIC WHERE IDREMITO=" & pIdRemitoAnterior 
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
			cantidad = -CDbl(rs("CANTIDAD"))
			strSQL="Insert into TBLREMPIC values(" & REM_id & ", " & Clng(rs("IDPIC")) & ", " & clng(rs("IDARTICULO")) & ", " & cantidad & ", " & rs("IDAREA") & ", " & rs("IDDETALLE") & ")"
			'Response.Write strSQL & "<br>"
			call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
			REM_idPIC = rs("IDPIC")
			REM_idArticulo = rs("IDARTICULO")
			REM_Cantidad = cantidad
			actualizarStockREM()	
			cantidad = 0		
		rs.movenext
	wend		
End Function
'---------------------------------------------------------------------------------------------
Function grabarRelacionREMPIC(pListaArticulos, pDicArticulos)
	dim strSQL, rs, dic, k, datos

	Set dic = determinarPICsAfectados(pListaArticulos, pDicArticulos)
	'Se guardan los PICs afectados	
	for each k in dic.Keys
		datos= split(dic(k), STRING_DELIMITER)
		strSQL="Insert into TBLREMPIC values(" & REM_id & ", " & datos(0) & ", " & datos(1) & ", " & datos(2) & ", " & datos(3) & ", " & datos(4) & ")"
		'Response.Write strSQL & "<br>"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		REM_idPIC = datos(0)
		REM_idArticulo = datos(1)
		REM_Cantidad = datos(2)		
		actualizarStockREM()
	next
	'Response.end
End Function
'---------------------------------------------------------------------------------------------
Function actualizarStockREM()
	Dim strSQL, rs, existencia, sobrante
	
	strSQL= "select * from  TBLARTICULOS ART left join tblarticulosdatos ARTD on ART.IDARTICULO=ARTD.IDARTICULO and ARTD.idalmacen = " & REM_idAlmacen & " where ART.idarticulo = " & REM_idArticulo
	'response.write strSQL & "<br>"			
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		existencia=0
		sobrante=0
		'Si el PIC asociado tiene obra, el stock es de producto sobrante.
		strSQL = "Select * from TBLCTZCABECERA where IDCOTIZACION=" & REM_idPIC & " and IDOBRA=0"
		call executeQueryDb(DBSITE_SQL_INTRA, rsP, "OPEN", strSQL)
		if ((rs("BIENUSO") = ES_BIEN_DE_USO) or (rsP.eof)) then
			sobrante= REM_Cantidad
		else
			existencia= REM_Cantidad
		end if
		if (isNull(rs("IDALMACEN"))) then
			strSQL= "Insert into tblarticulosdatos(IDARTICULO, IDALMACEN, EXISTENCIA, SOBRANTE, CDUSUARIO, MOMENTO) values(" & REM_idArticulo & ", " & REM_idAlmacen & ", " & existencia & ", " & sobrante & ", '" & session("Usuario") & "', '" & session("MmtoSistema") & "')"
		else
			'Actualizo stock de articulo que ya existe
			strSQL= "update tblarticulosdatos set existencia = (existencia + " & existencia & "), sobrante=(sobrante + " & sobrante & "), cdusuario = '" & session("Usuario") & "', momento = '" & session("MmtoSistema") & "' where idalmacen = " & REM_idAlmacen & " and idarticulo = " & REM_idArticulo
		end if
		'response.write strSql
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		'Actualizo el remito agregando la especie de la cantidad (Existencia o sobrante)
		'Siempre para un articulo y un PIC solo se asigna o existencia o sobrante.
		strSQL = "Update TBLREMDETALLE set EXISTENCIA=EXISTENCIA + " & existencia & ", SOBRANTE=SOBRANTE + " & sobrante & " where IDREMITO=" & REM_id & " and idarticulo = " & REM_idArticulo
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if
End Function
'---------------------------------------------------------------------------------------------
function loadCantidadesREM(pIdRemitoAnterior, pIdArticulo, byref pExistencia, byref pSobrante)
dim strSQL, rs, oConn
	strSQL = "Select * from TBLREMCABECERA CAB INNER JOIN TBLREMDETALLE DET ON CAB.IDREMITO=DET.IDREMITO WHERE CAB.IDREMITO=" & pIdRemitoAnterior & " AND DET.IDARTICULO=" & pIdArticulo 
	'Response.Write "<br>loadCantidadesREM(" & strSQL & ")"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		pExistencia = rs("EXISTENCIA") 
		pSobrante	= rs("SOBRANTE") 
	end if
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
end function
'---------------------------------------------------------------------------------------------
'Borra todas las variables del HeaderREM
Function clearHeaderREM()
	REM = 0
	REM_idRemito = 0
	REM_idPIC = 0
	REM_nroRemito = 0
	REM_idAlmacen = 0
	REM_Fecha = ""
	REM_idProveedor = ""
	REM_dsProveedor = ""
	REM_usuario = ""
	REM_momento = ""
	REM_hayCabecera = false
End function
'---------------------------------------------------------------------------------------------
Function clearArticulo()
	REM_idArticulo = 0
	REM_dsArticulo = ""
	REM_cantidad = 0
	REM_idUnidad=0
	REM_abreviaturaUnidad=""
End Function
'---------------------------------------------------------------------------------------------
sub addArticulosConErrores (ByRef p_arrArticulosConErrores, ByRef p_idArticulo)
dim iNewUBound
	iNewUBound = UBound(p_arrArticulosConErrores) + 1
	redim preserve p_arrArticulosConErrores(iNewUBound)
	p_arrArticulosConErrores(iNewUBound) = p_idArticulo
end sub
'---------------------------------------------------------------------------------------------
sub cambiarEstadoRemito(pIdRemito, pEstado)
dim rs, oConn, strSQL
strSQL = "UPDATE TBLREMCABECERA SET ESTADO=" & pEstado & " where IDREMITO=" & pIdRemito
call executeQueryDb(DBSITE_SQL_INTRA, rsP, "EXEC", strSQL)
end sub
'---------------------------------------------------------------------------------------------
%>