<%
dim VS, VS_idVale, VS_FechaSolicitud, VS_FechaRequerido, VS_cdSolicitante, VS_dsSolicitante
dim VS_usuario, VS_momento, VS_PartidaPendiente, VS_estado, VS_nrVale
dim VS_CantArticulos, VS_ArticuloActual, VS_idPedido, VS_existencia, VS_sobrante
dim VS_hayCabecera, VS_abreviaturaUnidad, VS_idUnidad, VS_idBudgetArea, VS_idBudgetDetalle
dim VS_idObra, VS_idAlmacen, VS_idAlmacenDest, VS_Fecha, VS_cdVale, VS_saldo, VS_cumplido, VS_cantidad, VS_idArticulo, VS_dsArticulo, VS_Aju_Entregados, VS_Aju_Pedidos
dim arrArticulosConErrores, VS_cdInterno, VS_comentario, VS_idSector
dim gDicArticulos

arrArticulosConErrores = Array()

Const PREFIX_VS = "VS"

'/*Codigos DE UN VS */
Const CODIGO_VS_SALIDA = "VMS"
Const CODIGO_VS_SALIDA_X = "XMS"
Const CODIGO_VS_PRESTAMO = "VMP"
Const CODIGO_VS_PRESTAMO_X = "XMP"
Const CODIGO_VS_DEVOLUCION = "VMD"
Const CODIGO_VS_DEVOLUCION_X = "XMD"
Const CODIGO_VS_TRANSFERENCIA = "VMT"
Const CODIGO_VS_TRANSFERENCIA_X = "XMT"
Const CODIGO_VS_RECEPCION = "VMR"
Const CODIGO_VS_RECEPCION_X = "XMR"
Const CODIGO_VS_ENTRADA = "VME"
Const CODIGO_VS_ENTRADA_X = "XME"
Const CODIGO_VS_AJUSTE_VALE = "AJU"
Const CODIGO_VS_AJUSTE_VALE_X = "XJU"
Const CODIGO_VS_AJUSTE_STOCK = "AJS"
Const CODIGO_VS_AJUSTE_STOCK_X = "XJS"
Const CODIGO_VS_AJUSTE_PEDIDO = "AJP"
Const CODIGO_VS_AJUSTE_PEDIDO_X = "XJP"
Const CODIGO_VS_AJUSTE_TRANSFERENCIA = "AJT"
Const CODIGO_VS_AJUSTE_TRANSFERENCIA_X = "XJT"
const CODIGO_PM = "PM"
Const CODIGO_VS_FIX = "SYS"
Const CODIGO_VS_RECLASIFICACION_STOCK = "VRS"
Const CODIGO_VS_RECLASIFICACION_STOCK_X = "XRS"

Const LEYENDA_VS_SALIDA = "Vale de Salida"
Const LEYENDA_VS_SALIDA_X = "Anulacion Vale de Salida"
Const LEYENDA_VS_PRESTAMO = "Vale de Prestamo"
Const LEYENDA_VS_PRESTAMO_X = "Anulacion Vale de Prestamo"
Const LEYENDA_VS_DEVOLUCION = "Vale de Devolucion"
Const LEYENDA_VS_DEVOLUCION_X = "Anulacion Vale de Devolucion"
Const LEYENDA_VS_TRANSFERENCIA = "Vale de Transferencia"
Const LEYENDA_VS_TRANSFERENCIA_X = "Anulacion Vale de Transferencia"
Const LEYENDA_VS_RECEPCION = "Vale de Recepcion"
Const LEYENDA_VS_RECEPCION_X = "Anulacion Vale de Recepcion"
Const LEYENDA_VS_ENTRADA = "Vale de Entrada"
Const LEYENDA_VS_ENTRADA_X = "Anulacion Vale de Entrada"
Const LEYENDA_VS_AJUSTE_VALE = "Vale de Ajuste de movimientos"
Const LEYENDA_VS_AJUSTE_VALE_X = "Anulacion Vale de Ajuste de movimientos"
Const LEYENDA_VS_AJUSTE_STOCK = "Vale de Ajuste de Stock"
Const LEYENDA_VS_AJUSTE_STOCK_X = "Anulacion Vale de Ajuste de Stock"
Const LEYENDA_VS_AJUSTE_PEDIDO = "Vale de Ajuste Pedido"
Const LEYENDA_VS_AJUSTE_PEDIDO_X = "Anulacion Vale de Ajuste Pedido"
Const LEYENDA_VS_AJUSTE_TRANSFERENCIA = "Vale de Ajuste de Transferencia"
Const LEYENDA_VS_AJUSTE_TRANSFERENCIA_X = "Anulacion Vale de Ajuste de Transferencia"
Const LEYENDA_PM = "Pedido de Materiales"
Const LEYENDA_VS_FIX = "Vale Correccion del Sistema"
Const LEYENDA_VS_RECLASIFICACION_STOCK = "Vale de Reclasificacion de Stock"
Const LEYENDA_VS_RECLASIFICACION_STOCK_X = "Anulacion Vale de Reclasificacion de Stock"
Const VS_FIRMA_RESPONSABLE = 0
Const VS_FIRMA_GERENTE     = 1
Const VS_FIRMA_COORD_AUDIT = 2
Const VS_FIRMA_DIRECTOR    = 3 'Para ajustes de stock que necesiten ser firmados por Dirección
'Para el registro de las firmas de ajuste de stock
Const VS_NO_USER = "XXX"
Const VS_AUDIT_USER = "AIS"
Const VS_PORT_SUPERVISOR_USER = "RPS"
Const VS_PORT_GERENTE_USER = "GPS"
'Codigo para identificar la norma de auditoria que tiene el importe del vale de ajuste para que sea autorizado por el director
Const NORMA_VS_AJUSTE_AUTORIZADO = "VLAJSAUTH"
'---------------------------------------------------------------------------------------------
Function initHeaderVale(pIdVale)
	Call clearHeaderVale()
	if (isFormSubmit()) then
		Call initHeaderValeParams(pIdVale)
	else
		if (pIdVale=0) then 
			Call initHeaderValeNuevo()
		else
			Call initHeaderValeDB(pIdVale)
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderValeNuevo() 
	VS_cdVale = GF_PARAMETROS7("cdVale","",6)	
	VS_FechaSolicitud = left(GF_VERFECHADATO(),10)	
	VS_FechaRequerido = left(GF_VERFECHADATO(),10)
	VS_cdSolicitante = ""
	VS_idAlmacen = 0
	VS_idAlmacenDest = 0
	VS_idObra = 0
	VS_idSector = 0
	VS_idBudgetArea = 0
	VS_idBudgetDetalle = 0
	VS_nrVale = ""
	VS_hayCabecera = false
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderValeParams(pIdVale)
	dim strSQL, rs, km, kc
	VS = pIdVale
	VS_cdVale = GF_PARAMETROS7("cdVale","",6)
	VS_FechaSolicitud = GF_PARAMETROS7("issuedate","",6)
	VS_FechaRequerido = GF_PARAMETROS7("closingdate","",6)
	VS_cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)
	VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
	VS_idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
	VS_idAlmacenDest = GF_PARAMETROS7("idAlmacenDest",0,6)
	VS_idObra = GF_PARAMETROS7("idObra",0,6)
	VS_idSector = GF_PARAMETROS7("idSector",0,6)
	VS_idBudgetArea = GF_PARAMETROS7("idBudgetArea",0,6)
	VS_idBudgetDetalle = GF_PARAMETROS7("idBudgetDetalle",0,6)
	VS_estado = GF_PARAMETROS7("estado",0,6)
	VS_usuario = session("Usuario")
	VS_momento = session("MmtoSistema")
	VS_nrVale = ""
	VS_hayCabecera = True	
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderValeDB(pIdVale)
	dim strSQL, rs, km, kc, tmp, pos
	
	VS_hayCabecera = False
	strSQL="select * from TBLVALESCABECERA where IDVALE=" & pIdVale
 	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (not rs.eof) then
		VS = rs("IDVALE")
		vs_idVale = rs("IDVALE")		
		vs_cdVale = rs("CDVALE")
		VS_FechaSolicitud = GF_FN2DTE(rs("FECHA"))
		VS_FechaRequerido = GF_FN2DTE(rs("FECHA"))
		VS_cdSolicitante = rs("CDSOLICITANTE")
		VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
		VS_idAlmacen = rs("IDALMACEN")
		VS_idObra = rs("IDOBRA")
		if (VS_idObra = "") then VS_idObra = 0		
		VS_idSector = rs("IDSECTOR")
		if (VS_idSector = "") then VS_idSector = 0		
		VS_idBudgetArea = rs("IDBUDGETAREA")
		if (VS_idBudgetArea = "") then VS_idBudgetArea = 0	
		VS_idBudgetDetalle = rs("IDBUDGETDETALLE")
		if (VS_idBudgetDetalle = "") then VS_idBudgetDetalle = 0	
		VS_usuario = rs("CDUSUARIO")
		VS_momento = rs("MOMENTO")
		VS_PartidaPendiente = rs("PARTIDAPENDIENTE")
		VS_estado = rs("ESTADO")
		VS_nrVale = Trim(rs("NRVALE"))
		VS_comentario = getComentarioVale(pIdVale)
		VS_hayCabecera = True
	end if		
End Function
'---------------------------------------------------------------------------------------------
Function initArticulosVale(pIdVale)
	VS_CantArticulos = GF_PARAMETROS7("cantArticulos",0,6)
	VS_ArticuloActual=0
	initArticulosVale = true
	if (pIdVale > 0) then
		initArticulosVale = initArticulosValeDB(pIdVale)
	end if
End Function
'---------------------------------------------------------------------------------------------
function initArticulosValeDB(pIdVale)
	dim strSQL, rs, km, kc
	VS_CantArticulos = GF_PARAMETROS7("cantArticulos",0,6)	
	initArticulosValeDB = false
	if (VS_hayCabecera) then
		strSQL="select * from TBLVALESDETALLE where IDVALE=" & pIdVale
		Call executeQueryDB(DBSITE_SQL_INTRA, rsArticulos, "EXEC", strSQL)		
		if (not rsArticulos.eof) then
			initArticulosValeDB = true
		end if
	end if
end function
'---------------------------------------------------------------------------------------------
Function readNextArticuloVale(pIdVale)
dim ret

	Call clearArticulo()
	if (isFormSubmit()) then
		ret = readNextArticuloValeParams()
	else 
		if (pIdVale > 0) then
			ret = readNextArticuloValeDB()
		end if
	end if		
	readNextArticuloVale = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloValeDB()
	dim strSQL, rs, km
	
	readNextArticuloValeDB = false	
	if (not rsArticulos.eof) then
		VS_idArticulo = rsArticulos("IDARTICULO")
		strSQL="select A.*, B.CDINTERNO from TBLARTICULOS A left join TBLARTICULOSDATOS B on A.IDARTICULO=B.IDARTICULO and B.IDALMACEN=" & VS_idAlmacen & " where A.IDARTICULO=" & VS_idArticulo
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
		if (not rs.eof) then
			VS_dsArticulo = rs("DSARTICULO")
			VS_idUnidad = rs("IDUNIDAD")
			VS_abreviaturaUnidad = getAbreviaturaUnidad(rs("IDUNIDAD"))
			VS_cdInterno = rs("CDINTERNO")
		end if		
		VS_cantidad = rsArticulos("CANTIDAD")
		VS_existencia = rsArticulos("EXISTENCIA")
		VS_sobrante = rsArticulos("SOBRANTE")
		VS_saldo = GF_PARAMETROS7("saldo" & VS_ArticuloActual,0,6)	
		VS_cumplido = GF_PARAMETROS7("cumplido" & VS_ArticuloActual,0,6)		
		rsArticulos.MoveNext()
		readNextArticuloValeDB = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function readNextArticuloValeParams()
	dim strSQL, rs, ret
	ret = false		
	VS_idArticulo = GF_PARAMETROS7("item" & VS_ArticuloActual, "",6)	
	if (VS_idArticulo <> "") then
		strSQL="select A.*, B.CDINTERNO from TBLARTICULOS A left join TBLARTICULOSDATOS B on A.IDARTICULO=B.IDARTICULO and B.IDALMACEN=" & VS_idAlmacen & " where A.IDARTICULO=" & VS_idArticulo
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
		if (not rs.eof) then
			VS_dsArticulo = rs("DSARTICULO")
			VS_idArticulo = rs("IDARTICULO")
			VS_idUnidad = rs("IDUNIDAD")
			VS_abreviaturaUnidad = getAbreviaturaUnidad(rs("IDUNIDAD"))
			VS_cdInterno = rs("CDINTERNO")
		end if				
		VS_cantidad = GF_PARAMETROS7("amount" & VS_ArticuloActual,3,6)
		VS_saldo = GF_PARAMETROS7("saldo" & VS_ArticuloActual,3,6)	
		VS_cumplido = GF_PARAMETROS7("cumplido" & VS_ArticuloActual,3,6)	
		VS_existencia = GF_PARAMETROS7("existencia" & VS_ArticuloActual,3,6)	
		VS_sobrante = GF_PARAMETROS7("sobrante" & VS_ArticuloActual,3,6)			
		ret = true
	end if
	VS_ArticuloActual = VS_ArticuloActual + 1
	readNextArticuloValeParams = ret
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de Vale.
Function controlarVale(pIdVale)	
	Dim tmp, cantProv, index
	Dim artDic,cantArt
	Set artDic = Server.CreateObject("Scripting.Dictionary")
	
	Set gDicArticulos = Server.CreateObject("Scripting.Dictionary") 
	controlarVale = false		
	if (controlarHeaderVale()) then
		tmp = true
		index=0
		cantArt = 0
		while ((readNextArticuloVale(pIdVale)) and (tmp))			
			index=1
 			if VS_cdVale = CODIGO_VS_AJUSTE_STOCK then
				tmp = controlarArticuloValeAjusteStock()
			elseif VS_cdVale = CODIGO_VS_AJUSTE_PEDIDO then
				tmp = controlarArticuloValeAjustePedido()
			elseif VS_cdVale = CODIGO_VS_AJUSTE_VALE then
				tmp = controlarArticuloValeAjusteVale()
			elseif VS_cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA then
				tmp = controlarArticuloValeAjusteTransf()
			elseif VS_cdVale = CODIGO_VS_RECLASIFICACION_STOCK then
				cantArt = cantArt + 1				
				tmp = controlarArticuloVRS()
				'controlo que no haya articulos duplicados
				if (not artDic.Exists(VS_idArticulo)) then
					Call artDic.Add(VS_idArticulo,VS_idArticulo)
				else
					Call setError(ARTICULO_DUPLICADO)
					tmp = false
				end if 
			else
				tmp = controlarArticuloVale()
			end if	
		wend		
		if (index = 0) then 
			Call setError(POCOS_ARTICULOS)	
			tmp = false	'No se cargaron articulos
		end if
		'controla que haya 2 articulos en la reclasificacion de stock
		if (cantArt < 2 and VS_cdVale = CODIGO_VS_RECLASIFICACION_STOCK and tmp = true) then
			Call setError(FALTA_ARTICULO)
			tmp = false
		end if
		controlarVale = tmp
	end if
	Set gDicArticulos = nothing
	VS_ArticuloActual = 0	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de la cabecera cargada.
Function controlarHeaderVale()
	dim idDivisionObra, idDivisionVS
	controlarHeaderVale = false
		if (VS_cdSolicitante = "") then call setError (SOLICITANTE_NO_EXISTE)
		if (VS_idAlmacen = 0) then call setError (ALMACEN_NO_EXISTE)
		'control de division de la partida y del vale
		if (VS_idObra > 0) then		
			idDivisionVS = getDivisionAlmacen(VS_idAlmacen)
			Call getDivisionObraFull(VS_idObra, idDivisionObra, "")
			if (idDivisionVS <> idDivisionObra) then setError(DIVISION_PM_VS_DIFF_OBRA)
		end if		
		'Solo un pañolero puede grabar un vale en su almacen.
		if (VS_cdVale = CODIGO_VS_RECEPCION) then
		    if ((isAuditorAL(VS_idAlmacenDest)) or (isSolicitanteAL(VS_idAlmacenDest))) then call setError (ALMACEN_NO_GUARDAR)
		else
		    if ((isAuditorAL(VS_idAlmacen)) or (isSolicitanteAL(VS_idAlmacen))) then call setError (ALMACEN_NO_GUARDAR)
		end if
		if (trim(VS_comentario) = "") then
			'El comentario en los ajustes es requerido
			if (VS_cdVale = CODIGO_VS_AJUSTE_VALE or VS_cdVale = CODIGO_VS_AJUSTE_PEDIDO or VS_cdVale = CODIGO_VS_AJUSTE_STOCK or VS_cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA) then 
				call setError (COMENTARIO_REQUERIDO)
			end if
		end if				
		'Control por tipo de vale
		select case ucase((VS_cdVale))
			case CODIGO_VS_TRANSFERENCIA, CODIGO_VS_RECEPCION
				if (VS_idAlmacenDest = 0) then call setError (ALMACENDEST_NO_EXISTE)
				if (VS_idAlmacenDest = VS_idAlmacen) then call setError (ALMACENDEST_IGUAL_ORIGEN)
			case CODIGO_VS_AJUSTE_PEDIDO, CODIGO_VS_AJUSTE_VALE, CODIGO_VS_AJUSTE_TRANSFERENCIA
				if (idPMReferencia = 0) then call setError (PM_REQUERIDO) 
			case CODIGO_VS_AJUSTE_STOCK
				'Nada por ahora
			case CODIGO_VS_ENTRADA 
				'Nada por ahora
			case CODIGO_VS_RECLASIFICACION_STOCK
				'Nada por ahora
			case CODIGO_PM
				'Si no es trabsferencia, debe especificar destino de los consumos (PArtida o sector)
				if (VS_idAlmacenDest = 0) then
					'Si no hay obra debe si o si debe haber sector
					if ((VS_idObra = 0) or (VS_idBudgetArea = 0) or (VS_idBudgetDetalle = 0)) then	
						if (not validarSector(VS_idSector)) then												
							call setError(FALTA_ASIGNAR_OBRA_SECTOR)
						end if
					end if
				else
					'Es pedido de transferencia, debe ser a otro almacen.
					if (VS_idAlmacenDest = VS_idAlmacen) then call setError (ALMACENDEST_IGUAL_ORIGEN)
				end if				
			case else
				'Si no hay obra debe si o si debe haber sector
				if ((VS_idObra = 0) or (VS_idBudgetArea = 0) or (VS_idBudgetDetalle = 0)) then	
					if (not validarSector(VS_idSector)) then												
						call setError(FALTA_ASIGNAR_OBRA_SECTOR)
					end if
				end if
		end select		
	
if not hayError() then controlarHeaderVale = true
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloVale() 	
	Dim rsStockArticulo, strSQL, cantIngresada
	Dim stockDisponible
	
	controlarArticuloVale = false
	'Controlo si el articulo existe
	if (VS_saldo > 0) then
		if (not controlarArticulo(VS_idArticulo)) then Call setError(ARTICULO_NO_EXISTE)	
	end if
	if (not hayError() ) then
		if (controlarRepetido(VS_idArticulo)) then
			flagValePP = llevaPartidaVale(VS_cdVale)
			flagPP = llevaPartida(VS_idArticulo)
			if (VS_idBudgetArea=0 and flagPP and flagValePP) then
				setError(ARTICULO_REQUIERE_PP)	
			else	
				if (VS_idSector=0 and not flagPP and flagValePP) then 
					setError(ARTICULO_REQUIERE_SEC)	
				else
					if ((VS_cdVale <> CODIGO_PM) and (CDBL(VS_saldo) > 0)) then	
						'Controlo el Stock del Articulo.
						strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & vs_idArticulo & " and IDALMACEN=" & VS_idAlmacen
						Call executeQueryDB(DBSITE_SQL_INTRA, rsStockArticulo, "OPEN", strSQL)		
						if not rsStockArticulo.eof then	
							'Control de stock suficiente
							Select case(VS_cdVale)
								'case CODIGO_VS_TRANSFERENCIA
								'	stockDisponible = CDbl(rsStockArticulo("EXISTENCIA")) 
								'	if (getDivisionAlmacen(VS_idAlmacen) = getDivisionAlmacen(VS_idAlmacenDest)) then
								'		stockDisponible = stockDisponible + CDbl(rsStockArticulo("SOBRANTE"))
								'	end if			
								'	if (round(CDbl(VS_saldo), 3) > stockDisponible) then setError(STOCK_EXISTENCIA_INSUFICIENTE)
								case CODIGO_VS_DEVOLUCION, CODIGO_VS_RECEPCION
									'No se controla el stock.
								case else							
									if (round(CDbl(VS_saldo), 3) > round(CDbl(rsStockArticulo("EXISTENCIA")) + CDbl(rsStockArticulo("SOBRANTE")), 3)) then 
										setError(STOCK_INSUFICIENTE)					
									end if		
							end select
										
							'Se controla que la cantidad este en el rango permitido.
							select case(VS_cdVale)
								case CODIGO_VS_PRESTAMO, CODIGO_VS_SALIDA, CODIGO_VS_TRANSFERENCIA
									if (VS_saldo > (VS_cantidad - VS_cumplido)) then
										setError(APRESTAR_MAY_PEDIDOS)	
									end if
								case CODIGO_VS_DEVOLUCION
									cantIngresada = VS_saldo
									devueltos = getCantidadDevuelta(idPMReferencia, VS_idArticulo)
									VS_saldo = VS_cumplido - CDbl(devueltos)					
									if (cantIngresada > VS_saldo) then
										setError(ADEVOLVER_MAY_PRESTADOS)	
									end if
								case CODIGO_VS_RECEPCION 
									devueltos = getCantidadRecibida(idPMReferencia, VS_idArticulo)
									VS_cumplido = CDbl(VS_cantidad) - CDbl(VS_saldo) 
									VS_saldo = VS_cumplido - CDbl(devueltos)
									if ((VS_saldo > VS_cumplido) or (VS_saldo<0)) then
										setError(ARECIBIR_MAY_TRANSFERIDOS)
									end if
							end select	
						else
							Call setError(STOCK_INSUFICIENTE)			
						end if			
					end if
				end if	
			end if	
		else
			setError(DETALLE_DUPLICADO)			
		end if
	else
		Call setError(ARTICULO_NO_EXISTE)			
	end if	
	if not hayError then 
		controlarArticuloVale = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function llevaPartida(pIdArticulo)
dim strSQL, rtrn
rtrn = false
strSQL = "SELECT * FROM TBLARTICULOS ART INNER JOIN TBLARTCATEGORIAS CAT " & _
		 " ON ART.IDCATEGORIA = CAT.IDCATEGORIA " & _
		 " WHERE ART.IDARTICULO=" & pIdArticulo & " AND CAT.ESMANTENIMIENTO='" & TIPO_AFIRMACION & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then rtrn = true
llevaPartida = rtrn
End Function
'---------------------------------------------------------------------------------------------
Function llevaPartidaVale(pCdVale)
dim strSQL, rtrn
rtrn = false
strSQL = "SELECT * FROM TBLVALESMAESTRO " & _
		 " WHERE ASIGNACION='" & TIPO_AFIRMACION & "' AND CDTIPO='" & pCdVale & "'"
		 'Response.Write strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if (not rs.eof) then 
	'Salvo que sea una transferencia, debe llevar partida. (VMT, VMR o PM)
	if (VS_idAlmacenDest = 0) then rtrn = true
end if
llevaPartidaVale = rtrn
End Function
'---------------------------------------------------------------------------------------------
Function controlarRepetido(pIdArticulo)
	Dim clave
	
	clave= "A" & pIdArticulo
	if (gDicArticulos.Exists(clave)) then
		controlarRepetido = false
	else
		gDicArticulos.Add clave, 1
		controlarRepetido = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloValeAjusteVale() 	
	Dim rsStockArticulo, strSQL
	controlarArticuloValeAjusteVale = false
	'Controlo si el articulo existe	
	Call controlarArticulo(VS_idArticulo)
	if (VS_saldo < 0) then Call setError(CANTIDAD_NO_NEGATIVA)			
	if ((VS_cantidad - VS_cumplido) < VS_saldo) then Call setError(NUEVA_MAY_NO_DEVUELTO)
	if not hayError then 
		controlarArticuloValeAjusteVale = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloValeAjusteStock() 	
	Dim rsStockArticulo, strSQL
	
	
	controlarArticuloValeAjusteStock = false
	'Controlo si el articulo existe	
	Call controlarArticulo(VS_idArticulo)
	if (VS_saldo < 0) then setError(CANTIDAD_NO_NEGATIVA)			
	if not hayError() then 
		controlarArticuloValeAjusteStock = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloValeAjustePedido() 	
	Dim rsStockArticulo, strSQL
	controlarArticuloValeAjustePedido = false
	'Controlo si el articulo existe, solo si tiene saldo pendiente.
	if (VS_saldo > 0) then
		if (not controlarArticulo(VS_idArticulo)) then Call setError(ARTICULO_NO_EXISTE)	
	end if
	if (VS_saldo < 0) then Call setError(CANTIDAD_NO_NEGATIVA)
	if ((VS_cantidad - VS_saldo) > VS_cumplido) then Call setError(NUEVA_MAY_CUMPLIDOS)			
	if (VS_cumplido > VS_cantidad) then Call setError(NUEVA_MAY_ORIGINAL)
	if not hayError then 
		controlarArticuloValeAjustePedido = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloValeAjusteTransf() 	
	controlarArticuloValeAjusteTransf = false
	'Controlo si el articulo existe	
	Call controlarArticulo(VS_idArticulo)
	if (VS_saldo <= 0) then Call setError(CANTIDAD_NO_NEGATIVA)
	if ((VS_cantidad - VS_cumplido) < VS_saldo) then Call setError(NUEVA_MAY_RECIBIDOS)
	if not hayError then 
		controlarArticuloValeAjusteTransf = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function controlarArticuloVRS()
	Dim strSQL, idDivision
	controlarArticuloVRS = false
	'Controlo si el articulo existe	
	Call controlarArticulo(VS_idArticulo)
	if (VS_saldo < 0) then 
		setError(CANTIDAD_NO_NEGATIVA)			
	else
		idDivision = getDivisionAlmacen(VS_idAlmacen)
		strSQL =" SELECT D.IDARTICULO, D.CANTIDAD, D.FACTURADO " &_
				" FROM (SELECT * FROM TBLCTZCABECERA WHERE IDDIVISION=" & idDivision & " and ESTADO NOT IN ('" & CTZ_FACTURADA & "', '" & CTZ_ANULADA & "')) C " &_ 
				" INNER JOIN TBLCTZDETALLE D ON C.IDCOTIZACION = D.IDCOTIZACION AND D.IDARTICULO = " & VS_idArticulo 				
		Call executeQueryDB(DBSITE_SQL_INTRA, rsCTZ, "OPEN", strSQL)			
		while (not rsCTZ.Eof)
		    if (CDbl(rsCTZ("CANTIDAD")) <> CDbl(rsCTZ("FACTURADO"))) then Call setError(COMPRAS_EN_CURSO)
		    rsCTZ.MoveNext()
		wend
		strSQL = " SELECT SUM(SALDO) AS SALDO, IDARTICULO " & _
				 " FROM ( " & _
				 "	  SELECT  D.IDARTICULO, " & _
				 "            C.CDVALE, " & _
				 "            C.PARTIDAPENDIENTE, "& _
				 "            CASE(C.CDVALE) WHEN '"& CODIGO_VS_DEVOLUCION &"'  THEN SUM(-D.CANTIDAD) "& _
				 "                           WHEN '"& CODIGO_VS_PRESTAMO &"'    THEN SUM(D.CANTIDAD)  "& _
				 "							 WHEN '"& CODIGO_VS_AJUSTE_VALE &"' THEN SUM(-D.CANTIDAD) "& _
				 "			  END  " & chr(34) & "SALDO" & chr(34) & _
				 "	  FROM TBLVALESCABECERA C "& _
		         "        INNER JOIN TBLVALESDETALLE D ON C.IDVALE = D.IDVALE "&_
		         "		  INNER JOIN TBLPMCABECERA PM ON C.PARTIDAPENDIENTE = PM.IDPEDIDO "&_
		         "	  WHERE C.CDVALE IN ('"& CODIGO_VS_PRESTAMO &"','"& CODIGO_VS_DEVOLUCION &"','"& CODIGO_VS_AJUSTE_VALE &"') "& _
		         "             AND C.IDALMACEN = "& VS_idAlmacen &" AND C.ESTADO = "& ESTADO_ACTIVO &" AND D.IDARTICULO = " & VS_idArticulo & _
				 "    GROUP BY C.PARTIDAPENDIENTE,D.IDARTICULO, C.CDVALE "& _
				 " ) T1 GROUP BY T1.IDARTICULO "		
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.EoF then
			' Si el Saldo es distinto de 0, el articulo tiene prestamos activos
			if Cdbl(rs("SALDO")) <> 0 then Call setError(PRESTAMOS_EN_CURSO)
		end if
	end if	
	if not hayError() then
		controlarArticuloVRS = true
	else 
		controlarArticuloVRS = false
	end if	
End Function 
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticuloValeAju() 	
	Dim rsStockArticulo, strSQL
	controlarArticuloValeAju = false
	'Controlo si el articulo existe	
	if (controlarArticulo(VS_idArticulo)) then
		strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & vs_idArticulo & " and IDALMACEN=" & VS_idAlmacen
		Call executeQueryDB(DBSITE_SQL_INTRA, rsStockArticulo, "EXEC", strSQL)		
		if not rsStockArticulo.eof then
			if CDbl(VS_saldo) > (CDbl(rsStockArticulo("EXISTENCIA")) + CDbl(rsStockArticulo("SOBRANTE"))) and VS_cdVale <> CODIGO_PM then setError(STOCK_INSUFICIENTE)
			devueltos = 0
			if VS_cantidad < VS_Aju_Entregados then
				setError(ADEVOLVER_MAY_PRESTADOS)	
			end if
		else
			if VS_cdVale <> CODIGO_PM then
				setError(ARTICULO_NO_EXISTE)			
			 end if	
		end if
	else
		setError(ARTICULO_NO_EXISTE)			
	end if
	if not hayError then 
		controlarArticuloValeAju = true
	else 
		addArticulosConErrores arrArticulosConErrores, vs_idArticulo
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Devuelve la cantidad de articulos que tiene un pedido
Function getCantidadArticulos(pIdPedido)
	Dim rs, strSQL, rtrn
	rtrn = 0
	strSQL="select count(*) as Cantidad from TBLVALESDETALLE where IDPEDIDO=" & pIdPedido
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (not rs.eof) then
		if not isnull(rs("Cantidad")) then rtrn = rs("Cantidad")
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)		
	getCantidadArticulos = rtrn
End Function
'---------------------------------------------------------------------------------------------
Sub grabarDetalles(p_idPM, p_idVS) 
	if (initArticulosVale()) then		
		while ((readNextArticuloVale(p_idVS)))
			if estaPMReferencia then
				call actualizarPMDetalle(p_idPM, VS_saldo)
			else
				call grabarPMDetalle(p_idPM, VS_cantidad, VS_saldo)
			end if
			if (grabarValeDetalle(p_idVS, p_idPM)) then
				call actualizarStock()
			end if
		wend	
	end if		
End sub
'---------------------------------------------------------------------------------------------
Function grabarHeaderPMVale()
	grabarHeaderPMVale = grabarHeaderPMInsert()
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderVale(byref pIdVale, pPartPend)
	Dim strSQL, rs, dte, idPedido
	if pIdVale <> 0 then
		strSQL= "Update TBLVALESCABECERA SET IDOBRA=" & VS_idObra & ", IDBUDGETAREA=" & VS_idBudgetArea & ", IDBUDGETDETALLE=" & VS_idBudgetDetalle & _
				", IDSECTOR=" & VS_idSector & " where IDVALE= " & pIdVale 
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
	else
		nrvale = getNumeracionVale(VS_idAlmacen)
		strSQL= "Insert into TBLVALESCABECERA(CDVALE, IDALMACEN, NRVALE, FECHA, IDOBRA, IDBUDGETAREA, IDBUDGETDETALLE, CDSOLICITANTE, CDUSUARIO, MOMENTO, PARTIDAPENDIENTE, IDSECTOR) values(" 
		strSQL = strSQL & "'" & VS_cdVale & "', " & VS_idAlmacen & ", '" & nrvale & "', " & GF_DTE2FN(VS_FechaSolicitud) & ", " & VS_idObra & ", " & VS_idBudgetArea & ", " & VS_idBudgetDetalle & ", '" & VS_cdSolicitante & "' "
		strSQL = strSQL & ", '" & session("Usuario") & "', '" & session("MmtoSistema") & "'," & pPartPend & ", " & VS_idSector & ")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
		strSQL = "Select MAX(IDVALE) as IDVALE from TBLVALESCABECERA where IDALMACEN=" & VS_idAlmacen
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
		pIdVale = rs("IDVALE")
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)		
	end if
	grabarHeaderVale = true
End Function
'-------------------------------------------------------------------------------
'Funcion responsable de grabar las solicitudes de firmas para los vales de ajuste 
'y reclasificación de stock y sus correspondientes anulaciones. 
Function grabarRegistroFirmas(pIdVale)
	Dim strSQL, conn, rs, cdGerente,cdCoordAud, cdResponsable

	cdResponsable = VS_cdSolicitante
	if (cdResponsable = "") then cdResponsable = session("usuario")
	
	cdGerente   = VS_NO_USER
	cdCoordAud	= VS_PORT_SUPERVISOR_USER
				
	strSQL = "Insert into TBLVALESFIRMAS (IDVALE,SECUENCIA,CDUSUARIO,FECHAFIRMA,HKEY) "
	strSQL = strSQL & "VALUES("&pIdVale&","&VS_FIRMA_RESPONSABLE&",'"&cdResponsable&"',"
	if (cdResponsable = session("usuario")) then
		strSQL = strSQL & session("MmtoDato") & ", '" & A_MANO & "')"
	else
		strSQL = strSQL & " null, null)"
	end if
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
	strSQL = "insert into TBLVALESFIRMAS (IDVALE,SECUENCIA,CDUSUARIO,FECHAFIRMA,HKEY) VALUES("&pIdVale&","&VS_FIRMA_GERENTE&",'"&cdGerente&"',null,null)"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
	strSQL = "insert into TBLVALESFIRMAS (IDVALE,SECUENCIA,CDUSUARIO,FECHAFIRMA,HKEY) VALUES("&pIdVale&","&VS_FIRMA_COORD_AUDIT&",'"&cdCoordAud&"',null,null)"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
	
End Function
'---------------------------------------------------------------------------------------------
Function grabarValeDetalle(p_idVS, p_idPM)
	Dim strSQL, rs, diff, ndiff

	grabarValeDetalle = false
	select case	(VS_cdVale)
			case CODIGO_VS_TRANSFERENCIA
			    VS_existencia = 0
	            VS_sobrante = 0
	            'JAS - SE HABILITO LA TRANSFERENCIA DE SOBRANTES - TAMBIEN SE MODIFICO EL CONTROL (20/11/2015)
				'if (getDivisionAlmacen(VS_idAlmacen) <> getDivisionAlmacen(VS_idAlmacenDest)) then
					'Si la transferencia es entre almacenes diferentes solo se puede transferir existencia.
				'	VS_existencia = VS_saldo
				'else					
					'La transferencia es entre almacenes del mismo puerto.
					strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & vs_idArticulo & " and IDALMACEN=" & VS_idAlmacen
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)							
					diff = CDbl(VS_saldo) - CDbl(rs("SOBRANTE"))					
					if (diff > 0) then
						'Se transfirienron más articulos que el sobrante.
						VS_sobrante = CDbl(rs("SOBRANTE"))
						VS_existencia = diff
					else
						'Se transfirienron menos articulos que el sobrante.						
						VS_sobrante = CDbl(VS_saldo)
						VS_existencia = 0						
					end if					
				'end if				
			case CODIGO_VS_AJUSTE_STOCK	
			    VS_existencia = 0
	            VS_sobrante = 0
				'Los ajustes de stock nunca pueden agregar existencia!
				'Si el AJS agrega stock sera siempre Sobrante y si quita primero quitara todo el sobrante
				' y si necesita más recien ahi quitara existencias
				strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & vs_idArticulo & " and IDALMACEN=" & VS_idAlmacen
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
				if not rs.eof then					
					diff = CDbl(VS_saldo) - CDbl(VS_cantidad)
					ndiff = CDbl(VS_cantidad) - CDbl(VS_saldo)
					if (diff > 0) then
						'Hay articulos de mas
						VS_sobrante = diff
						VS_existencia = 0
					else
						'Hay articulos de menos
						if (ndiff <= CDbl(rs("SOBRANTE"))) then						
							VS_sobrante = diff
							VS_existencia = 0
						elseif (ndiff <= (CDbl(rs("EXISTENCIA")) + CDbl(rs("SOBRANTE")))) then																					
							VS_sobrante = -CDbl(rs("SOBRANTE"))
							VS_existencia = CDbl(rs("SOBRANTE")) - ndiff
						end if
					end if
					VS_saldo = diff
				else
					VS_sobrante = VS_saldo
					VS_existencia = 0
				end if
			case CODIGO_VS_SALIDA, CODIGO_VS_PRESTAMO
			    VS_existencia = 0
	            VS_sobrante = 0				
				strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & vs_idArticulo & " and IDALMACEN=" & VS_idAlmacen
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
				if not rs.eof then
					if CDbl(VS_saldo) <= CDbl(rs("SOBRANTE")) then
						VS_sobrante = VS_saldo
					elseif (round(CDbl(VS_saldo), 3) <= round((CDbl(rs("EXISTENCIA")) + CDbl(rs("SOBRANTE"))), 3)) then
						VS_sobrante = CDbl(rs("SOBRANTE"))
						VS_existencia = VS_saldo - VS_sobrante
					end if
				end if
			case CODIGO_VS_ENTRADA
			    VS_existencia = 0	            
				VS_sobrante = VS_saldo
			case CODIGO_VS_DEVOLUCION
			        VS_existencia = 0
	                VS_sobrante = 0
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_PRESTAMO, exVMP, soVMP)
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_PRESTAMO_X, exXMP, soXMP)
					Call getCantidadDevueltaExSo(p_idPM, vs_idArticulo, exD, soD)
					VS_existencia = exVMP - exXMP - exD
					VS_sobrante   = soVMP - soXMP - soD
					if (CDbl(VS_saldo) <= CDbl(VS_existencia)) then
						VS_existencia = CDbl(VS_saldo)
						VS_sobrante = 0
					elseif CDbl(VS_saldo) <= (CDbl(VS_existencia) + CDbl(VS_sobrante)) then
						VS_sobrante = VS_saldo - VS_existencia
					end if
			case CODIGO_VS_RECEPCION 
			        VS_existencia = 0
	                VS_sobrante = 0
					'Tomo la transferido.
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_TRANSFERENCIA, exVMT, soVMT)
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_TRANSFERENCIA_X, exXMT, soXMT)
					'Tomo lo recibido. Se asume que siempre se recibe primero la existencia.
					Call getCantidadRecibidaExSo(p_idPM, vs_idArticulo, exR, soR)
					'Se cargan las variables de Existencia y sobrante
					'	Existencia = lo transferido - las transferencias anuladas - lo recibido (Pendiente de recepcion)
					'	Sobrante = lo transferido - las transferencias anuladas - lo recibido (Pendiente de recepcion)
					VS_existencia = exVMT - exXMT - exR
					VS_sobrante   = soVMT - soXMT - soR		
					
					if (CDbl(VS_saldo) <= CDbl(VS_existencia)) then
						'El valor que se pretende recibir es cubierto por lo que resta recibir de existencia
						VS_existencia = CDbl(VS_saldo)
						VS_sobrante = 0
					elseif CDbl(VS_saldo) <= (CDbl(VS_existencia) + CDbl(VS_sobrante)) then
						'El valor que se pretende recibir NO es cubierto por lo que resta recibir de existencia
						'Existencia queda cargado con el valor que resta recibir y sobrante se carga con lo que resta
						VS_sobrante = VS_saldo - VS_existencia
					end if					
			case CODIGO_VS_AJUSTE_TRANSFERENCIA 
			        VS_existencia = 0
	                VS_sobrante = 0
					'Tomo la transferido.
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_TRANSFERENCIA, exVMT, soVMT)
					Call getCantidadExSo(p_idPM, vs_idArticulo, CODIGO_VS_TRANSFERENCIA_X, exXMT, soXMT)
					'Tomo lo recibido.
					Call getCantidadRecibidaExSo(p_idPM, vs_idArticulo, exR, soR)
					VS_existencia = exVMT - exXMT - exR
					VS_sobrante   = soVMT - soXMT - soR
					'Se asume que primero se pierda el sobrante.
					if (CDbl(VS_saldo) <= CDbl(VS_sobrante)) then
						VS_sobrante = CDbl(VS_saldo)
						VS_existencia = 0
					elseif CDbl(VS_saldo) <= (CDbl(VS_existencia) + CDbl(VS_sobrante)) then
						VS_existencia = VS_saldo - VS_sobrante
					end if
			case CODIGO_VS_AJUSTE_VALE
			    VS_existencia = 0
	            VS_sobrante = 0
				call getCantidadAjuExSo(p_idPM, vs_idArticulo, VS_existencia, VS_sobrante)
				if CDbl(VS_saldo) <= CDbl(VS_sobrante) then
					VS_sobrante = CDbl(VS_saldo)
					VS_existencia = 0
				elseif CDbl(VS_saldo) <= (CDbl(VS_existencia) + CDbl(VS_sobrante)) then
					VS_existencia = VS_saldo - VS_sobrante
				end if	
            case CODIGO_VS_RECLASIFICACION_STOCK
                '/* SE SETEAN LOS VALORES DIRECTAMENTE EN EL ARCHIVO almacenValesVRS.asp */				
			case else
			    VS_sobrante = 0
				VS_existencia = VS_saldo 					
	end select	
	'Response.Write "(" & VS_existencia & "),(" & VS_sobrante & ")"
	if ((VS_existencia <> 0) or (VS_sobrante <> 0)) then 
		strSQL= "Insert into TBLVALESDETALLE(IDVALE, IDARTICULO, CANTIDAD, EXISTENCIA, SOBRANTE) values(" & p_idVS & ", " & VS_idArticulo & ", " & VS_saldo & "," & VS_existencia & "," & VS_sobrante & ")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
		grabarValeDetalle = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
Sub getCantidadAjuExSo(pIdPM, pIdArt, byref pExistencia, byref pSobrante)
'XXX SQL juntar	
	Dim exVMP, soVMP, exAJU, soAJU, exVMD, soVMD
	Dim exXMP, soXMP, exXJU, soXJU, exXMD, soXMD
	pExistencia = 0
	pSobrante = 0
	'Todo lo que ya se presto para ese PM-ART
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_PRESTAMO		, exVMP, soVMP)
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_PRESTAMO_X	, exXMP, soXMP)
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE	, exAJU, soAJU)
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE_X	, exXJU, soXJU)
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION	, exVMD, soVMD)
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION_X	, exXMD, soXMD)
	
	pExistencia	= (exVMP - exXMP) - (exAJU - exXJU) - (exVMD - exXMD)
	pSobrante	= (soVMP - soXMP) - (soAJU - soXJU) - (soVMD - soXMD)
		
End Sub
'---------------------------------------------------------------------------------------------
Function getCantidadRecibidaExSo(pIdPM, pIdArt, byref pExistencia, byref pSobrante)

	Dim existenciaVMR, sobranteVMR, existenciaAJT, sobranteAJT
	Dim	existenciaXMR, sobranteXMR, existenciaXJT, sobranteXJT

	'Se suman las devoluciones.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_RECEPCION, existenciaVMR, sobranteVMR)	
	'Se suman los ajustes anulados.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_TRANSFERENCIA, existenciaAJT, sobranteAJT)
	'Se suman las devoluciones anuladas.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_RECEPCION_X, existenciaXMR, sobranteXMR)
	'Se suman los ajustes anulados.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_TRANSFERENCIA_X, existenciaXJT, sobranteXJT)
	'Se totaliza	
	pExistencia = existenciaVMR + existenciaAJT - existenciaXMR - existenciaXJT
	pSobrante   = sobranteVMR + sobranteAJT - sobranteXMR - sobranteXJT
End Function
'---------------------------------------------------------------------------------------------
Function getCantidadDevueltaExSo(pIdPM, pIdArt, byref pExistencia, byref pSobrante)
	
	Dim exVMD, soVMD, exXMD, soXMD
	Dim	exAJU, soAJU, exXJU, soXJU

	'Se suman las devoluciones.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION, exVMD, soVMD)
	'Se suman los ajustes.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE, exAJU, soAJU)
	'Se suman las devoluciones anuladas.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION_X, exXMD, soXMD)
	'Se suman los ajustes anulados.
	Call getCantidadExSo(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE_X, exXJU, soXJU)
	'Se totaliza
	pExistencia = exVMD + exAJU + exXMD + exXJU
	pSobrante   = soVMD + soAJU + soXMD + soXJU
End Function
'---------------------------------------------------------------------------------------------
'Devuelve las cantidades de Existencia y Sobrante que fueron registrdas en los 
'vales del tipo indicado y que estan asociados al pedido  de materiales indicado
'OJO! toma toso los vales, esten anulados o no!
Sub getCantidadExSo(pIdPM, pIdArt, pCdVale, byref pExistencia, byref pSobrante)
	Dim strSQL, rs, oConn, rtrn
	pExistencia = 0
	pSobrante = 0
	strSQL= "SELECT SUM(EXISTENCIA) AS EXISTENCIA, SUM(SOBRANTE) AS SOBRANTE FROM TBLVALESCABECERA C INNER JOIN TBLVALESDETALLE D " & _
			" ON C.IDVALE=D.IDVALE WHERE C.PARTIDAPENDIENTE=" & pIdPM & " AND D.IDARTICULO=" & pIdArt & _
			" AND C.CDVALE='" & pCdVale & "' GROUP BY D.IDARTICULO"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not rs.eof then 
		if (rs("EXISTENCIA") <> "") then pExistencia = CDbl(rs("EXISTENCIA"))
		if (rs("SOBRANTE") <> "") then pSobrante = CDbl(rs("SOBRANTE"))
	end if		
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)	
End Sub
'---------------------------------------------------------------------------------------------
'Se calcula la cantidad recibida por el destino.
Function getCantidadRecibida(pIdPM, pIdArt)
	Dim vmr, aju, xmd, xju

	'Se suman las devoluciones.
	vmr = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_RECEPCION)	
	'Se suman los ajustes anulados.
	ajt = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_AJUSTE_TRANSFERENCIA)
	'Se suman las devoluciones anuladas.
	xmr = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_RECEPCION_X)
	'Se suman los ajustes anulados.
	xjt = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_AJUSTE_TRANSFERENCIA_X)
	'Se totaliza
	getCantidadRecibida=(cdbl(vmr) - cdbl(xmr))	+ (cdbl(ajt) - cdbl(xjt))
End Function
'---------------------------------------------------------------------------------------------
'Se calcula la cantidad del artículo que fue prestado y ya fue devuelto al almacen.
Function getCantidadDevuelta(pIdPM, pIdArt)
	Dim vmd, aju, xmd, xju

	'Se suman las devoluciones.
	vmd = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION)
	'Se suman los ajustes.
	aju = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE)
	'Se suman las devoluciones anuladas.
	xmd = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_DEVOLUCION_X)
	'Se suman los ajustes anulados.
	xju = getTotalArticuloxVale(pIdPM, pIdArt, CODIGO_VS_AJUSTE_VALE_X)
	'Se totaliza
	getCantidadDevuelta=(cdbl(vmd) - cdbl(xmd))	+ (cdbl(aju) - cdbl(xju))
End Function
'---------------------------------------------------------------------------------------------
Function getTotalArticuloxVale(pIdPM, pIdArt, pCdVale)
	Dim strSQL, rs, oConn, rtrn
	rtrn = 0
	strSQL= "SELECT SUM(CANTIDAD) AS DEVUELTOS FROM TBLVALESCABECERA C INNER JOIN TBLVALESDETALLE D " & _
			" ON C.IDVALE=D.IDVALE WHERE C.PARTIDAPENDIENTE=" & pIdPM & " AND D.IDARTICULO=" & pIdArt & _
			" AND C.CDVALE='" & pCdVale & "' and C.ESTADO=" & ESTADO_ACTIVO & " GROUP BY D.IDARTICULO"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not rs.eof then rtrn = rs("DEVUELTOS")	
	getTotalArticuloxVale = rtrn
End Function
'---------------------------------------------------------------------------------------------
Function actualizarStock()
	Dim strSQL, rs
	strSQL = "SELECT * FROM TBLARTICULOSDATOS WHERE IDARTICULO=" & VS_idArticulo & " AND IDALMACEN=" & VS_idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not rs.eof then 
		'Actualizo lo articulo
		strSQL= "UPDATE TBLARTICULOSDATOS " 
		select case (VS_cdVale)
				case CODIGO_VS_ENTRADA
					strSQL = strSQL & " SET SOBRANTE = (SOBRANTE + " & VS_saldo & ")," 						
				case CODIGO_VS_DEVOLUCION, CODIGO_VS_RECEPCION, CODIGO_VS_AJUSTE_STOCK, CODIGO_VS_RECLASIFICACION_STOCK
					strSQL = strSQL & " SET EXISTENCIA = (EXISTENCIA + " & VS_existencia & ")," 
					strSQL = strSQL & " SOBRANTE = (SOBRANTE + " & VS_sobrante & ")," 
				case else
					strSQL = strSQL & " SET EXISTENCIA = (EXISTENCIA - " & VS_existencia & ")," 
					strSQL = strSQL & " SOBRANTE = (SOBRANTE - " & VS_sobrante & ")," 
		end select
		strSQL = strSQL & " CDUSUARIO = '" & session("Usuario") & "', MOMENTO = '" & session("MmtoSistema") & "' WHERE IDARTICULO= " & VS_idArticulo
		strSQL = strSQL & " AND IDALMACEN = " & VS_idAlmacen
	else
		'Inserto el articulo
		strSQL= "INSERT INTO TBLARTICULOSDATOS VALUES(" & VS_idArticulo 
		select case (VS_cdVale)
				case CODIGO_VS_ENTRADA
					strSQL = strSQL & ",0," & VS_saldo & "," 		
				case else
					strSQL = strSQL & ", " & VS_existencia & "," & VS_sobrante & "," 
		end select
		strSQL = strSQL & "'" & session("Usuario") & "','" & session("MmtoSistema") & "',"
		strSQL = strSQL & VS_idAlmacen & ",0,0,0,0,'')"
	end if
	'Response.Write strSQL & "<br>"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
End Function
'---------------------------------------------------------------------------------------------
'Borra todas las variables del HeaderVS
Function clearHeaderVale()
	VS = 0
	VS_idPedido = ""
	VS_idAlmacen = 0
	VS_idAlmacenDest = 0
	VS_FechaSolicitud = ""
	VS_FechaRequerido = ""
	VS_cdSolicitante = ""
	VS_idObra = 0
	VS_idSector = 0
	VS_idBudgetArea = 0	
	VS_idBudgetDetalle = 0	
	VS_usuario = ""
	VS_momento = ""
	VS_hayCabecera = false
End function
'---------------------------------------------------------------------------------------------
Function clearArticulo()
	VS_idArticulo = 0	
	VS_dsArticulo = ""
	VS_cantidad = 0
	VS_saldo = 0
	VS_idUnidad=0
	VS_ArticuloActual=0
	VS_abreviaturaUnidad=""
	VS_cumplido = 0
	VS_existencia = 0 
	VS_sobrante = 0
End Function
'---------------------------------------------------------------------------------------------
'Devuelve lista de almacenes disponibles
function obtenerListaAlmacenes(p_idAlmacen)
	Dim strSQL, rtrn, rsOLA
	rtrn = 0
	strSQL="select * from TBLALMACENES "
	if not (p_idAlmacen = 0) then strSQL = strSQL & " where IDALMACEN = " & p_idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rsOLA, "OPEN", strSQL)		
	Set obtenerListaAlmacenes = rsOLA
End function
'---------------------------------------------------------------------------------------------
sub addArticulosConErrores (ByRef p_arrArticulosConErrores, ByRef p_idArticulo)
	dim iNewUBound
	iNewUBound = UBound(p_arrArticulosConErrores) + 1
	redim preserve p_arrArticulosConErrores(iNewUBound)
	p_arrArticulosConErrores(iNewUBound) = p_idArticulo
end sub
'---------------------------------------------------------------------------------------------
sub grabarComentarioVale(pIdVale, pComentario)
	Dim strSQL, rs, rsExec, connExec 
	if len(pComentario) > 0 then
		strSQL= "Select * from TBLVALESCOMENTARIOS where IdVale=" & pIdVale
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
			if not rs.eof then
				strSQL= "Update TBLVALESCOMENTARIOS set Comentario='" & pComentario & "' where IdVale=" & pIdVale
			else
				strSQL= "Insert into TBLVALESCOMENTARIOS values(" & pIdVale & ",'" & pComentario & "')" 
			end if
			Call executeQueryDB(DBSITE_SQL_INTRA, rsExec, "EXEC", strSQL)		
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)		
	end if
End sub
'---------------------------------------------------------------------------------------------
Function ConfirmaValeAnular(pIdVale)
	Dim strSQL, rsVale, rtrn, estadoVale
	strSQL= "SELECT ESTADO FROM TBLVALESCABECERA WHERE IDVALE=" & pIdVale
	Call executeQueryDB(DBSITE_SQL_INTRA, rsVale, "OPEN", strSQL)		
	if not rsVale.eof then estadoVale = rsVale("ESTADO")
	Call executeQueryDB(DBSITE_SQL_INTRA, rsVale, "OPEN", strSQL)		
	if (estadoVale = ESTADO_ACTIVO) then
		rtrn = true
	else
		rtrn = false
	end if
	ConfirmaValeAnular = rtrn
End Function
'---------------------------------------------------------------------------------------------
function getValeRelacionado(pIdVale)
	Dim strSQL, rs, rtrn
	rtrn = ""
	strSQL= "SELECT *  FROM TBLVALESRELACIONES WHERE IDVALE_1=" & pIdVale & " or IDVALE_2=" & pIdVale
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	if (not rs.eof) then 
		if (CLng(rs("IDVALE_1")) = pIdVale) then
			rtrn  = rs("IDVALE_2")
		else
			rtrn  = rs("IDVALE_1")
		end if
		strSQL= "SELECT *  FROM TBLVALESCABECERA WHERE IDVALE=" & rtrn
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
		if (not rs.eof) then rtrn = rs("NRVALE")
	end if	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getValeRelacionado = rtrn
End function
'---------------------------------------------------------------------------------------------
function getComentarioVale(pIdVale)
	Dim strSQL, rs, rtrn
	rtrn = ""
	strSQL= "SELECT COMENTARIO FROM TBLVALESCOMENTARIOS WHERE IDVALE=" & pIdVale
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	if not rs.eof then rtrn = rs("COMENTARIO")
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	getComentarioVale = rtrn
End function
'---------------------------------------------------------------------------------------------
function getLeyendaCdVale(pCdVale)
	select case pCdVale
		case CODIGO_VS_SALIDA
			getLeyendaCdVale = LEYENDA_VS_SALIDA
		case CODIGO_VS_SALIDA_X
			getLeyendaCdVale = LEYENDA_VS_SALIDA_X
		case CODIGO_VS_ENTRADA
			getLeyendaCdVale = LEYENDA_VS_ENTRADA
		case CODIGO_VS_ENTRADA_X
			getLeyendaCdVale = LEYENDA_VS_ENTRADA_X
		case CODIGO_VS_PRESTAMO
			getLeyendaCdVale = LEYENDA_VS_PRESTAMO
		case CODIGO_VS_PRESTAMO_X
			getLeyendaCdVale = LEYENDA_VS_PRESTAMO_X
		case CODIGO_VS_DEVOLUCION
			getLeyendaCdVale = LEYENDA_VS_DEVOLUCION
		case CODIGO_VS_DEVOLUCION_X
			getLeyendaCdVale = LEYENDA_VS_DEVOLUCION_X
		case CODIGO_VS_TRANSFERENCIA
			getLeyendaCdVale = LEYENDA_VS_TRANSFERENCIA
		case CODIGO_VS_TRANSFERENCIA_X
			getLeyendaCdVale = LEYENDA_VS_TRANSFERENCIA_X
		case CODIGO_VS_RECEPCION
			getLeyendaCdVale = LEYENDA_VS_RECEPCION
		case CODIGO_VS_RECEPCION_X
			getLeyendaCdVale = LEYENDA_VS_RECEPCION_X
		case CODIGO_VS_AJUSTE_VALE
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_VALE
		case CODIGO_VS_AJUSTE_VALE_X
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_VALE_X
		case CODIGO_VS_AJUSTE_STOCK
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_STOCK
		case CODIGO_VS_AJUSTE_STOCK_X
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_STOCK_X
		case CODIGO_VS_AJUSTE_PEDIDO
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_PEDIDO
		case CODIGO_VS_AJUSTE_PEDIDO_X
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_PEDIDO_X
		case CODIGO_VS_AJUSTE_TRANSFERENCIA
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_TRANSFERENCIA
		case CODIGO_VS_AJUSTE_TRANSFERENCIA_X
			getLeyendaCdVale = LEYENDA_VS_AJUSTE_TRANSFERENCIA_X
		case CODIGO_PM
			getLeyendaCdVale = LEYENDA_PM
		case CODIGO_VS_FIX
			getLeyendaCdVale = LEYENDA_VS_FIX
		case CODIGO_VS_RECLASIFICACION_STOCK
			getLeyendaCdVale = LEYENDA_VS_RECLASIFICACION_STOCK
		case CODIGO_VS_RECLASIFICACION_STOCK_X
			getLeyendaCdVale = LEYENDA_VS_RECLASIFICACION_STOCK_X
	end select
end function
'----------------------------------------------------------------------------
Function obtenerDivisionVale(pIdVale)
	Dim strSQL, oConn, rs,rtrn
	
	rtrn = ""
	
	strSQL = "select div.IDDIVISION" & _
			" from TBLVALESCABECERA v " & _
			" inner join TBLALMACENES alm on alm.IDALMACEN = v.IDALMACEN " & _
			" inner join TBLDIVISIONES div on div.IDDIVISION = alm.IDDIVISION " & _
			" where v.IDVALE = " & pIdVale
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	
	if (not rs.EoF) then
		rtrn = cstr(rs("IDDIVISION"))
	end if
	
	obtenerDivisionVale = rtrn
End Function
'----------------------------------------------------------------------------
'/**
' * Funcion    : getNroVale
' * Descripcion: Obtiene el Nro de vale que fue asignado al vale indicado.
' * Parametros : idAlmacen [in] ID del Almacen al quen pertenece el vale.
' *
' * Valor Devuelto:
' * Devuelve el Nro de Vale asigando según ID indicado.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 21/09/2010
' */
Function getNroVale(pIdVale)
	Dim strSQL, oConn, rs,rtrn
	
	strSQL="Select NRVALE from TBLVALESCABECERA where IDVALE=" & pIdVale
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	rtrn = ""
	if (not rs.eof) then rtrn = Trim(rs("NRVALE"))
	getNroVale = rtrn
End Function
'----------------------------------------------------------------------------
'/**
' * Funcion    : getNumeracionVale
' * Descripcion: Funcion responsable de generar la numeración del vale
' *				 respetando el formato establecido.
' * Parametros : idAlmacen [in] ID del Almacen al quen pertenece el vale.
' *
' * Valor Devuelto:
' * Devuelve un string con el formato correspondiente al proximo numero de vale.
' * Formato <AÑO>-<ID ALMACEN>-<NUMERO>
' *
' * Autor: Javier A. Scalisi
' * Fecha: 21/09/2010
' */
Function getNumeracionVale(idAlmacen)
	Dim strSQL, oConn, rs,rtrn, nr, clave
	
	clave= PREFIX_VNR & "_" & idAlmacen & "_" & Right(year(now()),2)
	strSQL="Select * from TBLNUMERACION where PREFIJO='" & PREFIX_VNR & "' and CLAVE='" & clave & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	if (not rs.eof) then
		nr=CLng(rs("VALOR"))+1
		strsql = "Update TBLNUMERACION set VALOR=" & nr & " where CLAVE = '" & clave & "' and PREFIJO='" & PREFIX_VNR & "'"
	else
		nr=1
		strSQL = "Insert into TBLNUMERACION values('" & PREFIX_VNR & "','" & clave & "', 1)"
	end if	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)		
	getNumeracionVale = Right(year(now()),2) & "-" & idAlmacen & "-" & GF_nDigits(nr, 6)	
End Function

'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/12/10
' Objetivo:	
'			obtener el importe total de los vales de una obra
' Parametros:
'			idObra  [int]
'			pArea	[int]		Area del budget a consultar
'			pDetalle[int]		Detalle del budget a consultar
'			pVale	[string]	Tipo de vale que se quiere totalizar, si es vacio ("") se
'								totalizaran todos
'			pMoneda [string]	Moneda con la que se procesara la operacion
' Devuelve:
'			[int] importe Total 
'--------------------------------------------------------------------------------------------------
Function obtenerTotalValesObra(pIdObra,pArea,pDetalle,pVale,pMoneda)
	Dim strSQL, conn, rs, campoImporte,rtrn
	
	campoImporte = "det.vludolares"
	if (pMoneda = MONEDA_PESO) then campoImporte = "det.vlupesos"
	
	strSQL =          "select sum(det.existencia*"&campoImporte&") total from tblvalescabecera cab "
	strSQL = strSQL & " inner join tblvalesdetalle det on cab.idvale = det.idvale "
	strSQL = strSQL & " where cab.estado = " & ESTADO_ACTIVO
	strSQL = strSQL & " and cab.idobra = " & pIdObra
	if (pArea <> "" and pArea <> 0) then strSQL = strSQL & " and cab.idbudgetarea = " & pArea
	if (pDetalle <> "" and pDetalle <> 0) then strSQL = strSQL & " and cab.idbudgetdetalle = " & pDetalle
	if (pVale <> "") then strSQL = strSQL & " and cab.cdvale = '"&ucase(pVale)&"'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	
	rtrn = rs("total")
	if (isnull(rtrn)) then rtrn = 0
	
	obtenerTotalValesObra = rtrn
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/12/10
' Objetivo:	
'			Obtiene el importe total de cada Partida Presupuestaria de una obra
' Parametros:
'			pIdObra  		[int]
'			pMoneda 		[string]	Moneda con la que se procesara la operacion
'			pFechaLimite	[int]		Fecha limite de los vales a consultar
' Devuelve:
'			[dictionary] Key: area-detalle | value: importe 
'--------------------------------------------------------------------------------------------------
Function obtenerTotalValesObraPorPP(pIdObra,pMoneda,pFechaLimite)
	Dim strSQL,conn,rs,rtrn,campoImporte
	
	Set rtrn = createObject("Scripting.Dictionary")
	
	campoImporte = "det.vludolares"
	if (pMoneda = MONEDA_PESO) then campoImporte = "det.vlupesos"
	
	strSQL = "SELECT SUM("&campoImporte&"*cantidad) AS importe,cab.idbudgetarea area,cab.idbudgetdetalle detalle"
	strSQL = strSQL & " FROM tblvalescabecera cab "
	strSQL = strSQL & " INNER JOIN tblvalesdetalle det ON cab.idvale = det.idvale "
	strSQL = strSQL & " WHERE cab.idobra = " & pIdObra
	strSQL = strSQL & " AND cab.estado = " & ESTADO_ACTIVO
	strSQL = strSQL & " AND cab.momento <= " & pFechaLimite
	strSQL = strSQL & " GROUP BY cab.idbudgetarea,cab.idbudgetdetalle"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		

	while not rs.EoF
		myAreaDetalle = rs("area")&"-"&rs("detalle")
		if (not rtrn.exists(myAreaDetalle)) then
			myImporte = cdbl(rs("importe"))
			if (isnull(myImporte)) then myImporte = 0
			
			call rtrn.add(myAreaDetalle,myImporte)
		end if
		rs.MoveNext
	wend	
		
	set obtenerTotalValesObraPorPP = rtrn
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	01/12/10
' Objetivo:	
'			Obtiene el importe total de cada Partida Presupuestaria de una obra
' Parametros:
'			pIdObra  		[int]
'			pMoneda 		[string]	Moneda con la que se procesara la operacion
'			pFechaLimite	[int]		Fecha limite de los vales a consultar
' Devuelve:
'			[dictionary] Key: area-detalle | value: importe 
'--------------------------------------------------------------------------------------------------
Function obtenerTotalValesObraPorPPArea(pIdObra,pArea, pMoneda,pFechaLimite)
	Dim strSQL,conn,rs,rtrn,campoImporte
	
	Set rtrn = createObject("Scripting.Dictionary")
	
	campoImporte = "det.vludolares"
	if (pMoneda = MONEDA_PESO) then campoImporte = "det.vlupesos"
	
	strSQL = "SELECT SUM("&campoImporte&"*EXISTENCIA) AS importe,cab.idbudgetdetalle detalle"
	strSQL = strSQL & " FROM tblvalescabecera cab "
	strSQL = strSQL & " INNER JOIN tblvalesdetalle det ON cab.idvale = det.idvale "
	strSQL = strSQL & " WHERE cab.idobra = " & pIdObra
	strSQL = strSQL & " AND cab.estado = " & ESTADO_ACTIVO
	strSQL = strSQL & " AND cab.momento <= " & pFechaLimite
	strSQL = strSQL & " AND cab.idbudgetarea = " & pArea
	strSQL = strSQL & " GROUP BY cab.idbudgetdetalle"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		

	while not rs.EoF
		myDetalle = rs("detalle")
		if (not rtrn.exists(myDetalle)) then
			myImporte = cdbl(rs("importe"))
			if (isnull(myImporte)) then myImporte = 0
			
			call rtrn.add(myDetalle,myImporte)
		end if
		rs.MoveNext
	wend	
		
	set obtenerTotalValesObraPorPPArea = rtrn
End Function
'---------------------------------------------------------------------------------------------
' Autor: 	Ezequiel Bacarini
' Fecha: 	07/11/2011
' Objetivo:	
'			Indica si la transferencia es interdivional o no. 
' Parametros:
'			pIdVale [int] 		Id del vale que se va a analizar. El vale debe ser la recepcion
' Devuelve:
'			true:	es interdivisional
'			false:	no es interdivisional
Function esTransferenciaInterdivisional(pIdVale)
Dim strSQL, rs, cn, rtrn
rtrn = false
strSQL =	"SELECT * FROM TBLVALESCABECERA VCT " & _
			"	INNER JOIN TBLALMACENES ALM1 ON ALM1.IDALMACEN=VCT.IDALMACEN " & _
			"    INNER JOIN TBLVALESCABECERA VCR ON VCT.PARTIDAPENDIENTE=VCR.PARTIDAPENDIENTE AND VCT.CDVALE='" & CODIGO_VS_TRANSFERENCIA & "'" & _
			"    INNER JOIN TBLALMACENES ALM2 ON ALM2.IDALMACEN=VCR.IDALMACEN " & _
			"		WHERE ALM1.IDDIVISION<>ALM2.IDDIVISION AND VCR.IDVALE=" & pIdVale 
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.eof) then rtrn = true
end function
'---------------------------------------------------------------------------------------------
' Autor: 	Nahuel Ajaya
' Fecha: 	08/11/2013
' Objetivo:	
'			Devuelve la valuacion de un vale de Ajuste de Stock, a traves de la valorización de los items
' Parametros:
'			pIdVale [int] 		Id del Vale
'			pMoneda [char]		Moneda
' Devuelve:
'			precio vale
Function getValuacionValeAjuste(pIdVale ,pMoneda)
	Dim strSQL, rs, rtrn, campoImporte
	rtrn = 0
	campoImporte = "VLUDOLARES"
	if (pMoneda = MONEDA_PESO) then campoImporte = "VLUPESOS"
	if VS_cdVale = CODIGO_VS_AJUSTE_STOCK then
		strSQL = "SELECT A.IDVALE, B.MONTO "&_ 
				 "FROM TBLVALESCABECERA A " & _
				 "	  INNER JOIN ( SELECT ABS(SUM((EXISTENCIA * "& campoImporte &" + SOBRANTE * "& campoImporte &")/100)) AS MONTO, IDVALE"&_
				 "				   FROM TBLVALESDETALLE "&_
				 "				   WHERE IDVALE = "& pIdVale &_
				 "				   GROUP BY IDVALE) B ON B.IDVALE = A.IDVALE "&_
				 "WHERE A.ESTADO = " & ESTADO_ACTIVO				 
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then rtrn = Cdbl(rs("MONTO"))
	end if
	getValuacionValeAjuste = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
'Function grabarFirmasValeAJS : Graba las firmas del Vale de Ajuste de Stock
Function grabarFirmasValeAJS(pIdVale)
	Dim importeVale, rs, strSQL, limite
	Call grabarRegistroFirmas(pIdVale)
	'Obtengo el importe total del Vale y el importe limite (norma de auditoria)	
	limite		= getValorNorma(NORMA_VS_AJUSTE_AUTORIZADO)
	importeVale = getValuacionValeAjuste(idVale, getUnidadNorma(NORMA_VS_AJUSTE_AUTORIZADO))
	'Si supera el monto limite establecido agrego la firma del director
	if (cdbl(importeVale) > cdbl(limite))  then
		strSQL = "insert into TBLVALESFIRMAS (IDVALE,SECUENCIA,CDUSUARIO) VALUES("&pIdVale&","&VS_FIRMA_DIRECTOR&",'"&DIRECTOR_USER & ")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)			
	end if
End Function


%>