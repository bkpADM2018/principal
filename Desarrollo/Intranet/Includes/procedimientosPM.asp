<%
Dim pm, PM_FechaSolicitud,PM_idCotizacion, PM_FechaRequerido, PM_cdSolicitante, PM_dsSolicitante
Dim PM_usuario, PM_momento, PM_comentario
Dim PM_idArticulo, PM_dsArticulo, PM_cantidad, PM_saldo
Dim rsArticulos, PM_articuloStock, PM_idSector
Dim PM_CantArticulos, PM_ArticuloActual, PM_idPedido
Dim PM_hayCabecera, PM_abreviaturaUnidad, PM_idUnidad, PM_cdInterno
Dim PM_idObra, PM_idAlmacen, PM_idAlmacenDest, PM_idBudgetArea, PM_idBudgetDetalle
Dim PM_CantDetalle, PM_DetalleActual,PM_Tipo, PM_Transferir, PM_idDivision, PM_articuloError
'---------------------------------------------------------------------------------------------
Function initHeaderPMDB(pIdPedido)
	dim strSQL, rs, km, kc, tmp, pos
	
	PM_hayCabecera = False
	strSQL="select * from TBLPMCABECERA where IDPEDIDO=" & pIdPedido	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		PM = rs("IDPEDIDO")
		PM_FechaSolicitud = GF_FN2DTE(rs("FechaSolicitud"))
		PM_FechaRequerido = GF_FN2DTE(rs("FechaRequerido"))
		PM_cdSolicitante = rs("CDSOLICITANTE")
		PM_dsSolicitante = getUserDescription(PM_cdSolicitante)
		PM_idAlmacen = rs("IDALMACEN")
		PM_idAlmacenDest = rs("IDALMACENDEST")
		PM_idObra = rs("IDOBRA")		
		if (PM_idObra = "") then PM_idObra = 0
		PM_idSector = rs("IDSECTOR")
		if (PM_idSector = "") then PM_idSector = 0
		PM_idPedido = rs("idPedido")
		PM_usuario = rs("CDUSUARIO")
		PM_momento = rs("MOMENTO")
		PM_idBudgetArea = rs("idBudgetArea")
		PM_idBudgetDetalle = rs("idBudgetDetalle")
		PM_comentario = rs("COMENTARIOS")
		PM_hayCabecera = True
	end if		
End Function
'---------------------------------------------------------------------------------------------
function initArticulosDB(pIdPedido)
	dim strSQL, rs	
	PM_CantArticulos = 0
	PM_ArticuloActual=0
	initArticulosDB = false	
	if (PM_hayCabecera) then
		strSQL="select * from TBLPMDETALLE where IDPEDIDO=" & pIdPedido
		call executeQueryDb(DBSITE_SQL_INTRA, rsArticulos, "OPEN", strSQL)
		if (not rsArticulos.eof) then
			PM_CantArticulos = rsArticulos.RecordCount
			initArticulosDB = true
		end if
	end if
end function
'---------------------------------------------------------------------------------------------
Function readNextArticuloDB()
	dim strSQL, rs, km
	
	readNextArticuloDB = false	
	if (not rsArticulos.eof) then
		PM_idArticulo = rsArticulos("IDARTICULO")
		strSQL="select A.*, B.CDINTERNO, B.EXISTENCIA+B.SOBRANTE STOCK from TBLARTICULOS A left join TBLARTICULOSDATOS B on A.IDARTICULO=B.IDARTICULO and B.IDALMACEN=" & PM_idAlmacen & " where A.IDARTICULO=" & PM_idArticulo
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			PM_dsArticulo = rs("DSARTICULO")
			PM_idArticulo = rs("IDARTICULO")
			PM_idUnidad = rs("IDUNIDAD")
			PM_abreviaturaUnidad = getAbreviaturaUnidad(rs("IDUNIDAD"))
			PM_articuloStock = rs("STOCK")
			PM_cdInterno = rs("CDINTERNO")			
		end if		
		PM_cantidad = rsArticulos("CANTIDAD")
		PM_saldo = rsArticulos("SALDO")		
		rsArticulos.MoveNext()
		readNextArticuloDB = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderPMInsert()
	Dim strSQL, rs, dte, idPedido
	strSQL= "Insert into TBLPMCABECERA(IDALMACEN, IDOBRA, CDSOLICITANTE, FechaSolicitud, FechaRequerido, CDUSUARIO, MOMENTO, IDALMACENDEST, IDBUDGETAREA, IDBUDGETDETALLE, COMENTARIOS, IDSECTOR) values(" 
	strSQL = strSQL & PM_idAlmacen & ", " & PM_idObra & ", '" & PM_cdSolicitante & "', " & GF_DTE2FN(PM_FechaSolicitud) & ", " & GF_DTE2FN(PM_FechaRequerido) 
	strSQL = strSQL & ", '" & session("Usuario") & "', '" & session("MmtoSistema") & "'," & PM_idAlmacenDest & "," & PM_idBudgetArea & "," & PM_idBudgetDetalle & ", '" & PM_comentario & "', " & PM_idSector & ")"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = "Select MAX(IDPEDIDO) IDPEDIDO from TBLPMCABECERA where IDALMACEN=" & PM_idAlmacen
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	pm = rs("IDPEDIDO")
	grabarHeaderPMInsert = pm
End Function
'---------------------------------------------------------------------------------------------
'Borra todas las variables del HeaderPM
Function clearHeaderPM()
	pm = 0
	PM_idPedido = ""
	PM_idAlmacen = 0
	PM_idAlmacenDest = 0
	PM_FechaSolicitud = ""
	PM_FechaRequerido = ""
	PM_cdSolicitante = ""
	PM_idObra = 0
	PM_idSector = 0
	PM_usuario = ""
	PM_momento = ""
	PM_idBudgetArea = 0
	PM_idBudgetDetalle = 0
	PM_hayCabecera = false
End function
'---------------------------------------------------------------------------------------------
Function clearArticulo()
	PM_idArticulo = 0
	PM_dsArticulo = ""
	PM_cantidad = 0
	PM_idUnidad=0
	PM_abreviaturaUnidad=""
End Function
'---------------------------------------------------------------------------------------------
Function actualizarPMDetalle(pId, pIdArticulo, pSaldo)
	Dim strSQL, rs
	'Grabo los articulos
	strSQL= "UPDATE TBLPMDETALLE SET SALDO = (SALDO -" & pSaldo & ") "
	strSQL = strSQL & " WHERE IDPEDIDO = " & pId & " AND IDARTICULO= " & pIdArticulo
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Function
'---------------------------------------------------------------------------------------------
Function actualizarPMDetalleAju(pId, pIdArticulo, pCantidadOriginal, pNuevaCantidad, pSaldo)
	Dim strSQL, rs, nuevoSaldo
	'Grabo los articulos
	nuevoSaldo = pNuevaCantidad - (pCantidadOriginal-pSaldo)
	strSQL= "UPDATE TBLPMDETALLE SET SALDO =" & nuevoSaldo
	strSQL = strSQL & " WHERE IDPEDIDO = " & pId & " AND IDARTICULO= " & pIdArticulo
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Function
'---------------------------------------------------------------------------------------------
Function grabarPMDetalle(pId, pIdArticulo, pCantidad, pSaldo)
	Dim strSQL, rs
	'Grabo los articulos
	strSQL= "Insert into TBLPMDETALLE(IDPEDIDO, IDARTICULO, CANTIDAD, SALDO) values(" & pId & ", " & pIdArticulo & ", " & pCantidad & ", " & (pCantidad - pSaldo) & ")"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End Function
'---------------------------------------------------------------------------------------------
'Funcion que verifica si se puede acceder a modificar el PM, es decir, si el usuario es U o A del almacen o bien cargo el PM.
Function checkControlPM(idPM)
	
	Dim rs, rs1, flag, ret, strSQL, conn
	
	ret = false	
	strSQL = "Select * from TBLPMCABECERA where IDPEDIDO=" & idPM
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		set rs1 = obtenerListaAlmacenesUA()
		flag = false
		while ((not rs1.eof) and (not flag))
			if (CLng(rs1("IDALMACEN")) = CLng(rs("IDALMACEN"))) then flag = true
			rs1.MoveNext()
		wend
		if (flag) then
			'Tiene acceso al almacen como U o A.	
			ret = true
		else
			'Tiene otro acceso o no tiene.
			'Se averigua si el usuario cargo el pedido.					
			if (rs("CDUSUARIO") = session("Usuario")) then ret = true			
		end if
	end if	
	checkControlPM = ret
End Function
'----------------------------------------------------------------------------------------------------
Function clearArticuloPM()
	PM_idArticulo = 0	
	PM_dsArticulo = ""
	PM_cantidad = 0
	PM_articuloStock = 0	
	PM_saldo = 0	
End Function
'------------------------------------------------------------------------------------------------------
Function clearDetallePM()
	PM_idSector = 0	
	PM_idObra = 0
	PM_idBudgetArea = 0
	PM_idBudgetDetalle = 0
	PM_CantArticulos = 0
	PM_Tipo = 0
End Function
'----------------------------------------------------------------------------------------------
' Función:	  tieneAccesoTransferenciaPM
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  06/08/2013
' Objetivo:	  Controla si el usuario es Pañolero del Almacen para poder realizar pedidos a otros 
'			  almacenes de la empresa.
' Parametros: -
' Devuelve:	  true - false
'----------------------------------------------------------------------------------------------
Function tieneAccesoTransferenciaPM(pIdAlmacen)
	Dim strSQL, rs, flagControl
	flagControl = false 
	strSQL = " Select * FROM ( "&_
			 "		Select IDALMACEN from TBLALMACENESUSUARIO "&_
			 "		WHERE CDUSUARIO='" & session("usuario") & "'"&_
			 "			and NIVEL  = '" &  ALMACEN_USUARIO & "' and IDALMACEN = "& pIdAlmacen &" ) A "&_
			 "		INNER JOIN TBLALMACENES B ON  A.IDALMACEN = B.IDALMACEN AND B.ESTADO = "& ESTADO_ACTIVO			 
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.EoF then flagControl = true
	tieneAccesoTransferenciaPM = flagControl
End Function
'------------------------------------------------------------------------------------------------------
function readNextArticuloParamsPM()
	Dim index
	index = PM_DetalleActual
	PM_CantArticulos = 	GF_PARAMETROS7("rowTblArticulos_" & PM_DetalleActual,0,6)
	Call clearArticuloPM()
	readNextArticuloParamsPM = false	
	if PM_ArticuloActual < PM_CantArticulos then		
		PM_idArticulo = GF_PARAMETROS7("item" & index & "_" & PM_ArticuloActual,0,6)
		strSQL="select A.*, B.CDINTERNO from TBLARTICULOS A left join TBLARTICULOSDATOS B on A.IDARTICULO=B.IDARTICULO and B.IDALMACEN=" & PM_idAlmacen & " where A.IDARTICULO=" & PM_idArticulo
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			PM_idUnidad = rs("IDUNIDAD")
			PM_abreviaturaUnidad = getAbreviaturaUnidad(rs("IDUNIDAD"))
			PM_cdInterno = rs("CDINTERNO")
		end if		
		PM_dsArticulo = GF_PARAMETROS7("articuloItem" & index & "_" & PM_ArticuloActual &"_text","",6)
		PM_cantidad = GF_PARAMETROS7("amount" & index & "_" & PM_ArticuloActual,3,6)		
		readNextArticuloParamsPM = true		
	end if		
End Function
'------------------------------------------------------------------------------------------------------
Function readNextDetalleParamsPM()
	Call clearDetallePM()
	readNextDetalleParamsPM = false
	if (PM_Transferir = 0) then
		if PM_DetalleActual < PM_CantDetalle then	
			PM_Tipo = GF_PARAMETROS7("hidBoleanPartidaSector_" & PM_DetalleActual,0,6)
			if(PM_Tipo > 0) then							
				PM_idSector = GF_PARAMETROS7("cmbPartidaSector_" & PM_DetalleActual,0,6)
			else
				PM_idObra = GF_PARAMETROS7("cmbPartidaSector_" & PM_DetalleActual,0,6)
				PM_idBudgetArea = GF_PARAMETROS7("hiddenValueArea_" & PM_DetalleActual,0,6)		
				PM_idBudgetDetalle = GF_PARAMETROS7("cmbAreaDetalle_" & PM_DetalleActual,0,6)				
			end if			
			readNextDetalleParamsPM = true
		end if
	else
		if PM_DetalleActual = 0 then readNextDetalleParamsPM = true 		
	end if			
End Function
'--------------------------------------------------------------------------------------------------------
' Función:	  controlHeaderPM
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/08/2013
' Objetivo:	  Controla la cabecera del PM.
' Parametros: -
' Devuelve:	  true - false
'----------------------------------------------------------------------------------------------
Function controlHeaderPM()
	controlHeaderPM = false	
	if (PM_cdSolicitante = "") then call setError (SOLICITANTE_NO_EXISTE)
	if (PM_idAlmacen = 0) then call setError (ALMACEN_NO_EXISTE)
	if (isAuditorAL(PM_idAlmacen)) then Call setError (ALMACEN_NO_GUARDAR)
	if (PM_Transferir <> 0) then
		if (PM_idAlmacenDest = 0)then 
			setError(ALMACENDEST_NO_EXISTE)
		else
			if (PM_idAlmacenDest = PM_idAlmacen) then call setError (ALMACENDEST_IGUAL_ORIGEN)
		end if	
	end if
	if not hayError() then controlHeaderPM = true
End Function
'----------------------------------------------------------------------------------------------------
' Función:	  generarKeyDetalle
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/08/2013
' Objetivo:	  Genera una clave para usar en el Diccionario Detalle, dependiendo del detalle(Obra/Sector y Area/Detalle)
' Parametros: -
' Devuelve:	  clave [string]
'----------------------------------------------------------------------------------------------
Function generarKeyDetalle()		
	if (PM_idSector > 0 ) then	generarKeyDetalle = Trim("1" & PM_idSector)
	if (PM_idObra > 0 )then generarKeyDetalle = Trim("2" & PM_idObra & "|" & PM_idBudgetArea & "|" & PM_idBudgetDetalle)	
End function
'----------------------------------------------------------------------------------------------------
' Función:	  controlarArticuloPM
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/08/2013
' Objetivo:	  Controla los datos de un articulo del PM
' Parametros: -
' Devuelve:	  True - False
'----------------------------------------------------------------------------------------------
Function controlarArticuloPM()
	Dim flagPP,flagValePP, rsArt, strSQL, rtrn
	rtrn = true	
	if (controlarArticulo(PM_idArticulo))then
			VS_idAlmacenDest = PM_idAlmacenDest
			flagValePP = llevaPartidaVale(CODIGO_PM)
			flagPP = llevaPartida(PM_idArticulo)			
			if (PM_idBudgetArea=0 and flagPP and flagValePP) then
				setError(ARTICULO_REQUIERE_PP)
			else
				if (PM_idSector=0 and not flagPP and flagValePP) then 	
					setError(ARTICULO_REQUIERE_SEC)
				end if	
			end if		
	else
		Call setError(ARTICULO_NO_EXISTE)
	end if
	if hayError() then 
		rtrn = false
		PM_articuloError = PM_idArticulo
	end if	
	controlarArticuloPM = rtrn
End Function
'----------------------------------------------------------------------------------------------
' Función:	  controlarDetallePM
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/08/2013
' Objetivo:	  Controla los datos de un Detalle del PM, esto es la Obra/Sector y el Area/Detalle.
' Parametros: -
' Devuelve:	  True - False
'----------------------------------------------------------------------------------------------
'Controla los datos de un Detalle (Obra/Sector y Area/Detalle).
Function controlarDetallePM()
	Dim rtrn
	rtrn = true
	if PM_idObra > 0 then
		if ((PM_idBudgetArea > 0)and(PM_idBudgetDetalle > 0)) then
			Call getDivisionObraFull(PM_idObra, idDivisionObra, "")
			if (PM_idDivision <> idDivisionObra) then	Call setError(DIVISION_PM_VS_DIFF_OBRA)						
		else
			Call setError(FALTA_ASIGNAR_OBRA_SECTOR)
		end if
	else
		if (not validarSector(PM_idSector)) then call setError(FALTA_ASIGNAR_OBRA_SECTOR)
	end if	
	if hayError() then rtrn = false
	controlarDetallePM = rtrn
End Function 
'----------------------------------------------------------------------------------------------
' Función:	  controlarPM
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/08/2013
' Objetivo:	  Controla todos los datos del Pedido de Material
' Parametros: -
' Devuelve:	  True - False
'----------------------------------------------------------------------------------------------
Function controlarPM()	
	Dim flagDetalle, flagArticulo, cantProv, artDic, detDic
	Set artDic = Server.CreateObject("Scripting.Dictionary")
	Set detDic = Server.CreateObject("Scripting.Dictionary")
	controlarPM = false
	if (controlHeaderPM()) then
		flagDetalle = true
		PM_DetalleActual = 0
		PM_idDivision = getDivisionAlmacen(PM_idAlmacen)				
		while ((readNextDetalleParamsPM()) and (flagDetalle))
			if (PM_Transferir = 0) then						
				'No es transferencia, se controla el Sector/Obra , Area/Detalle y los Articulos
				if controlarDetallePM then
					myKey = generarKeyDetalle()
					if (not detDic.Exists(myKey)) then
						Call detDic.Add(myKey,"")
					else
						Call setError(DETALLE_DUPLICADO)
						flagDetalle = false
					end if
				else					
					flagDetalle = false
				end if
			end if			
			PM_ArticuloActual = 1			
			flagArticulo = true
			while ((readNextArticuloParamsPM()) and (flagDetalle) and (flagArticulo))					
				if (PM_idArticulo > 0) then
					if controlarArticuloPM() then
						if (not artDic.Exists(PM_idArticulo)) then
							Call artDic.Add(PM_idArticulo,"")
						else
							Call setError(ARTICULO_DUPLICADO)
						end if
					end if
				end if
				PM_ArticuloActual = PM_ArticuloActual + 1
				if ((PM_ArticuloActual = PM_CantArticulos)and(artDic.Count = 0)) then Call setError(POCOS_ARTICULOS)
				if hayError() then flagArticulo = false
			wend			
			artDic.RemoveAll
			if not flagArticulo then flagDetalle = false
			PM_DetalleActual = PM_DetalleActual +  1
		wend
	end if	
	Set artDic = nothing
	Set detDic = nothing
	PM_DetalleActual = 0	
	if not hayError() then controlarPM = true	
End Function
'---------------------------------------------------------------------------------------------

%>