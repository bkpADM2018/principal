<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<%
Const MIN_LINEAS = 10
Const MSG_ALERT_EXCEED = "ESTE PEDIDO CONSUME RECURSOS NO PRESUPUESTADOS.<br>"
Const MSG_ALERT_BONIF = "Se ha aplicado una bonificación.<br>"
			
Const ART_EXCLUSIVO_INV	=8835
Const ART_EXCLUSIVO_CTC	=11548

Const TODOS_ARTICULOS = 0

Const ITEM_ID = "ARTID_"
Const ITEM_ID_DIV = "ARTIDDIV_"
Const ITEM_DESC = "ARTDS_"
Const ITEM_UN_DESC = "ARTUNDS_"
Const ITEM_UN_DESC_DIV = "ARTUNDIV_"
Const ITEM_CANT = "CAN_"
Const ITEM_IMPP = "IMPP_"
	
dim rsDET, connDET, CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraCuentaDS, CAB_ObraImporte, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada
dim idPedido, CAB_idCotizacion, CAB_idPedido, CAB_cdPedido, CAB_idProveedor,CAB_idProveedorPCP, CAB_dsProveedor, CAB_fecEntrega, CAB_observaciones, CAB_importePesos, CAB_importeDolares
dim nroLinea, cantArt, index, accion, CAB_titulo, flagGrabar, CAB_idObra, tipoCambio, rsObras, CAB_estado, CAB_idContrato
dim IT_artID, IT_artDS, IT_cantidad, IT_importePesos, IT_importeDolares, IT_unidadID, IT_unidadDS, idCotizacionElegida, CAB_CdResponsable, CAB_DsResponsable
dim CAB_Moneda, CAB_IdDivision, CAB_ImportePlanilla, k, CAB_FechaBudget, IT_artBA, IT_artBD, esModificable, IT_unidadCD, dsUnidad
dim member1Cd, member2Cd, member3Cd, bonificacion, flagCTC
dim member1HK, member2HK, member3HK, member1FF, member2FF, member3FF, checkRtrn 
dim flagGrabarFirmasAuxiliares, monedaCargaReadOnly,lineClass, auxUser, auxTotal, subTotalImporte
Dim oDiccAFETotales,oDiccPICTotales,oDiccImpPP,oDiccAreasDetallesObras, flagDetallePresupuesto, useOrigen,isInPopUp
Dim verRemitos, descArticulo, pAbrev, auxremito, idremito, cdInterno, rsAlmacenes, rsArt,dicArtUCC, dicError,itemsNecesitanAFE
Dim oDiccPartidaExcedida,oDiccPartidaNoExcedida,strObserv, errAuditoria

Set oDiccAFETotales  = createObject("Scripting.Dictionary")
Set oDiccPICTotales  = createObject("Scripting.Dictionary")
Set oDiccImpPP       = createObject("Scripting.Dictionary")
Set oDiccAreasDetallesObras = createObject("Scripting.Dictionary")
flagGrabarFirmasAuxiliares = false
Set oDicNuevaPartida = Server.CreateObject("Scripting.Dictionary")
Set dicArtUCC = Server.CreateObject("Scripting.Dictionary")
Set dicError = Server.CreateObject("Scripting.Dictionary")
Set itemsNecesitanAFE = Server.CreateObject("Scripting.Dictionary")
Set oDiccPartidaExcedida = Server.CreateObject("Scripting.Dictionary")
Set oDiccPartidaNoExcedida = Server.CreateObject("Scripting.Dictionary")
'-----------------------------------------------------------------------------------------------
Function cargarDiccionario(pIdObra, pIdPic)
	Dim strSQL,conn,rs,rsaux,ultimaArea

	strSQL = "SELECT * FROM tbldatosafe WHERE idobra = " & pIdObra
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	while not rs.EOF
		Call agregarImporteDic(oDiccAFETotales,rs("IDAREA")&"-"&rs("IDDETALLE"),cdbl(rs("IMPORTEDOLARES")))
		rs.MoveNext
	wend
	
	strSQL = "SELECT idarea,iddetalle ,sum(d.importedolares) importe FROM tblctzcabecera c "
	strSQL = strSQL & " INNER JOIN tblctzdetalle d ON c.idcotizacion = d.idcotizacion "
	strSQL = strSQL & " WHERE c.idobra = " & pIdObra & " and c.idCotizacion<>" & pIdPic
	strSQL = strSQL & " group by idarea,iddetalle "
	'call logdebug(strSQL)
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.EOF) then ultimaArea = rs("IDAREA")
	while not rs.EOF
		'totales de los detalles
		Call agregarImporteDic(oDiccPICTotales,rs("IDAREA")&"-"&rs("IDDETALLE"),cdbl(rs("importe")))
		rs.MoveNext
	wend	

	'guardo los Areas-Detalles de la Obra Para saber cuales existen
	strSQL = "select * from tblbudgetobras where idobra = " & pidObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rsaux, "OPEN", strSQL)
	while not rsaux.EOF
		Call agregarImporteDic(oDiccAreasDetallesObras,rsaux("IDAREA")&"-"&rsaux("IDDETALLE"),cdbl(rsaux("DLBUDGET")))
		rsaux.MoveNext
	wend
	
End Function
'-----------------------------------------------------------------------------------------------
Function agregarImporteDic(byref pDic,pKey,pValue)
	if not pDic.Exists(pKey) then
		Call pDic.Add(pkey,pValue)
	else
		pDic.Item(pkey) = cdbl(pDic.Item(pkey)) + cdbl(pValue)
	end if
end Function
'-----------------------------------------------------------------------------------------------
Function errorAcceso() 
	response.redirect "comprasAccesoDenegado.asp"
End Function
'-----------------------------------------------------------------------------------------------
Function checkAjustar(pIdCotizacion, pIdArticulo, pArea, pDetalle)
    Dim rtrn, rs
    
    rtrn = ""
    
    'Verifico que el PIC este en un estado correcto o que el articulo sea la diferencia de cambio.
    if ((CAB_estado = CTZ_FIRMADA) or ((CAB_estado = CTZ_FACTURADA) and (pIdArticulo = CTZ_ITEM_DIFF_CAMBIO))) then 
        'Verifico que no haya ajustes en curso para el articulo- partida
        if (existeAjusteCotizacionArticulo(pIdCotizacion,pIdArticulo, pArea, pDetalle)) then
            rtrn = "No se puede ajustar. Hay otro ajuste pendiente para este articulo y partida. Eliminelo y ajuste nuevamente."
        else
            'Verifico que no se haya recibido en pañol nada de este articulo-partida.
            'Call executeQuery(rs, "OPEN", "Select * from TOEPFERDB.TBLREMPIC where IDPIC=" & pIdCotizacion & " and IDARTICULO=" & pIdArticulo & " and IDAREA=" & pArea & " and IDDETALLE=" & pDetalle)
            'if (not rs.eof) then rtrn="Ya se recibieron articulos en el pañol. No puede ajustarse el articulo."                
        end if            
    else        
        rtrn = "El " & ctz_docCode & " no acepta ajustes en este momento."        
    end if
    checkAjustar = rtrn
    
End Function
'-----------------------------------------------------------------------------------------------
Function puedeModificarImportes()
	puedeModificarImportes = true
	if ((CAB_estado = CTZ_FACTURADA) or (CAB_estado = CTZ_ANULADA)) then puedeModificarImportes = false
End Function
'-----------------------------------------------------------------------------------------------
Function puedeModificarCantidades()
	Dim strSQL, rs, conn
	puedeModificarCantidades = true	
	strSQL= "Select IDPIC from TBLREMPIC where IDPIC=" & CAB_idCotizacion
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if ((not rs.eof) or (CAB_estado = CTZ_ANULADA)) then	puedeModificarCantidades = false			
End Function
'-----------------------------------------------------------------------------------------------
sub RedimVarialbles(pCant)
	redim IT_cantidad(pCant)
	redim IT_importePesos(pCant)
	redim IT_importeDolares(pCant)
	redim IT_unidadID(pCant)
	redim IT_unidadCD(pCant)
	redim IT_unidadDS(pCant)	
	redim IT_artDS(pCant)
	redim IT_artID(pCant)
	redim IT_artBA(pCant)
	redim IT_artBD(pCant)
end sub

'**************************************************************************************************
'**----------------------------------------------------------------------------------------------**
'**----------------------------------------------------------------------------------------------**
'**---------------------------       SECTOR CONTROLES       -------------------------------------**
'**----------------------------------------------------------------------------------------------**
'**----------------------------------------------------------------------------------------------**
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	--/--/--
' Objetivo:	
'			Controla la cotizacion del Pic
' Devuelve:
'			True/False
' Modificaciones:
'			25/10/10 - GFG
'			15/11/10 - GFG
'--------------------------------------------------------------------------------------------------
Function controlarCotizacion()
	Dim ret, desc
	ret= false
	itemsNecesitanAFE.RemoveAll
	
	if (CAB_IdDivision <> SIN_DIVISION) then
		if (CAB_idObra <> 0)and(CAB_idObra <> OBRA_GEID) then			
			if (CAB_idDivision <> CAB_ObraDivID) then setError(DIVISION_PCT_DIFF_OBRA)			
		end if
		if (not hayError()) then
			'Se controla si tiene acceso		
			if (isAdmin(CAB_IdDivision) or isUser(CAB_IdDivision)) then
				'Se controla el proveedor	
				if (controlarProveedor(CAB_idProveedor)) then
			        'Controlo la norma de auditoria. Solo si es compra Directa
			        if (controlarFirmas()) then
				        'Controlar la norma de auditoria solo si el pago NO pertenece a un contrato.
				        if (not flagCTC) then								
					        if (idPedido > 0) then
						        if (CLng(CAB_idProveedor) = CLng(CAB_IdProveedorPCP)) then
						            'controlo si el pic tiene el mismo proveedor que el pct debe tener la misma partida presupuestaria.					            
							        if (pct_idArea > 0) then	
								        'los proveedores son los mismos todas las partidas presupuestarias deben ser igual al del pct
								        for i = 0 to ubound(IT_artID)
									        if (IT_artID(i) <> 0) then
										        if (cdbl(IT_artBA(i)) <> cdbl(pct_idArea)) then setError(PP_PIC_DISTIN_PCT)
										        if ((cdbl(IT_artBD(i)) <> cdbl(pct_idDetalle)) and (cdbl(pct_idDetalle) > 0)) then setError(PP_PIC_DISTIN_PCT)																																							
									        end if												
								        next										
							        end if 
						        else
							        Call setError(PROV_NO_COINCIDE_PCT)		
						        end if 
					        end if
					        if (not hayError()) then
						        if (controlPCP()) then
							        'La fecha de entrega es correcta, se controlan los articulos							
							        ret = controlarDetalle()
							        'En el detalle se agrego un control para saltear el control de auditoria, recien aca se valida si el control se hace o no.
							        if (flagGrabarFirmasAuxiliares) then setWarning(errAuditoria)
						        else								
							        Call setError(IMPORTE_NO_COINCIDE)						
						        end if
					        end if
				        else
					        ret= true
				        end if						
			        end if					
				end if
			else
				setError(USUARIO_NO_AUTORIZADO)
			end if
		end if	
	else
		setError(DIVISION_NO_EXISTE)
	end if	
	controlarCotizacion = ret
End Function
'-----------------------------------------------------------------------------------------------
Function controlarFirmas()
	Dim	rtrn
	
	rtrn = true
	
	if (member1Cd = "") then		
		rtrn = false
		Call setError(SOLICITANTE_NO_EXISTE)
	end if
	if (member2Cd = "") then		
		rtrn = false
		Call setError(AUTORIZANTE_NO_EXISTE)
	end if
	controlarFirmas = rtrn
	
End Function
'-----------------------------------------------------------------------------------------------
'Controla que el importe de la cotizacion coincida con el importe de la cotizacion seleccioanda en la Planilla Comparativa de Precios(PCP)
Function controlPCP()
	Dim strSQL, rsPCP, rsPCT, rsCTZ, con, importeCompra, myTotalPesos, myTotalDolares
	
	controlPCP = true	
'	response.write "CONTROL PCP!<br> "  
	if (pct_idPedido > 0) then	
		'if (CLng(CAB_idProveedorPCP) = CLng(CAB_idProveedor)) then
			controlPCP = false
			Call loadImporteAcumuladoPIC(pct_idPedido, CAB_idCotizacion, CAB_idProveedor, false, myTotalPesos, myTotalDolares)
			importeCompra = CAB_importePesos + myTotalPesos 'Asumo Pesos	
			if (CAB_Moneda = MONEDA_DOLAR) then importeCompra = CAB_importeDolares + myTotalDolares		
			'Controlo que el importe de la compra sea exactamente igual.				
			'Se comparan los importes por medio de restas para evitar problemas con diferencias menores a 1 centavo (paso una vez!)						
			if (Round(CDbl(CAB_ImportePlanilla), 0) => Round(CDbl(importeCompra), 0)) then controlPCP = true									
		'end if
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	--/--/--
' Objetivo:	
'			Controla los items de la cotizacion
' Devuelve:
'			True/False
' Modificaciones:
'			25/10/10 - GFG
'--------------------------------------------------------------------------------------------------
Function controlarDetalle()
	Dim ret, index, cantArt, listaART

	index = 0
	cantArt = 0
	ret = true
	listaART = "0"
	
	'Se controla artículo por artículo.	
	while ((index < UBound(IT_artID)) and (ret))
		if (IT_artID(index) > 0) then				    				
			ret = false			
			'Control del articulo
			if (controlarArticulo(IT_artID(index))) then
			    if (esArticuloElegiblePIC(IT_artID(index)))and(esArticuloElegibleObra(IT_artID(index),CAB_idObra)) then
				    'Controla la cantidad
				    if (IT_cantidad(index) > 0) then						
					    if (IT_importePesos(index) = "") then IT_importePesos(index) = 0					
					    'Controla que el articulo no este duplicado dentro del mismo area-detalle								
					    if (controlArticulosDuplicados(index)) then					    
						    Call agregarImporteDic(oDiccImpPP,IT_artBA(index)&"-"&IT_artBD(index),IT_importeDolares(index))
						    'Si es el artículo especial para contratos, entonces no se requieren firmas adicionales.
						    'if (CLng(IT_artID(index)) = ART_EXCLUSIVO_CTC) then flagGrabarFirmasAuxiliares = false
						    ret = true						
					    else  
						    setError(DETALLE_DUPLICADO)
					    end if
				    else							
					    setError(CANTIDAD_NO_EXISTE)
				    end if
				else
				    'El articulo no es elegible.
				    setError(ARTICULO_NO_ELEGIBLE)
				    dicError.Add IT_artID(index), IT_artID(index)
				end if
			else
				'El articulo no existe.
				dicError.Add IT_artID(index), IT_artID(index)
			end if
			cantArt = cantArt +1
			'Armo la lista de articulos para controles futuros.
			listaART = listaART & ", " & IT_artID(index)
		end if
		index = index+1
	wend
	if (cantArt = 0) then setError(POCOS_ARTICULOS)
	'Se controlan las reglas de Bienes de uso.	    	 
	ret = controlarBU(listaART, cantArt, CAB_idObra)
	if (ret)  then	    	    
		'Se controla si el precio de los items tiene sentido con los datos históricos
		Set dicArtUCC = controlarPrecioArticulo(MONEDA_PESO,IT_artID,IT_importePesos,IT_importeDolares,IT_cantidad,CAB_IdDivision, CAB_IDCotizacion)
		if (dicArtUCC.count > 0 ) then Call setWarning(PRECIO_DIFIERE_ULTIMO_REGISTRO)			
		ret = false
		if (controlPartidaPresupuestaria(CAB_idObra, CAB_idPedido, CAB_idCotizacion, oDiccImpPP)) then ret = true
	end if	
	
	controlarDetalle = ret
End Function
'-----------------------------------------------------------------------------------------------
Function controlarBU(pListaART, pCantArt, pIdObra)

    Dim rs, strSQL, flagBU, ret 
        
    flagBU=false
    ret = false
    '1ro - Verifico que no se mezclen Bienes de uso con otros articulos.
    strSQL="Select count(*) CANT from TBLARTICULOS where IDARTICULO in (" & pListaART & ") and BIENUSO='" & ES_BIEN_DE_USO & "'"    
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    if (not rs.eof) then        
        if ((CLng(rs("CANT")) <> pCantArt) and (CLng(rs("CANT")) <> 0)) then 
            setError(BU_MEZCLA_ARTICULOS)
            flagBU=true
        end if
    end if    
    '2do - Controlo que la partida solo sea de inversiones.
    if (flagBU and (pIdObra <> 0)) then
        strSQL="Select * from TBLDATOSOBRAS where IDOBRA=" & pIdObra & " and TIPOGASTO='" & OBRA_TIPO_INVERSION & "'"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
        if (rs.eof) then setError(BU_PARTIDA_INCORRECTA)        
    end if
    if (not hayError()) then ret = true
    controlarBU = ret
End Function
'-----------------------------------------------------------------------------------------------
Function controlArticulosDuplicados(pIdArticuloIndex)
	dim rtrn
	rtrn = true	
	for i = pIdArticuloIndex+1 to ubound(IT_artID)-1		
		if (IT_artID(i) = IT_artID(pIdArticuloIndex)) then
			if ((IT_artBA(i) = IT_artBA(pIdArticuloIndex)) and (IT_artBD(i) = IT_artBD(pIdArticuloIndex))) then				
				rtrn = false
			end if
		end if
	next
	controlArticulosDuplicados = rtrn
End Function
'-----------------------------------------------------------------------------------------------
'Controlar que la partida presupuestaria exista.
Function controlPartidaPresupuestaria(pIdObra, pPedido, idPIC, byref pDiccPP)
	Dim strSQL, conn, rsBudget
	ret = true			
	index = 0
    if (pIdObra <> OBRA_GEID) then
	    'Se controla que la partida presupuestaria sea valida.												
	    while ((index < UBound(IT_artID)) and (ret))
			    'Verifico si esta utilizando una partida preexistente		    
		        if ((IT_artID(index) <> 0) and (((IT_artBD(index) <> "") and (Trim(IT_artBD(index)) <> "0")) or ((IT_artBA(index) <> "") and (Trim(IT_artBA(index)) <> "0")) or (pIdObra <> 0)))then		    		    
		            if (not oDiccAreasDetallesObras.Exists(IT_artBA(index) & "-" & IT_artBD(index))) then			
		                'La partida no existe...!				
					    Call setError(BUDGET_NO_EXISTE)
					    ret = false				
				    end if			
		        end if
		    index=index+1
	    wend			
	    if (ret) then 
	        'Se notifica que falta info presupuestaria	
	        if (pIdObra = 0) then setWarning(OBRA_NO_SELECCIONADA)			
	        ret = controlPresupuestarioDetalle(pIdObra, pPedido, idPIC, pDiccPP)
	    end if
    else
        if (Trim(CAB_observaciones) = "") then Call setError(SM_OBS_REQUERIDAS)
    end if
	controlPartidaPresupuestaria = ret	
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	09/05/2012
' Objetivo:	
'			Deteminar si alguno de los artículos del PIC indicado pertenece a una categoría de mantenimiento,
' Parametros:
'			idPIC [int]	Id del PIC a analizar
' Devuelve:
'			True/False
Function hayArticuloMantenimiento(pPedido, arts)
	
	Dim strSQL, rs, lista
	
	
	lista = ""
	for index = 0 to ubound(arts)-1
		lista = lista & arts(index) & ", "
	next
	lista = left(lista, len(lista)-2)
	
	strSql=	"Select * from TBLARTCATEGORIAS where IDCATEGORIA in (" & _
			"	Select IDCATEGORIA from TBLARTICULOS where IDARTICULO in (" & lista & ")"
	if (pPedido > 0) then			
		strSql= strSql & " or IDARTICULO in (Select IDARTICULO from TBLCTZDETALLE D inner join TBLCTZCABECERA C on D.IDCOTIZACION=C.IDCOTIZACION where IDPEDIDO=" & pPedido & ")"
	end if						
	strSql= strSql & ") and ESMANTENIMIENTO='" & TIPO_AFIRMACION &"'"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	hayArticuloMantenimiento = not rs.eof
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	--/--/--
' Objetivo:	
'			Controla el detalle del PIC
' Parametros:
'			pIdObra 	[int] 
'			pDicPP		[dictionary]	Diccionario con las PP de todos los items agrupadas
' Devuelve:
'			TRUE/FALSE
' Modificaciones:
'			16/11/10 - GFG
'--------------------------------------------------------------------------------------------------
Function controlPresupuestarioDetalle(pIdObra, pPedido, idPIC, byRef pDicPP)
	Dim myImporte,myLimite,myArea,myDetalle,rtrn,aux, myGastoTotal, rsDesc
	Dim arrKeys,arrImporte, limiteExtra,totalDolaresControlAFE	,myAreaDetalle
	rtrn = true
	myImporte = 0
	
	flagBGT = false
	flagDetallePresupuesto = false
	arrKeys = pDicPP.Keys()
	arrImporte = pDicPP.Items()
	i=0		
	totalDolaresControlAFE = 0
	if (pIdObra = 0) then
		'Totalizo los items del PIC compra directa.
		if (hayArticuloMantenimiento(pPedido, IT_artID)) then
			for index = 0 to ubound(IT_importeDolares)			
				totalDolaresControlAFE = totalDolaresControlAFE + IT_importeDolares(index) - round((IT_importeDolares(index) * bonificacion)/100,0)			
			next		
			if (necesitaAFE(pIdObra,pPedido, idCotizacionElegida,totalDolaresControlAFE,0,0)) then setWarning(PIC_NECESITA_AFE)
		end if
	else
		while ((i <= ubound(arrKeys)) and (rtrn))		
			aux = split(arrKeys(i),"-")
			myArea = aux(0)
			myDetalle = aux(1)

				myImporte = cdbl(arrImporte(i))						
				if (necesitaAFE(pIdObra, pPedido, idCotizacionElegida, myImporte,myArea,myDetalle)) then 
					setWarning(PIC_NECESITA_AFE)
					Call itemsNecesitanAFE.add(arrKeys(i),arrKeys(i))
				end if
				'calculo si excedio el limite de la PP
				flagDetallePresupuesto = true
				myGastos = calcularGastosObra(MONEDA_DOLAR, pIdObra, myArea, myDetalle, false)
				Set rsDesc = obtenerDescripcionCompletaDetalle(pIdObra, myArea, myDetalle) 
				myAreaDetalle = myArea & "-" &  myDetalle		
				'Se asume que el importe del PIC se debe añadir a los gastos ya grabados de la obra, si se esta modificando un PIC (IDPIC <> 0) entonces los gastos calculados
				'en la funcion calcularGastosObra con el último parametro en false, ya incluyen este importe
				myGastoTotal = myGastos
				if (idPIC = 0) then myGastoTotal = myGastos + myImporte
				if(oDiccAreasDetallesObras.Item(arrKeys(i)) < myGastoTotal) then	
					Call setError(AVISO_BGT_EXCEDIDO)
					if not oDiccPartidaExcedida.Exists(myAreaDetalle) then Call loadDiccPartida(oDiccPartidaExcedida,myAreaDetalle, rsDesc("DSDETALLE"),oDiccAreasDetallesObras.Item(arrKeys(i)), myGastoTotal)
				else
					if not oDiccPartidaNoExcedida.Exists(myAreaDetalle) then Call loadDiccPartida(oDiccPartidaNoExcedida,myAreaDetalle, rsDesc("DSDETALLE"),oDiccAreasDetallesObras.Item(arrKeys(i)), myGastoTotal)
				end if									
                'Controlo que la partida no se encuentre en proceso de reasignacion o ajuste (se congelan los importes hasta que no se confirmen)
                if (tieneReasignacionAjusteActivo(pIdObra, myArea, myDetalle)) then Call setError(BUDGET_REASIGNACION_EN_PROCESO)
                if (hayError()) then rtrn = false
			i=i+1
		wend
	end if
	
	controlPresupuestarioDetalle = rtrn
	
End Function

'-----------------------------------------------------------------------------------------------
' Autor: 	Ajaya Nahuel	
' Fecha: 	05/12/12
' Objetivo:	
'			Carga en el Dictionary el Area - Detalle junto a los valores del presupuesto del Pic
' Parametros:
'			pDicc 			[Dictionary] 
'			pKey			[string]	Area + detalle
'			pPresupuesto	[int]
'			pGasto			[int]
' Devuelve:		-
Function loadDiccPartida(pDicc,pKey, DescPartida, pPresupuesto, pGasto)
	Dim pSaldos,strValues
	pSaldos =  pPresupuesto - pGasto	
	strValues = "Partida: " & pKey & " " & Left(DescPartida, 35) & " | Presupuesto Asignado: "& getSimboloMoneda(MONEDA_DOLAR) &"&nbsp;"& GF_EDIT_DECIMALS(pPresupuesto,2) & " | Saldo: " & getSimboloMoneda(MONEDA_DOLAR) &"&nbsp;"&GF_EDIT_DECIMALS(pSaldos,2)
	Call pDicc.Add(pKey,strValues)	
End Function
'-----------------------------------------------------------------------------------------------
'Funcion:       AlrtaCDEmailControl()
' Autor: 	    Jonathan G. Costilla
' Fecha: 	    04/08/2016
' Objetivo:	    Alerta. Enviar meil en caso de que se supere el monto de compra en el mes
' Parametros:   No recibe
Function AlrtaCDEmailControl()
    Dim mmtoDesde,retCant,retImporte,PIC_Momento,auxSendMail, myMoneda
    Dim auxAsunto,auxDescripcion,auxOrigen,auxDestino,auxDsProveedor
'   Detectar si es compra directa.
    IF((CAB_idPedido = 0) AND (PIC_idContrato = 0) and (CAB_idObra=0)) THEN
'       Recibo un IDProveedor se busca todas las compras asociadas a ese proveedor
        mmtoDesde = GF_DTEADD(session("mmtodato"),-30,"D")
        Call totalizarComprasDirectasProveedor(CAB_idProveedor, CAB_IdDivision, mmtoDesde, session("mmtodato"), CAB_Moneda, retCant,retImporte)
'       Determinar si se supero el monto de compras Directas.SI? Enviar Mail de alerta
        auxSendMail = getPICAuthorizationType(CAB_idPedido, PIC_idContrato, CAB_idProveedor, CDbl(retImporte)/100, CAB_Moneda)
        If ((auxSendMail = PIC_TYPE_PURCHASE_X_MEDIUM) OR (auxSendMail = PIC_TYPE_PURCHASE_LARGE)) Then
            auxDsProveedor = CAB_idProveedor&"-"&getDescripcionProveedor (CAB_idProveedor)
            limiteFirmaGte = CDbl(getValorNorma("VLMAXCD"))*100 'Limite Firma Gerente, Importe maximo autorizado para realizar Compras Directas
            unidadCD = getUnidadNorma("VLMAXCD")
            myImporte = retImporte
            myMoneda = unidadCD
	        if (CAB_Moneda <> unidadCD) then
		        if (CAB_Moneda = MONEDA_PESO) then	
			        myImporte = CDbl(myImporte) / getTipoCambio(MONEDA_DOLAR, "")
			        myMoneda = TIPO_MONEDA_DOLAR
		        else
			        myImporte = CDbl(myImporte) * getTipoCambio(MONEDA_DOLAR, "")
                    myMoneda = TIPO_MONEDA_PESO
		        end if
	        end if	
	        myImporte = myMoneda & " " & GF_EDIT_DECIMALS(myImporte,2)                                        
			limiteFirmaGte = myMoneda & " " & GF_EDIT_DECIMALS(limiteFirmaGte, 2)
'           Cargo Mail            
            auxAsunto = "Sistema de Compras Web - Alerta de Pedidos Fraccionados"
            auxDescripcion= "Se ha superado el monto maximo mensual de compras directas para el proveedor " & UCase(auxDsProveedor)& vbCrLf & vbCrLf &_ 
                            "Fecha Desde:                   " & GF_FN2DTE(LEFT(mmtoDesde,8)) & vbCrLf &_
                            "Fecha Hasta:                    " & GF_FN2DTE(LEFT(session("mmtodato"),8)) & vbCrLf &_
                            "Importe Total:                " & myImporte & vbCrLf &_
                            "Importe autorizado:     " & limiteFirmaGte
            auxOrigen = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_SENDER)
            auxDestino = getTaskMailList(TASK_PROV_MAIL_ALERT, MAIL_TASK_INFO_LIST)         
'           Envia Mail
            Call GP_ENVIAR_MAIL(auxAsunto,auxDescripcion,auxOrigen,auxDestino)                       
        End If
    END IF
End Function
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
nroLinea = 0
index = 0
monedaCargaReadOnly = false
flagCTC = false

idPedido = GF_Parametros7("idPedido",0,6)
idCotizacionElegida = GF_Parametros7("idCotizacionElegida",0,6)
accion = GF_PARAMETROS7("accion","",6)
nroLinea = GF_PARAMETROS7("nroLinea",0,6)
bonificacion = GF_PARAMETROS7("bonificacion", 2, 6)
useOrigen = GF_PARAMETROS7("useOrigen", 0, 6)
isInPopUp = GF_PARAMETROS7("isInPopUp", 0, 6)
uploadFilesName = GF_PARAMETROS7("uploadFilesName", "", 6)
'si muestra remitos
verRemitos = GF_PARAMETROS7("verRemitos","",6)

if bonificacion > 0 then 
	Call setWarning(BONIFICACION_PENDIENTE)
end if	
CAB_importePesos = 0
CAB_importeDolares = 0
CAB_observaciones = ""

tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
'Se inicializan las variables de artículos que estarán vacias inicialmente
RedimVarialbles(MIN_LINEAS)

Call GP_CONFIGURARMOMENTOS

CAB_estado = CTZ_PENDIENTE
CAB_titulo = "Cotizacion Elegida"
if nroLinea = 0 then 'Leer desde base
	'Leer Cabecera
	if idCotizacionElegida > 0 then		
		'Se paso como parametro el ID de una cotizacion cargada en el sistema, se debe ir a la base de cotizaciones cargadas.
		strSQL="SELECT * from TBLCTZCABECERA where IDCOTIZACION=" & idCotizacionElegida
		'Response.Write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then			
			if (rs("IDPEDIDO") = "0") then
				Call comprasControlAccesoCM(RES_CD)
			else
				Call comprasControlAccesoCM(RES_CC)
			end if
			
			CAB_idCotizacion = rs("IDCOTIZACION")
			CAB_idPedido = rs("IDPEDIDO")			
			Call initHeaderDB(CAB_idPedido)						
			CAB_cdPedido = pct_cdPedido					
			if CAB_cdPedido = "" then CAB_cdPedido = "Sin Pedido"				
			if (CAB_idPedido = 0) then CAB_titulo = "Compra Directa"
			CAB_idProveedor = rs("IDPROVEEDOR")
			CAB_dsProveedor = getDescripcionProveedor(CAB_idProveedor) 						
			CAB_fecEntrega = rs("FECHAENTREGA")
			CAB_observaciones = rs("OBSERVACIONES")
			CAB_observaciones = Replace(CAB_observaciones, MSG_ALERT_EXCEED, "")
			CAB_importePesos = CDbl(rs("IMPORTEPESOS"))
			CAB_importeDolares = CDbl(rs("IMPORTEDOLARES"))
			CAB_Moneda = rs("CDMONEDA")
			CAB_idObra = rs("IDOBRA")
			if(CAB_idObra > 0)then 				
				if(InStr(CAB_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO) > 0)then
					strObserv = Split(CAB_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO)
					CAB_observaciones = Left(CAB_observaciones,Len(strObserv(0)))			
				end if	
			end if	
			CAB_IdDivision = rs("IDDIVISION")
			CAB_estado = rs("ESTADO")
			CAB_idContrato = CLng(rs("IDCONTRATO"))
			tipoCambio = CDbl(rs("TIPOCAMBIO"))
			if (CAB_idObra <> OBRA_GEID) then
                Call loadDatosObra(CAB_idObra, CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada,CAB_CdResponsable, CAB_DsResponsable)
                if (CAB_ObraCD = "") then CAB_ObraCD = "Sin Partida"			
            else
                CAB_ObraCD = OBRA_GECD
                CAB_ObraDS = OBRA_GEDS
            end if
			'Leer detalles
			strSQL="SELECT * from TBLCTZDETALLE where IDCOTIZACION=" & idCotizacionElegida
			call executeQueryDb(DBSITE_SQL_INTRA, rsDET, "OPEN", strSQL)
			'cantArt = rsDET("Cantidad")			
			cantArt = rsDET.RecordCount
			if (CInt(cantArt) = 0) then cantArt = MIN_LINEAS
			RedimVarialbles(cantArt)
			while not rsDET.eof
				IT_artID(index) = rsDET("IDARTICULO")
				call getArticuloFull(IT_artID(index), IT_artDS(index), IT_unidadDS(index))
				IT_cantidad(index) = rsDET("CANTIDAD")
				IT_importePesos(index) = cdbl(rsDET("IMPORTEPESOS"))
				IT_importeDolares(index) = cdbl(rsDET("IMPORTEDOLARES"))
				IT_unidadID(index) = rsDET("IDUNIDAD")
				IT_artBA(index) = rsDET("IDAREA")
				IT_artBD(index) = rsDET("IDDETALLE")
				index = index + 1
				rsDET.movenext
			wend		
			strSQL = "Select * from TBLCTZFIRMAS where IDCOTIZACION=" & CAB_idCotizacion & " order by SECUENCIA"
			call executeQueryDb(DBSITE_SQL_INTRA, rsFirmas, "OPEN", strSQL)
		    if (not rsFirmas.eof) then
		        member1Cd = rsFirmas("CDUSUARIO")			
		        member1 = getUserDescription(member1Cd)
		        rsFirmas.MoveNext()
		        if (not rsFirmas.eof) then member2Cd = rsFirmas("CDUSUARIO")		        
		    end if		      
			' SI EL PIC TIENE UN CONTRATO ASOCIADO SIGNIFICA QUE ES UN PAGO DE CONTRATO
			flagCTC = false
			if(CAB_idContrato > 0)then flagCTC = true			
		else
			Call errorAcceso()										
		end if	
	else	
		CAB_idCotizacion = 0	
		CAB_idContrato = 0
		'No hay cotizacion elegida.
		if (idPedido > 0) then		
			Call comprasControlAccesoCM(RES_CC)			
			'Esta queriendo cargar una cotizacion nueva para una comparativa o un concurso de precios.
			'Cabecera
			monedaCargaReadOnly = true			
			strSQL="SELECT CAB.*, PCP.FECENTREGA, PCP.CDMONEDA from TBLPCTCABECERA CAB inner join TBLPCPDETALLE PCP on PCP.IDPEDIDO=CAB.IDPEDIDO and PCP.IDPROVEEDOR=CAB.IDPROVEEDOR where CAB.IDPEDIDO=" & idPedido						
			call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if not rs.eof then
				CAB_idPedido = rs("IDPEDIDO")
				Call initHeaderDB(CAB_idPedido)			
				if (not checkControlPCT()) then Call errorAcceso()
				CAB_cdPedido = pct_cdPedido				
				CAB_idObra = pct_idObra						
				CAB_IdDivision = pct_idDivision				
				if (CAB_idObra <> OBRA_GEID) then
                    Call loadDatosObra(CAB_idObra, CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada, CAB_CdResponsable, CAB_DsResponsable)
                    if (CAB_ObraCD = "") then 
					    CAB_ObraCD = "Sin Partida"				
				    else					
					    CAB_IdDivision = CAB_ObraDivID
				    end if	
                else
                    CAB_ObraCD = "GEID"
                    CAB_ObraDS = "GASTO ESPECIAL DE IIMPUTACION DIRECTA"
                end if
				CAB_Moneda = rs("CDMONEDA")
				CAB_idProveedor = rs("IDPROVEEDOR")
				CAB_dsProveedor = getDescripcionProveedor(CAB_idProveedor)
				CAB_fecEntrega  = rs("FECENTREGA")
				member1Cd = pct_cdSolicitante	
				member1Ds = getUserDescription(pct_cdSolicitante)
			else					
				Call errorAcceso()
			end if	
		else		
			Call comprasControlAccesoCM(RES_CD)		
			'No hay pedido ni cotizacion elegidam, se quiere cargar una nueva compra directa.		
			CAB_fecEntrega = Left(session("MmtoDato"), 8)			
			CAB_titulo="Compra directa"
			CAB_cdPedido = "Sin pedido"
			CAB_idObra = 0
			CAB_IdDivision = getDivisionID(CODIGO_EXPORTACION)
			CAB_idProveedor = 0
			CAB_ObraCD = "Sin Partida"
			CAB_idPedido = "0"				
			CAB_titulo = "Compra Directa"		
		end if	
	end if		

else 'Leer desde pagina	
	CAB_idPedido = GF_PARAMETROS7("CAB_idPedido",0,6)		
	if (CAB_idPedido > 0) then
		Call comprasControlAccesoCM(RES_CC)
	else
		Call comprasControlAccesoCM(RES_CD)
		CAB_titulo="Compra directa"
	end if
	'Leer Cabecera
	CAB_idCotizacion = GF_PARAMETROS7("CAB_idCotizacion",0,6)		
	Call initHeaderDB(CAB_idPedido)		
	CAB_idObra = GF_PARAMETROS7("CAB_idObra",0,6)	
	if (CAB_idObra <> OBRA_GEID) then
        Call loadDatosObra(CAB_idObra, CAB_ObraCD, CAB_ObraDS, CAB_ObraDivID, CAB_ObraDivDS, CAB_ObraImporte, CAB_FechaBudget, CAB_ObraMoneda, CAB_ObraFechaInicio, CAB_ObraFechaFin, CAB_ObraFechaAjustada, CAB_CdResponsable, CAB_DsResponsable)		
        if (CAB_ObraCD = "") then CAB_ObraCD = "Sin Partida"	
    else
        CAB_ObraCD = OBRA_GECD
        CAB_ObraDS = OBRA_GEDS
    end if
	tipoCambio = GF_PARAMETROS7("tipoCambio", 4,6)
	CAB_cdPedido = GF_PARAMETROS7("CAB_cdPedido","",6)
	CAB_idProveedor = GF_PARAMETROS7("CAB_idProveedor",0,6)	
	CAB_dsProveedor = GF_PARAMETROS7("CAB_dsProveedor","",6)
	CAB_fecEntrega = Left(session("MmtoDato"), 8)
	CAB_observaciones = GF_PARAMETROS7("CAB_observaciones","",6)
	CAB_observaciones = Replace(CAB_observaciones, MSG_ALERT_EXCEED, "")
	if(CAB_idObra > 0)then
		if(InStr(p_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO) > 0)then	
			strObserv = Split(CAB_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO)
			CAB_observaciones = Left(CAB_observaciones,Len(strObserv(0)))
		end if	
	end if
	CAB_Moneda = GF_PARAMETROS7("CAB_moneda","",6)	
	CAB_IdDivision = GF_PARAMETROS7("idDivision",0,6)
	CAB_estado = GF_PARAMETROS7("estado","",6)
	CAB_idContrato = GF_PARAMETROS7("CAB_idContrato",0,6)
	'Se lleen los valores de la planilla comparativa para controlar y determinar la moneda del pedido.
	'(Si no hay planilla, la moneda es PESOS).
	if (CAB_idPedido > 0) then Call obtenerGanadorPlanilla(CAB_Moneda, CAB_ImportePlanilla, CAB_IdProveedorPCP)
	RedimVarialbles(nroLinea) 	
	'Leer detalles
	'Response.Write nroLinea
	for index = 0 to nroLinea - 1
		IT_artID(index) = GF_PARAMETROS7(ITEM_ID & index,3,6)
		'Response.Write IT_artID(index) & "<br>"		
		if (IT_artID(index) <> 0)  then			
			IT_artDS(index) = GF_PARAMETROS7(ITEM_DESC & index,"",6)			
			IT_unidadDS(index) = GF_PARAMETROS7(ITEM_UN_DESC & index,"",6)
			IT_cantidad(index) = GF_PARAMETROS7(ITEM_CANT & index,3,6)									
			if (CAB_idObra = OBRA_GEID) then
				IT_artBA(index) = 0
				IT_artBD(index) = 0
			else				
				IT_artBA(index) = GF_PARAMETROS7("msBudgetArea" & index,0,6)
				IT_artBD(index) = GF_PARAMETROS7("msBudgetDetalle" & index, "",6)	
			end if				
			if (IT_artBD(index) = "") then IT_artBD(index) = 0

			myImporte = cdbl(GF_PARAMETROS7(ITEM_IMPP & index, 2,6))
			myImporte = cdbl(myImporte*100)
			
			if (CAB_Moneda = MONEDA_PESO) then			    			    
			    IT_importePesos(index) = myImporte
			    IT_importeDolares(index) = myImporte / tipoCambio
			    
			else
			    IT_importeDolares(index) = myImporte
			    IT_importePesos(index) = myImporte * tipoCambio
			end if
									            
			CAB_importeDolares = CAB_importeDolares + IT_importeDolares(index) - round((IT_importeDolares(index) * bonificacion)/100,0)
			CAB_importePesos = CAB_importePesos + IT_importePesos(index) - round((IT_importePesos(index) * bonificacion)/100,0)
			
			call getArticuloFull(IT_artID(index), IT_artDS(index), IT_unidadDS(index))										
			Call getUnidadArticulo(IT_artID(index), IT_unidadID(index), IT_unidadCD(index), dsUnidad)
		end if
	next		
	member1Cd = GF_PARAMETROS7("member1Cd","",6)	
	member2Cd = GF_PARAMETROS7("member2Cd","",6)	
	if ((accion = ACCION_CONTROLAR) or (accion = ACCION_GRABAR)) then
		'Se pidio controlar o grabar		
		flagGrabar = false		
		' SI EL PIC TIENE UN CONTRATO ASOCIADO SIGNIFICA QUE ES UN PAGO DE CONTRATO
		flagCTC = false		
		if(CAB_idContrato > 0)then flagCTC = true
		Call cargarDiccionario(CAB_idObra, CAB_idCotizacion)
		if (controlarCotizacion()) then
			if (accion = ACCION_GRABAR) then
				if ((bonificacion > 0) and (InStr(1, CAB_observaciones, MSG_ALERT_BONIF) = 0)) then CAB_observaciones = "Se ha aplicado una bonificación del " & bonificacion & " %.<br>" & CAB_observaciones				
				if (CAB_idCotizacion <> 0) then Call delCTZItems(CAB_idCotizacion)				
				Call addCTZCabecera(CAB_idCotizacion, CAB_idObra, CAB_idPedido, CAB_idProveedor, CAB_fecEntrega, editText4DB(CAB_observaciones), CAB_importePesos, CAB_importeDolares, tipoCambio, CAB_IdDivision, CAB_Moneda, CAB_idContrato)
				
				for index = 0 to UBound(IT_artID)
					if (bonificacion > 0) then 
						IT_importePesos(index) = cdbl(IT_importePesos(index) - round((IT_importePesos(index) * bonificacion)/100,0))
						IT_importeDolares(index) = cdbl(IT_importeDolares(index) - round((IT_importeDolares(index) * bonificacion)/100,0))
					else
						IT_importePesos(index) = cdbl(IT_importePesos(index) - (IT_importePesos(index) * bonificacion)/100)
						IT_importeDolares(index) = cdbl(IT_importeDolares(index) - (IT_importeDolares(index) * bonificacion)/100)
					end if							
					if (IT_artID(index) > 0) then call addCTZItems(CAB_idCotizacion, IT_artID(index), IT_cantidad(index), IT_unidadID(index), IT_artBA(index), IT_artBD(index), IT_importePesos(index), IT_importeDolares(index), tipoCambio)
				next
				
                Call confirmarBudget(CAB_idObra)
				'Se graban las firmas del PIC.
				Call addCTZFirmas(CAB_idCotizacion, member1Cd,  member2Cd)
				'subo los archivos a la base de datos
				auxFiles = split(uploadFilesName,",")
				for j = 0 to ubound(auxFiles)-1
					Call picFile2Binary(CAB_idCotizacion,PATH_COMPRAS_TEMP & "/"& auxFiles(j))
				next
				call AlrtaCDEmailControl()
				flagGrabar = true					
			end if
		end if
	end if	
	if (accion = ACCION_PROCESAR) then
	    'Se graban solo las firmas del PIC.
	    if (controlarFirmas()) then	        
	        Call addCTZFirmas(CAB_idCotizacion, member1Cd,  member2Cd)
	        strSQL="Update TBLCTZCABECERA set ESTADO='0' where IDCOTIZACION=" & CAB_idCotizacion
	        Call executeQueryDb(DBSITE_SQL_INTRA, rsx, "UPDATE", strSQL)    
        end if	        
	end if
end if

ctz_docCode = "PIC"
if (flagCTC) then ctz_docCode = "CEC"

member1Ds = getUserDescription(member1Cd)
						
if (verRemitos) then
	'obtengo remitos asociados	
	esModificableImportes = false
	esModificableCantidades = false
	esModificable = false
else
    esModificableImportes = puedeModificarImportes()
    esModificableCantidades = puedeModificarCantidades()
    esModificable = (esModificableImportes or esModificableCantidades)	
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Cotizacion") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link type="text/css" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<link rel="stylesheet" href="CSS/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.msgOK {
	font-weight: bold;
	font-size: 14px;
	color: #44CC66;
}
.trvisible { display:block; }
.troculto { display:none; }
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.ui-autocomplete-loading { background: white url('images/loading_small_green.gif') right center no-repeat; }

.ui-autocomplete-category {
	font-weight: bold;
	padding: .2em .4em;
	margin: auto;
	text-align:center;
	line-height: 1.5;
}
</style>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/uploadManager.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script language="javascript" src="scripts/magicSearchObj.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript">
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
	
	var ITEM_ID = "ARTID_";
	var ITEM_ID_DIV = "ARTIDDIV_";
	var ITEM_DESC = "ARTDS_";
	var ITEM_UN_DESC = "ARTUNDS_";
	var ITEM_UN_DESC_DIV = "ARTUNDIV_";
	var ITEM_CANT = "CAN_";
	var ITEM_IMPP = "IMPP_";
	
	var lastCategory = "";
	var uploadFilesName="";
	var importesItem = new Array();
	var importesScr = new Array();		//Backup de los valores enviados a pantalla (Sirve para detectar cambio)	
	var ch = new channel();
	var totalImporte = 0;
	var myAutoCompletesIndexs = new Array();
	
	function deleteCredito(img, idPic, idArticulo) {
		if (confirm("Esta seguro que desea eliminar este crédito pendiente?")){
			img.src='images/loading_small_green.gif';
			ch.bind('comprasAnularCreditoAjax.asp?idCotizacion=' + idPic + '&idArticulo=' + idArticulo,'deleteCredito_callback()');
			ch.send();
		}
	}
	
	function changeDivisionEvent() {	    
	    document.getElementById("idDivision").value = document.getElementById("slDivision").options[document.getElementById("slDivision").selectedIndex].value;
	    loadSignatureTable();
	}
	
	function loadSignatureTable() {
	    var idprov = document.getElementById("CAB_idProveedor").value;
	    if (document.getElementById("CAB_Moneda").selectedIndex) {
		    var cdmoneda = document.getElementById("CAB_Moneda").options[document.getElementById("CAB_Moneda").selectedIndex].value;
        } else {
            var cdmoneda = document.getElementById("CAB_Moneda").value;
        }		
		var iddivision = document.getElementById("idDivision").value;
		var cdSol = document.getElementById("member1Cd").value;		
		var cdAut = document.getElementById("member2Cd").value;		
		
	    ch.bind("comprasPICSignAjax.asp?pedido=<% =CAB_idPedido %>&cto=<% =CAB_idContrato %>&prov=" + idprov + "&autorizante=" + cdAut + "&solicitante=" + cdSol + "&importe=" + totalImporte + "&moneda=" + cdmoneda + "&division=" + iddivision, "signatureTableCallback()");
		ch.send();
	}
	
	function signatureTableCallback(){
		document.getElementById("signatureTableDiv").innerHTML = ch.response();		
		
		var msMember1 = new MagicSearch("", "member1", 40, 2, "comprasStreamElementos.asp?tipo=solicitantes");
		msMember1.setToken(";");
		msMember1.onBlur = seleccionarM1;
		msMember1.setValue(document.getElementById("member1Ds").value);					
	}	
	
	function seleccionAutorizante() {
	    if (document.getElementById("cmbUsrAut")) {
		    var e = document.getElementById("cmbUsrAut");
            document.getElementById("member2Cd").value = e.options[e.selectedIndex].value;
        }            
	}
	
	function deleteCredito_callback() {		
	    mostrarFacturas('<% =idCotizacionElegida %>', '<% =CAB_Moneda %>', '<% =CAB_importePesos %>', '<% =CAB_importeDolares%>');
	}
	
	function deleteAjuste(img, idAjuste, idCotizacion){
	if (confirm("Esta seguro que desea eliminar este ajuste?")){
		img.src='images/loading_small_green.gif';
		ch.bind('comprasAnularAJUPICAjax.asp?idAjuste=' + idAjuste + '&idCotizacion=' + idCotizacion,'deleteAjuste_callback()');
		ch.send();
		}
	}
 
	function deleteAjuste_callback() {		
		mostrarAjustes(<% =idCotizacionElegida %>, 0, 0);
	}	
	function volver() {	
		<% if isInPopUp = 0 then %>
			<% if (useOrigen = 0) then %>
			location.href = "comprasAdministrarCotizaciones.asp<% if (CAB_idPedido > 0) then response.write "?fromAP=1"%>";
			<% else %>
			location.href = "<% =session("Origen") %>";
			<% end if %>
		<% else %>
			parent.cerrarPopUpPics();
		<% end if %>
	}
	
	function closePopUp() {		
		location.href = "comprasAdministrarCotizaciones.asp<% if (CAB_idPedido > 0) then response.write "?fromAP=1"%>";				
	}
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	
	function irREMPIC() {
		mostrarRemitos(<% =idCotizacionElegida %>, 0, 0);
		location.href = "#REMPIC";
	}
	function irAJUPIC() {
		mostrarAjustes(<% =idCotizacionElegida %>, 0, 0);
		location.href = "#AJUPIC";
	}
	function irAjustePIC(pCotizacion,pArticulo, pIdArea, pIdDetalle, pCheckRtrn) {
	    if (pCheckRtrn == '') {
		    var myPage, w, h;
		    w=770;
		    h=450;
		    myPage = 'comprasAjustePIC.asp?idCotizacion=' + pCotizacion + '&idArticulo=' + pArticulo + '&idArea=' + pIdArea + '&idDetalle=' + pIdDetalle;
		    var puw = new PopUpWindow('popupAjuPIC',myPage, w, h,'Ajuste de PIC');		
		    puw.onHideEnd = "refreshPage(" + pCotizacion + ")";		
		 } else {
		    alert(pCheckRtrn);
		 }
	}
    function refreshPage(pIdCotizacion) {
		document.location.href = "comprasPIC.asp?verRemitos=true&idCotizacionElegida=" + pIdCotizacion;
	
	}
	function irFACPIC() {
		mostrarFacturas('<% =idCotizacionElegida %>', '<% =CAB_Moneda %>', '<% =CAB_importePesos %>', '<% =CAB_importeDolares%>');
		location.href = "#FACPIC";
	}
	
	function initImportes() {
		var fila = 0;
		var objP = document.getElementById(ITEM_IMPP + fila);
		while (objP) {
			importesItem[fila] = objP.value.replace(/,/,".");
			importesScr[fila] = objP.value;
			totalImporte += Number(importesItem[fila]);
			fila++;
			objP = document.getElementById(ITEM_IMPP + fila);
		}
	}
	
	function sumarTotal(fila) {		
		var aux1 = 0;		
		var objP = document.getElementById(ITEM_IMPP + fila);
		var objPValue = objP.value.replace(/,/,".");
		importesItem[fila] = objP.value;
		objP.value = editarImporte(objP.value);
		if (objP.value == 0) objP.value = "";
		importesScr[fila] = objP.value;
		calcularTotales();				
	}

    function calcularTotales() {
        var bonificacion = 0;
		var bonificacion = document.getElementById("bonificacion").value;
		totalImporte = 0;
		for (i in importesItem) {
			totalImporte += Number(importesItem[i]);
		}	
		document.getElementById("totalVisible").innerHTML = editarImporte(totalImporte.toString(),2);				
	    var totalBonif = totalImporte * bonificacion/100;
		document.getElementById("bonifVisible").innerHTML = editarImporte(totalBonif.toString(),2);				
		totalImporte = totalImporte - totalBonif;
		document.getElementById("totalVisible2").innerHTML = editarImporte(totalImporte.toString(),2);		
        loadSignatureTable();
    }
    
	function imprimirPIC() {
		window.open("comprasPICPrint.asp?idCotizacionElegida=<% =idCotizacionElegida %>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);				
	}
	
	function editarFirmas() {
	    loadSignatureTable();	
	    document.getElementById("imgEditFirma").style.display='none';
	    document.getElementById("imgSaveFirma").style.display='block';
	}
	function saveFirmas() {
	    submitInfo('<% =ACCION_PROCESAR %>')
	}
	
	function bodyOnLoad(){
		<% if (flagGrabar) then %>					
			pp = new winPopUp('pp', 'comprasConfirmacionNumeroPIC.asp?idPedido=<% =CAB_idPedido %>&idCotizacion=<% =CAB_idCotizacion %>&idObra=<% =CAB_idObra %>', '420', '140', 'Cotizacion Registrada', 'volver()');
		<% end if %>
		var	tb = new Toolbar('toolbar', 6, "images/compras/");
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% if (esModificable) then %>
			idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
			idBtnControl = tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
			startMagicSearch();
			//Se dibujan los items del detalle
		<%		if (not flagCTC) then	%>
			initArticulos();
		<%		end if	
			 end if %>
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		<% if (verRemitos) then %>
			tb.addButton("printer-16x16.png", "Imprimir <% =ctz_docCode %>", "imprimirPIC()");
			tb.addButton("see_all-16x16.png", "Remitos Asociados", "irREMPIC()");
			tb.addButton("see_all-16x16.png", "Facturas Asociadas", "irFACPIC()");
			tb.addButton("see_all-16x16.png", "Ajustes Asociados", "irAJUPIC()");

			var	tb2 = new Toolbar('toolbar2', 6, "images/compras/");
			tb2.addButton("Home-16x16.png", "Home", "irHome()");
			tb2.addButton("Previous-16x16.png", "Volver", "volver()");
			tb2.addButton("printer-16x16.png", "Imprimir <% =ctz_docCode %>", "imprimirPIC()");
			tb2.addButton("see_all-16x16.png", "Remitos Asociados", "irREMPIC()");
			tb2.addButton("see_all-16x16.png", "Facturas Asociadas", "irFACPIC()");
			tb2.addButton("see_all-16x16.png", "Ajustes Asociados", "irAJUPIC()");
			tb2.draw();
		<% else %>		
			nroLineaCarga = parseInt(document.getElementById("nroLinea").value);
		<% end if %>
		tb.draw();		
		initImportes();		
		<%  if (verRemitos) then	%>
			mostrarRemitos(<% =idCotizacionElegida %>, 0, 0);
			mostrarFacturas('<% =idCotizacionElegida %>', '<% =CAB_Moneda %>', '<% =CAB_importePesos %>', '<% =CAB_importeDolares%>');
			mostrarAjustes(<% =idCotizacionElegida %>, 0, 0);
	    <%  else %>
	        loadSignatureTable();
		<%  end if %>		
	}

	function mostrarRemitos_callback() {
		var resp = ch.response();
		document.getElementById("remitos").innerHTML=resp;
	}
	
	function mostrarRemitos(id, art, ty) {
		document.getElementById("remitos").innerHTML="<table align='center'><tr><td><img src='images/compras/loading_big.gif'></td></tr></table>";
		ch.bind('comprasPICRemitosAjax.asp?id=' + id + '&art=' + art + '&ty=' + ty,'mostrarRemitos_callback()');
		ch.send();
	}
	function mostrarAjustes(id, art, ty) {
		document.getElementById("ajustes").innerHTML="<table align='center'><tr><td><img src='images/compras/loading_big.gif'></td></tr></table>";
		ch.bind('comprasPICAjustesAjax.asp?id=' + id + '&art=' + art + '&ty=' + ty,'mostrarAjustes_callback()');
		ch.send();
	}	

	function mostrarAjustes_callback() {
		var resp = ch.response();
		document.getElementById("ajustes").innerHTML = resp;
	}

	function mostrarFacturas_callback() {
		var resp = ch.response();
		document.getElementById("facturas").innerHTML=resp;
	}
	
	function mostrarFacturas(id, mone, impPesos, impDolares) {
		document.getElementById("facturas").innerHTML="<table align='center'><tr><td><img src='images/compras/loading_big.gif'></td></tr></table>";
		ch.bind('comprasPICFacturasAjax.asp?id=' + id + '&cm=' + mone + '&pesos=' + impPesos + '&dolares=' + impDolares,'mostrarFacturas_callback()');
		ch.send();
	}

	function seleccionarM1(ms) {	    
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("member1Cd").value = arr[0];
			document.getElementById("member1Ds").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == "") {
			    document.getElementById("member1Cd").value = "";	
			    document.getElementById("member1Ds").value = "";						
			}
		}					
		loadSignatureTable();	
	}			
			
	function submitInfo(acc){	
		document.getElementById("uploadFilesName").value += uploadFilesName;
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}		
	var nroLineaCarga = 0
	
	function AddLineArticulo(){		
		var myTable = document.getElementById("tblDET");				
		var myNroLinea = nroLineaCarga;
		//myNroLinea = myNroLinea + 1;
		var rArticulo = myTable.insertRow(myNroLinea + 1);
        rArticulo.id = "TR_"+ parseInt(myNroLinea)
		var cCodigo = rArticulo.insertCell(0);
		var cDescripcion = rArticulo.insertCell(1);
		var cCantidad = rArticulo.insertCell(2);		
		var cBgtArea = rArticulo.insertCell(3);
		var cBgtDetalle = rArticulo.insertCell(4);
		var cImporte = rArticulo.insertCell(5);
		
		
		//Col 1
		cCodigo.align = 'center';
		var iArtDiv = document.createElement('div');
		iArtDiv.id = ITEM_ID_DIV + myNroLinea;
		cCodigo.appendChild(iArtDiv);

		//Col 2		
		var hArtID = document.createElement('input');
		hArtID.id = ITEM_ID + myNroLinea;
		hArtID.name = ITEM_ID + myNroLinea;
		hArtID.style.visibility = 'hidden';
		hArtID.style.position = 'absolute';
		cDescripcion.appendChild(hArtID);
		var hArtDS = document.createElement('input');
		hArtDS.id = ITEM_DESC + myNroLinea;
		hArtDS.name = ITEM_DESC + myNroLinea;
		hArtDS.style.visibility = 'visible';		
		hArtDS.style.position = 'relative';
		hArtDS.size = 60;
		cDescripcion.appendChild(hArtDS);
		var hArtUNDS = document.createElement('input');
		hArtUNDS.id = ITEM_UN_DESC + myNroLinea;
		hArtUNDS.name = ITEM_UN_DESC + myNroLinea;
		hArtUNDS.style.visibility = 'hidden';
		hArtUNDS.style.position = 'absolute';		
		cDescripcion.appendChild(hArtUNDS);				
		
		myAutoCompletesIndexs[ITEM_DESC + myNroLinea] = myNroLinea			
		createAutocompleteArticulo(ITEM_DESC + myNroLinea);
		
		//Col 3
		cCantidad.align = 'center';
		var hArtCAN = document.createElement('input');
		hArtCAN.id = ITEM_CANT + myNroLinea;
		hArtCAN.name = ITEM_CANT + myNroLinea;
		hArtCAN.size = '4';
		hArtCAN.style.textAlign = 'right';		
		if (isFirefox) {
			hArtCAN.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");		
		} else {
			hArtCAN.onkeypress = function() { return controlIngreso(this, event, 'N'); };		
		}		
		cCantidad.appendChild(hArtCAN);
		
		var hArtUNDDIV = document.createElement('span');
		hArtUNDDIV.id = ITEM_UN_DESC_DIV + myNroLinea;
		cCantidad.appendChild(hArtUNDDIV);
		
		//Col 4
		var idObra = document.getElementById("CAB_idObra").value;
		
		var iBgtArea = document.createElement('input');
		iBgtArea.id = "msBudgetArea" + myNroLinea;
		iBgtArea.name = "msBudgetArea" + myNroLinea;
		iBgtArea.size=5;
		cBgtArea.align="right";
		cBgtArea.appendChild(iBgtArea);						
		
		//Col 5
		var iBgtDetalle = document.createElement('input');
		iBgtDetalle.id = "msBudgetDetalle" + myNroLinea;
		iBgtDetalle.name = "msBudgetDetalle" + myNroLinea;
		iBgtDetalle.size=5;		
		cBgtDetalle.appendChild(iBgtDetalle);
		//Col 6
		cImporte.align = 'right';
		var hArtIMP = document.createElement('input');
		hArtIMP.id = ITEM_IMPP + myNroLinea;
		hArtIMP.name = ITEM_IMPP + myNroLinea;
		hArtIMP.size = '10';
		hArtIMP.style.textAlign = 'right';
		if (isFirefox) {
			hArtIMP.setAttribute('onkeypress', "return controlIngreso(this, event, 'I')");
			hArtIMP.setAttribute('onblur', "sumarTotal(" + myNroLinea + ")");
		} else {
			hArtIMP.onkeypress = function() { return controlIngreso(this, event, 'I'); };		
			hArtIMP.onblur = function() { sumarTotal(myNroLinea,'P'); };		
		}		
		cImporte.appendChild(hArtIMP);					
		
		document.getElementById("nroLinea").value = myNroLinea + 1; 
		nroLineaCarga++; 

	}
	
	function startMagicSearch(){
		$( "#CAB_dsProveedor").autocomplete({
			minLength: 3,				
			source: "comprasStreamElementos.asp?tipo=JQEmpresas",
			focus: function( event, ui ) {
				$( "#CAB_dsProveedor").val(ui.item.dsempresa);
				return false;
			},
			select: function( event, ui ) {
				var myIndex = myAutoCompletesIndexs[this.id];
				$( "#CAB_idProveedor").val (ui.item.idempresa);
				$( "#CAB_dsProveedor").val (ui.item.dsempresa);
				loadSignatureTable();
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item)
				{
					$( "#CAB_idProveedor").val ("");
					$( "#CAB_dsProveedor").val ("");
					loadSignatureTable();
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			li_Item = $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a><font style='font-size:10;'>" + item.idempresa + " - " + item.dsempresa + " - "+ item.cuit +"</font></a>" )
						.appendTo( ul );						
			return li_Item;
		};		
	}
	
	function initArticulos() {
		//ARTICULOS
		<%
			nroLinea = UBound(IT_artDS)
			if (nroLinea = 0) then nroLinea = MIN_LINEAS
			k = 0
			while (k < nroLinea)
		%>		
			myAutoCompletesIndexs[ITEM_DESC + "<% =k %>"] = <% =k %>;
			createAutocompleteArticulo(ITEM_DESC + "<% =k %>");
		<%		k=k+1	
			wend  %>		
	}
	function createAutocompleteArticulo(pID) {		
		$( "#"+ pID).autocomplete({
			minLength: 2,				
			source: "comprasStreamElementos.asp?tipo=JQArticulos",
			focus: function( event, ui ) {
				$( "#"+ pID).val(ui.item.dsarticulo);
				return false;
			},
			select: function( event, ui ) {
				var myIndex = myAutoCompletesIndexs[this.id];
				$( "#"+ITEM_ID + myIndex).val (ui.item.idarticulo);
				$( "#"+ITEM_ID_DIV + myIndex).html (ui.item.idarticulo);
				$( "#"+ITEM_DESC + myIndex).val (ui.item.dsarticulo);
				$( "#"+ITEM_UN_DESC + myIndex).val (ui.item.abreviatura);
				$( "#"+ITEM_UN_DESC_DIV + myIndex).html (ui.item.abreviatura);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item)
				{
					lastCategory = "";
					var myIndex = myAutoCompletesIndexs[this.id];
					$( "#"+ITEM_ID + myIndex).val ("");
					$( "#"+ITEM_ID_DIV + myIndex).html ("");
					$( "#"+ITEM_DESC + myIndex).val ("");
					$( "#"+ITEM_UN_DESC + myIndex).val ("");
					$( "#"+ITEM_UN_DESC_DIV + myIndex).html ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			li_Item = $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a><font style='font-size:10;'>" + item.idarticulo + " - " + item.dsarticulo + " ["+item.abreviatura+"]</font></a>" )
						.appendTo( ul );
						
			if (lastCategory != item.idcategoria) {
				lastCategory = item.idcategoria;
				return $(ul)
					.append( "<li class='ui-autocomplete-category'>" + item.dscategoria + "</li>" ).append(
						li_Item
					);
			} else {
				return li_Item;
			}
		};
	}
	
	function keyPressEvent(obj, evt) {		
		return controlIngreso(obj, evt, 'I');
	}
	function controlPercent(pObj){	
		if (pObj.value > 100){
			alert("El porcentaje no puede ser mayor a 100!");
			pObj.value = 0; 
		} else {
		    //Aplico el descuento en pantalla.
		    calcularTotales();
		}	
	}

	function REMprint(id) {
		location.href="almacenREM.asp?idRemito=" + id;
	}

	function verDetalle(img, pId) {
		if (document.getElementById(pId).style.display == 'none'){
			img.src = "images/compras/Menos.gif";
			document.getElementById(pId).style.display = '';
		}
		else{
			img.src = "images/compras/Mas.gif";
			document.getElementById(pId).style.display = 'none';
		}
	}
	function abrirFAC() {
	    alert('<% =GF_TRADUCIR("Función no Implementada") %>');		
	}
	function lightOn(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdosHL reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdosHL";
		}
	}
	function lightOff(tr, estado) {
		if (estado == <%=ESTADO_BAJA%>) {
			tr.className = "reg_Header_navdos reg_header_rejected";
		}
		else{
			tr.className = "reg_Header_navdos";
		}
	}
	function abrirCTC(id){
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,resizable=yes",false);		
	}
	
	function abrirPCT(id){
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + id + "&tab=1", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}

</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel" action="comprasPIC.asp?verRemitos=<% =verRemitos %>&idPedido=<%=pct_idPedido%>">
<div id="toolbar"></div><br /><br />
<table class="reg_header" align="center" width="1024px" border="0" >				
	<tr>
		<td colspan="4"><% call showErrors() %></td>
	</tr>
	<tr>
		<td colspan="4">
			<div id='PartidaExcedida' style='font-weight:  bold;color:#FF0000;'>
				<%
				 	Dim j					 		
					for each j in oDiccPartidaExcedida.Keys
						Response.Write oDiccPartidaExcedida(j) & "<br>"
					Next
				%>
			</div>		
			<div id='PartidaNoExcedida' style='font-weight:  bold;'>			
				<%	
					for each j in oDiccPartidaNoExcedida.Keys
						Response.Write oDiccPartidaNoExcedida(j) & "<br>"
					Next						
				
				%>
			<div>
		</td>
	</tr>
	<%	if (idCotizacionElegida <> 0) then %>	        
	<tr>								
		<td align="right" class="numberStyle" colspan="4"><% =GF_TRADUCIR("Id " & ctz_docCode & ":") %>&nbsp;<% =idCotizacionElegida %></td>
	</tr>
	<%	end if	%>
	<tr>
		<td class="reg_header_nav" colspan="4"><% =GF_TRADUCIR("Datos del Pedido") %></td>				
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Ptda. Presup.") %></td>	
		<td colspan="3">
		<% if (not verRemitos) then %>
			<%	if ((pct_idPedido = 0 or CAB_idObra = 0) and (not flagCTC)) then 
					Set rsObras = obtenerListaObras("", "", "","", OBRA_ACTIVA)
					%>						
					<select id="CAB_idObra" name="CAB_idObra">
						<option value="0" >- <% =GF_TRADUCIR("Sin Partida") %> -
						<%	
						while (not rsObras.eof)	%>
							<option value="<% =rsObras("IDOBRA") %>"  <% if (rsObras("IDOBRA") = CAB_idObra) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsObras("CDOBRA")) %> - <% =GF_TRADUCIR(rsObras("DSOBRA")) %>
						<%	rsObras.MoveNext()
						wend 	%>
                        <option value="<% =OBRA_GEID %>"  <% if (OBRA_GEID = Cdbl(CAB_idObra)) then response.write "selected='true'" %>><% =OBRA_GECD %> - <% =GF_TRADUCIR(OBRA_GEDS) %>
					</select>
                    <input type="hidden" name="obraOld" id="obraOld" value="<%=CAB_idObra %>">
					<% 
				else
					%>
					<input type="hidden" name="CAB_idObra" id="CAB_idObra" value="<%=CAB_idObra%>">
					<input type="hidden" name="CAB_obraCD" id="CAB_obraCD" value="<%=CAB_obraCD%>">
					<input type="hidden" name="CAB_obraDS" id="CAB_obraDS" value="<%=CAB_obraDS%>">
					<%
					Response.Write GF_TRADUCIR(CAB_obraCD) & " - " & GF_TRADUCIR(CAB_obraDS)
				end if %>
		<% else
			Response.Write GF_TRADUCIR(CAB_obraCD) & " - " & GF_TRADUCIR(CAB_obraDS)
		   end if %>
		</td>
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Pedido") %></td>	
		<td>
			<input type="hidden" name="CAB_idCotizacion" value="<%=CAB_idCotizacion%>">
			<input type="hidden" name="CAB_idPedido" id="CAB_idPedido" value="<%=CAB_idPedido%>">
			<input type="hidden" name="CAB_cdPedido" id="CAB_cdPedido" value="<%=CAB_cdPedido%>">
			<% if(CAB_idPedido >  0)then %>
				<a><img id="imgPCT" src="images/compras/PCT-16X16.png" style="cursor:pointer" onclick="abrirPCT(<%=CAB_idPedido %>)" title="Abrir Pedido" ></a>&nbsp&nbsp;
			<% end if %>
			<%=CAB_cdPedido%>
		</td>
		
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Division") %></td>	
		<td>					
		<% if (not verRemitos) then %>
			<% 
			if (CAB_idPedido = 0 or pct_idObra = 0) then
				strSQL="Select * from TBLDIVISIONES"
				call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
			%>
				<select id="slDivision" name="slDivision" onchange="changeDivisionEvent()"> 
					<option value="<% =SIN_DIVISION %>" selected="true">- <% =GF_TRADUCIR("Seleccione") %> -
					<%	
					while (not rsDivision.eof) 	
						if (checkPointAcceso(rsDivision("IDDIVISION"))) then 
							if not isAuditor(rsDivision("IDDIVISION"))	then
							%>
								<option value="<% =rsDivision("IDDIVISION") %>" <% if (CAB_IdDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION")	%>
							<%			
							end if
						end if
						rsDivision.MoveNext()
						wend	
					%>								
				</select>						
			<%		
			else
				Response.Write getDivisionDS(CAB_IdDivision) %>				
			<%
			end if
		else
			Response.Write getDivisionDS(CAB_IdDivision) %>			
<%		end if%>
            <input type="hidden" name="idDivision" id="idDivision" value="<%=CAB_idDivision%>">
		</td>
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Proveedor") %></td>
		<td>
			<% if (not verRemitos) then %>
				<input type="text" id="CAB_dsProveedor" name="CAB_dsProveedor" size="80" value="<%=CAB_dsProveedor%>">				
			<% else %>
				<%=CAB_dsProveedor%>
				<input type="hidden" id="CAB_dsProveedor" name="CAB_dsProveedor" value="<%=CAB_dsProveedor%>">				
			<% end if %>
			<input type="hidden" id="CAB_idProveedor" name="CAB_idProveedor" value="<%=CAB_idProveedor%>">            
		</td>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Moneda") %></td>
		<td>
		<% if (not verRemitos) then %>
			<select id="CAB_Moneda" name="CAB_Moneda" onchange="loadSignatureTable()">
			    <option value="<% =MONEDA_PESO %>" <% if (CAB_Moneda = MONEDA_PESO) then response.write "selected" end if %>><% =getNombreMoneda(MONEDA_PESO) %></option>
			    <option value="<% =MONEDA_DOLAR %>" <% if (CAB_Moneda = MONEDA_DOLAR) then response.write "selected" end if %>><% =getNombreMoneda(MONEDA_DOLAR) %></option>
			</select>
		<% else %>
			<% =getNombreMoneda(CAB_MONEDA) %>
			<input type="hidden" name="CAB_Moneda" id="CAB_Moneda" value="<%=CAB_Moneda%>">
		<% end if %>
		</td>
	</tr>
	<% if (CAB_idContrato > 0) then %>
		<tr>
			<td class="reg_header_navdos"><% =GF_TRADUCIR("Contrato") %></td>
			<td>
				<a><img id="imgCTC" src="images/compras/ctc-16x16.png" style="cursor:pointer" onclick="abrirCTC(<%=CAB_idContrato%>)" title="Abrir CTC" ></a>&nbsp&nbsp;
				<%=getCodigoCTC(CAB_idContrato)%>
			</td>
		</tr>
	<% end if %>	
	<tr>
		<td class="reg_header_nav" colspan="4"><% =GF_TRADUCIR("Detalle") %></td>
	</tr>	
	<tr>
		<td colspan="4" align="center">
			<table width="100%" id="tblDET">
				<tr>
					<td class="reg_header_nav" align="center">	<% =GF_TRADUCIR("Cód") %></td>	
					<td class="reg_header_nav" width="50%">	<% =GF_TRADUCIR("Descripción") %></td>	
					<td class="reg_header_nav" align="center">	<% =GF_TRADUCIR("Cant") %></td>	
					<td class="reg_header_nav" colspan="2" align="center">	<% =GF_TRADUCIR("Ptda. Presup.")	%></td>	
					<td class="reg_header_nav" align="center" width="20%">	<% =GF_TRADUCIR("Importe Total (S/IVA)") %></td>						
				</tr>
				<%
				nroLinea = 0
				diffCambioPesos = 0				
				for index = 0 to ubound(IT_artID) - 1		
				    'Si el artículo es la diferencia de cambios, el mismo no se muestra, sino que se acumula su importe para mostrar al final.
				    if (IT_artID(index) = CTZ_ITEM_DIFF_CAMBIO) then
				        diffCambioPesos = IT_importePesos(index)
				    else				    
					    lineClass = ""
					    if (dicError.Exists(IT_artID(index))) then
						    lineClass = "reg_header_error" 
					    elseif ((dicArtUCC.Exists(IT_artID(index))) or (itemsNecesitanAFE.exists(IT_artBA(indes)&"-"&IT_artBD(index)))) then 
						    lineClass = "reg_header_warning" 
					    end if
					    subTotalImportePesos = subTotalImportePesos + IT_importePesos(index)
					    subTotalImporteDolares = subTotalImporteDolares + IT_importeDolares(index)
				    %>
					    <tr id="TR_<%=index%>" class='<% =lineClass %>'  >
						    <td align="center">							
							    <div id="ARTIDDIV_<%=index%>"><%=IT_artID(index)%></div>		
						    </td>						
						    <td>
							    <% if ((not verRemitos) and (not flagCTC)) then %>
								    <input type="input" size="90" id="ARTDS_<%=index%>" name="ARTDS_<%=index%>" value="<%=IT_artDS(index)%>">
							    <% else %>
								    <%=IT_artDS(index)%>
									<input type="hidden" id="ARTDS_<%=index%>" name="ARTDS_<%=index%>" value="<%=IT_artDS(index)%>">
							    <% end if %>
							    <input type="hidden" id="ARTID_<%=index%>" name="ARTID_<%=index%>" value="<%=IT_artID(index)%>">							    
							    <input type="hidden" id="ARTUNDS_<%=index%>" name="ARTUNDS_<%=index%>" value="<%=IT_unidadDS(index)%>">							
						    </td>	

						    <td align="center">
						    <%	if ((esModificableCantidades) and (not flagCTC)) then	%>
								    <input style="text-align:right;" type="text" size="4" name="CAN_<%=index%>" id="CAN_<%=index%>" value="<%=IT_cantidad(index) %>" onKeyPress="return controlIngreso(this, event, 'N')"><!--&nbsp;&nbsp;-->
						    <%	else
								    Response.Write IT_cantidad(index)	%>
								    <input type="hidden" name="CAN_<%=index%>" id="CAN_<%=index%>" value="<%=IT_cantidad(index) %>">
						    <%	end if								%>
							    <span id="ARTUNDIV_<%=index%>"><%=IT_unidadDS(index)%></span>
						    </td>	
					    <%  'Se determina que importe mostrar según la moneda en la que está nominado el PIC.
					        if (CAB_Moneda = MONEDA_PESO) then			    			    
			                    myImporteScr = IT_importePesos(index)
			                else
			                    myImporteScr = IT_importeDolares(index)
			                end if
					        if ((not verRemitos) and (not flagCTC)) then %>						
						    <td align="right">
								    <input type="text" id="msBudgetArea<% =index %>" name="msBudgetArea<% =index %>" value="<%=IT_artBA(index)%>" size="5">
						    </td>
						    <td align="left">
								    <input type="text" id="msBudgetDetalle<% =index %>" name="msBudgetDetalle<% =index %>" value="<%=IT_artBD(index)%>" size="5" >
						    </td>
						    <td align="right">
								    <input style="text-align:right;" type="text" size="20" name="IMPP_<%=index%>" id="IMPP_<%=index%>" value="<% if (myImporteScr <> "") then Response.write myImporteScr/100 %>" onBlur="sumarTotal(<% =index %>)" onKeyPress="return keyPressEvent(this, event)"> 
						    </td>
					    <% else %>
						    <td align="right"><%=IT_artBA(index)%>-</td>						
						    <td align="left"><%=IT_artBD(index)%></td>
						    <td align="right"><% if (myImporteScr <> "") then Response.write GF_EDIT_DECIMALS(myImporteScr,2) %></td>
						    <% checkRtrn = checkAjustar(idCotizacionElegida, IT_artID(index),IT_artBA(index),IT_artBD(index)) %>
							<td> <img src="images/compras/ajustes.gif" style="cursor:pointer;" title="Ajustar Producto" onclick="irAjustePIC('<%=idCotizacionElegida%>','<%=IT_artID(index)%>','<%=IT_artBA(index)%>','<%=IT_artBD(index)%>', '<% =checkRtrn %>')"> </td>						    
						    <td> </td>
						    <input type="hidden" id="msBudgetArea<% =index %>" name="msBudgetArea<% =index %>" value="<%=IT_artBA(index)%>">
						    <input type="hidden" id="msBudgetDetalle<% =index %>" name="msBudgetDetalle<% =index %>" value="<%=IT_artBD(index)%>">
						    <input type="hidden" name="IMPP_<%=index%>" id="IMPP_<%=index%>" value="<% if (myImporteScr <> "") then Response.write myImporteScr/100 %>"> 						    
					    <% end if %>
					    </tr>	
					    <% if(dicArtUCC.Exists(IT_artID(index)))then%>
					    <tr>
						    <td colspan="5" class="reg_Header_Warning" style='font-weight:  bold;color:#FF0000;'>
							    &nbsp;&nbsp;<% =dicArtUCC.Item(IT_artID(index))%>
						    </td>					
					    </tr>
					    <%	end if
						    nroLinea = nroLinea + 1
				    end if
				next 
			        if (CAB_Moneda = MONEDA_PESO) then			    			    
	                    subTotalImporte = subTotalImportePesos
	                    auxTotal = CAB_importePesos
	                    myImporteBonifScr = round(subTotalImportePesos * bonificacion/100, 2)
	                else
	                    subTotalImporte = subTotalImporteDolares
	                    auxTotal = CAB_importeDolares
	                    myImporteBonifScr = round(subTotalImporteDolares * bonificacion/100, 2)
	                end if
					%>
					<tr>
						<td class="reg_header_navdos" colspan="5" align="right"><font size="+1"><b><% =GF_TRADUCIR("Sub-Total") %>&nbsp;&nbsp;</b></font></td>	
						<td align="right"><font size="+1"><b><div id="totalVisible"><% =GF_EDIT_DECIMALS(subTotalImporte,2)%></div></b></font></td>
					</tr>
					<tr>
						<td colspan="4" align="right"><font size="+1"><b><% =GF_TRADUCIR("Bonificación") %> (%)&nbsp;&nbsp;</b></font></td>	
						<td align="right" width="30px">
						<% if (not verRemitos) then %>
							<input style="text-align:right;" onBlur="controlPercent(this)" type="text" id="bonificacion" name="bonificacion" value="<%=bonificacion%>" size="5" onKeyPress="return keyPressEvent(this, event)">
						<% else %>
							<b><% response.write bonificacion %></b>
						<% end if %>
						
						</td>
						<td align="right"><font size="+1">
							<div id="bonifVisible"><%=GF_EDIT_DECIMALS(myImporteBonifScr,2)%></div></font>
						</td>						
					</tr>
					<%  'La diferencia de cambios solo tiene sentido para los PICs en dolares. Y solo puede agregarse si el PIC ya está autorizado.
					    if ((CAB_MONEDA = MONEDA_DOLAR) and (verRemitos))then %>
					<tr>
						<td colspan="4" align="right"><font size="+1"><b><% =GF_TRADUCIR("Difrencia de Cambio (AR$)") %>&nbsp;&nbsp;</b></font></td>							
						<td align="right"></td>
						<td align="right">
						    <font size="+1">
							    <div id="Div1"><%=GF_EDIT_DECIMALS(diffCambioPesos,2)%></div>
							</font>
					    </td>						
						<% checkRtrn = checkAjustar(idCotizacionElegida, CTZ_ITEM_DIFF_CAMBIO, 0, 0) %>
						<td> <img src="images/compras/ajustes.gif" style="cursor:pointer;" title="Agregar Diferencia" onclick="irAjustePIC('<%=idCotizacionElegida%>','<%=CTZ_ITEM_DIFF_CAMBIO%>','0','0', '<% =checkRtrn %>')"> </td>
					</tr>
					<% end if %>
					<tr>
						<td class="reg_header_navdos22" colspan="5">&nbsp;</td>	
						<td colspan="2"><hr></td>	
						<td width="2%" align="center">
						<% if (not verRemitos) then %>
							<img src="images/add.gif" onClick="AddLineArticulo()" style="cursor:pointer" id="addItem">
						<% end if %>
						</td>
					</tr>
					<tr>
						<td class="reg_header_navdos" colspan="5" align="right"><font size="+1"><b><% =GF_TRADUCIR("Total") %>&nbsp;&nbsp;</b></font></td>	
						<td align="right"><font size="+1"><b><div id="totalVisible2"><%=GF_EDIT_DECIMALS(auxTotal ,2)%></div></b></font></td>	
					</tr>
			</table>	
		</td>
	</tr>	
	
<%	    if (verRemitos) then        %>	
    <tr>
	    <td class="reg_header_nav" colspan="4"><% =GF_TRADUCIR("Observaciones") %></td>		
	</tr>
	<tr>
			<td colspan="2" align="left"><%=CAB_observaciones%></td>
    </tr>
    <tr>
        <td class="reg_header_nav" colspan="4"><% =GF_TRADUCIR("Firmas") %></td>
        <td  style="background: #FFFFFF; cursor: pointer;" align="center">
<%          if ((CAB_estado =  CTZ_PENDIENTE) or (CAB_estado =  CTZ_EN_FIRMA)) then %> 
            <img title="<%=GF_TRADUCIR("Editar Firmas")%>" id="imgEditFirma" src="images\compras\edit-16x16.png" onclick='editarFirmas()'>
            <img title="<%=GF_TRADUCIR("Guardar Firmas")%>" style="display: none;" id="imgSaveFirma" src="images\save_b-16.png" onclick='saveFirmas()'>
<%          end if   %>                        
        </td>
    </tr>
    <tr>			
            <td colspan="4">		                    
                <div id="signatureTableDiv">
                <table align="center" width="80%" border="1" cellspacing=0 cellpadding=0>
<%             Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLCTZFIRMAS_GET_BY_IDCOTIZACION", CAB_IdCotizacion)		             %>
		            <tr>
<%                  while (not rsFirmas.eof) %>		            
			            <td align="center" width="33%">
				            <%	if (rsFirmas("HKEY")  <> "") then %>
					            <img src="images/firmas/<% =obtenerFirma(rsFirmas("CDUSUARIO")) %>"><br>					            
                            <%  else %>					            
                                <br /><br /><br /><br /><br />
				            <%	end if	%>
					        <% =getUserDescription(rsFirmas("CDUSRROL")) %>
			            </td>			    			            
<%                      if (CInt(rsFirmas("SECUENCIA")) = 3) then   %>
                            </tr>
                            <tr>
<%                      end if
                        rsFirmas.MoveNext()  
                    wend			             %>
                    </tr>			                    
	            </table>   		
	            </div>
            </td>
	</tr>
<%      else     %>		
    <tr>
		<td class="reg_header_nav" colspan="3"><% =GF_TRADUCIR("Observaciones") %></td>
		<td class="reg_header_nav" ><% =GF_TRADUCIR("Firmas") %></td>
    </tr>				    
    <tr>
        <td  align="center" colspan="3">
			<textarea type="text" cols="100" rows="7" name="CAB_observaciones"><% =editText4Input(CAB_observaciones) %></textarea>
		</td>
		<td >		    
	        <div id="signatureTableDiv"></div>
	    </td>
    </tr>	    
<%      end if   %>		
	
	
</table>

<br>
<% if (idCotizacionElegida > 0) then %>
	<iframe src="compraspicfiles.asp?idcotizacion=<%=idCotizacionElegida%>&showuploader=true" style="border: 0 none;width: 100%;"></iframe>
<% else %>
	<iframe src="compraspicfiles.asp?idcotizacion=<%=idCotizacionElegida%>&showuploader=true&origen=nuevo&uploadFilesName=<%=uploadFilesName%>" style="border: 0 none;width: 100%;"></iframe>
<% 	
end if %>
<br />
<a name="REMPIC">	
<div id="toolbar2"></div><br>
<span id="remitos"></span>
<br>
<a name="FACPIC">
<span id="facturas"></span>
<br>
<a name="AJUPIC">	
<span id="ajustes"></span>
<br>
<input type="hidden" id="nroLinea" name="nroLinea" value="<%=nroLinea%>">
<input type="hidden" name="accion" id="accion">
<input type="hidden" name="estado" id="estado" value="<% =CAB_estado %>">
<input type="hidden" name="tipoCambio" id="tipoCambio" value="<% =tipoCambio %>">
<input type="hidden" name="idCotizacionElegida" id="idCotizacionElegida" value="<% =idCotizacionElegida %>">
<input type="hidden" name="uploadFilesName" id="uploadFilesName" value="<%=uploadFilesName%>">
<input type="hidden" name="isInPopUp" id="isInPopUp" value="<%=isInPopUp%>">
<input type="hidden" name="CAB_idContrato" id="CAB_idContrato" value="<% =CAB_idContrato %>">
<input type="hidden" id="member1Cd" name="member1Cd" value="<%=member1Cd%>">
<input type="hidden" id="member1Ds" name="member1Ds" value="<% =member1Ds  %>">
<input type="hidden" id="member2Cd" name="member2Cd" value="<%=member2Cd%>">
</form>
</body>
</html>