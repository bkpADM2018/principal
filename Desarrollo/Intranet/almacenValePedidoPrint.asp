<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->   
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Const SEPARATION = 10
Const PAGE_HEIGHT_SIZE = 570
Const MAXREGISTROS = 40
Const TEXTOFIRMAS = 0
Const FIRMAS = 1
Const SEPARATION_IMAGE_SIGN = 190
Const SEPARATION_IMAGE_SIGN_AJS = 145
'------------------------------------------------------------------------------------------------------
'Function checkValuacionValeAJS : verifica si el vale de Ajuste de Stock tiene asignado la firma para el Director
Function checkValuacionValeAJS()
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA , rs, "TBLVALESFIRMAS_GET_BY_FILTERS", idVale & "||"&VS_FIRMA_DIRECTOR&"||||||")	
	If not rs.Eof then checkValuacionValeAJS = true
End function
'------------------------------------------------------------------------------------------------------
Function isAjuste(cdVale)
	isAjuste = false
	if ((cdVale = CODIGO_VS_AJUSTE_VALE)	or _
		(cdVale = CODIGO_VS_AJUSTE_STOCK)	or _
		(cdVale = CODIGO_VS_AJUSTE_PEDIDO)	or _
		(cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA)) then
		isAjuste = true
	end if
End Function
'------------------------------------------------------------------------------------------------------
Function isAnulacion(cdVale)
	isAnulacion = false
	if (cdVale = CODIGO_VS_AJUSTE_VALE_X	or _
		cdVale = CODIGO_VS_AJUSTE_STOCK_X	or _
		cdVale = CODIGO_VS_AJUSTE_PEDIDO_X	or _
		cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA_X) then
		isAnulacion = true
	end if
End Function
'------------------------------------------------------------------------------------------------------
Function isTransferencia(cdVale)	
	if (cdVale = CODIGO_VS_TRANSFERENCIA	or _
		cdVale = CODIGO_VS_RECEPCION	or _
		cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA	or _
		cdVale = CODIGO_VS_TRANSFERENCIA_X	or _
		cdVale = CODIGO_VS_AJUSTE_TRANSFERENCIA_X	or _
		cdVale = CODIGO_VS_RECEPCION_X) then
		isTransferencia = true
	end if	
End Function
'------------------------------------------------------------------------------------------------------
sub PM2VS()
	'VS = PM
	VS_FechaSolicitud = PM_FechaSolicitud
	VS_FechaRequerido = PM_FechaRequerido
	VS_cdSolicitante = PM_cdSolicitante
	VS_dsSolicitante = PM_dsSolicitante
	VS_idPedido = PM_idPedido
	VS_idAlmacen = PM_idAlmacen
	VS_idObra = PM_idObra	
	VS_usuario = PM_usuario
	VS_momento = PM_momento
	VS_hayCabecera = PM_hayCabecera
end sub
'------------------------------------------------------------------------------------------------------
Function armadoCabeceraBox()
	'dibuja celdas y lineas
	Call GF_squareBox(oPDF, 2, 2, 590, 848, 0, "", "#000000", 2, PDF_SQUARE_ROUND)
    Call GF_horizontalLine(oPDF,10,80,570)
	Call GF_squareBox(oPDF, 10, 90, 570, 15, 0, "#80A2B7", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 105, 570, 10, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 115, 570, 15, 0, "#80A2B7", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 130, 570, 10, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 140, 570, 15, 0, "#80A2B7", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, 100, 115, 25)
	Call GF_verticalLine(oPDF, 190, 90, 50)
	Call GF_verticalLine(oPDF, 390, 115, 25)
	Call GF_verticalLine(oPDF, 510, 90, 25)
end Function
'------------------------------------------------------------------------------------------------------
Function armadoCabeceraInfo(p_Titulo, p_NroOrden, p_esPedido)
	'escribe datos	
	Call GF_writeImage(oPDF, Server.MapPath("images\ADMlogo2.jpg"), 10, 10, 60, 55, 0)
	call GF_setFont(oPDF,"ARIAL", 24,0)
	Call GF_writeTextAlign(oPDF,10, 20, GF_TRADUCIR(p_Titulo), 570, PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"ARIAL",16,0)
	if (p_esPedido) then
		Call GF_writeTextAlign(oPDF,380, 60, GF_TRADUCIR("Pedido Nro: ") & p_NroOrden, 200, PDF_ALIGN_RIGHT)
	else
		Call GF_writeTextAlign(oPDF,380, 60, GF_TRADUCIR("Vale Nro: ") & p_NroOrden, 200, PDF_ALIGN_RIGHT)
	end if
	Call GF_setFontColor("#FFFFFF")
	call GF_setFont(oPDF,"ARIAL",10,0)
	Call GF_writeTextAlign(oPDF,190, 92, GF_TRADUCIR("Part. Presup./Budget | Sector")	, 320, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 10, 92, GF_TRADUCIR("Almacén")				, 180, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,510, 92, GF_TRADUCIR("Referencia")				,  70, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 10, 117, GF_TRADUCIR("Solicitado el")			,  90, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,100, 117, GF_TRADUCIR("Requerido para el")		,  90, PDF_ALIGN_CENTER)
	if ((VS_cdVale <> CODIGO_VS_RECEPCION) and (VS_cdVale <> CODIGO_VS_RECEPCION_X)) then
		Call GF_writeTextAlign(oPDF,190, 117, GF_TRADUCIR("Solicitante"), 200, PDF_ALIGN_CENTER)
	else
		Call GF_writeTextAlign(oPDF,190, 117, GF_TRADUCIR("Origen"), 200, PDF_ALIGN_CENTER)
	end if
	if (isAjuste(VS_cdVale)) then
		auxEntrego = "Ajusto"
	elseif (isAnulacion(VS_cdVale)) then
		auxEntrego = "Anulo"
	elseif (VS_cdVale = CODIGO_VS_RECEPCION) then
		auxEntrego = "Recibio"
	else
		auxEntrego = "Entrego"
	end if	
	Call GF_writeTextAlign(oPDF,390, 117, GF_TRADUCIR(auxEntrego)  , 190, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 15, 142, GF_TRADUCIR("Articulos"), 500, PDF_ALIGN_LEFT)
	Call GF_setFontColor("#000000")
end Function
'------------------------------------------------------------------------------------------------------
Function armadoCabecera(p_Titulo, p_NroOrden, p_partPresupuestaria, p_sector, p_almacen, p_almacenDestino, p_referencia, p_fechaSolicitado, p_fechaRequerido, p_solicitante, p_entrego, p_esPedido)
	Dim partida	
	Call armadoCabeceraBox()
	Call armadoCabeceraInfo(p_Titulo, p_NroOrden, p_esPedido)
	call GF_setFont(oPDF,"ARIAL",8,0)
	if (p_partPresupuestaria <> "") then	
		partida = p_partPresupuestaria	
	else	
		partida = p_sector
	end if
	call GF_writeTextAlign(oPDF,190, 106, partida	, 320, PDF_ALIGN_CENTER)	
	if (p_almacen <> "")			then	Call GF_writeTextAlign(oPDF, 10, 106, p_almacen				, 180, PDF_ALIGN_CENTER)
	if (p_referencia <> "")			then	Call GF_writeTextAlign(oPDF,510, 106, "PM: " & p_referencia	,  70, PDF_ALIGN_CENTER)
	if (p_fechaSolicitado <> "")	then	Call GF_writeTextAlign(oPDF, 10, 131, p_fechaSolicitado		,  90, PDF_ALIGN_CENTER)
	if (p_fechaRequerido <> "")		then	Call GF_writeTextAlign(oPDF,100, 131, p_fechaRequerido		,  90, PDF_ALIGN_CENTER)	
	if (p_almacenDestino <> "")	then
		Call GF_writeTextAlign(oPDF,190, 131, p_almacenDestino, 200, PDF_ALIGN_CENTER)
	else
		Call GF_writeTextAlign(oPDF,190, 131, p_solicitante, 200, PDF_ALIGN_CENTER)
	end if
	if (p_entrego <> "")	then	Call GF_writeTextAlign(oPDF,390, 131, p_entrego, 190, PDF_ALIGN_CENTER)
end Function
'------------------------------------------------------------------------------------------------------
Function armadofindepaginaBox()
	Call GF_squareBox(oPDF, 10, 600, 570, 15, 0, "#80A2B7", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 615, 570, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 715, 570,  15, 0, "#80A2B7", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, 730, 570, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",10,0)
	Call GF_setFontColor("#FFFFFF")
	Call GF_writeTextAlign(oPDF,15, 602, GF_TRADUCIR("Observaciones :"), 190, PDF_ALIGN_LEFT)
	if (VS_cdVale <> CODIGO_VS_AJUSTE_STOCK and VS_cdVale <> CODIGO_VS_AJUSTE_STOCK_X) and (VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK and VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK_X) then
		Call GF_writeTextAlign(oPDF, 10, 717, GF_TRADUCIR("Aprobación Resp. Almacén"),   	   190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,200, 717, GF_TRADUCIR("Aprobación Solicitante"),		   190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,390, 717, GF_TRADUCIR("Aprobación Supervisor (Opcional)"), 190, PDF_ALIGN_CENTER)
	else
		auxScope = 190
		auxSeparation = 190
		if (isAjsMayor) then
			auxScope = 130
			auxSeparation = 145
		end if
		px = 10
		Call GF_writeTextAlign(oPDF,px, 717,  GF_TRADUCIR("Visado Responsable"),   	      auxScope, PDF_ALIGN_CENTER)
		px = px + auxSeparation
		Call GF_writeTextAlign(oPDF,px, 717, GF_TRADUCIR("Visado Gerente de Planta"),    auxScope, PDF_ALIGN_CENTER)
		if (VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK and VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK_X) then
			px = px + auxSeparation	
			Call GF_writeTextAlign(oPDF,px, 717, GF_TRADUCIR("Visado"),auxScope, PDF_ALIGN_CENTER)
		end if
		px = px + auxSeparation
		if (isAjsMayor) then Call GF_writeTextAlign(oPDF,px, 717, GF_TRADUCIR("Aprobación Dirección"),auxScope, PDF_ALIGN_CENTER)		
	end if	
	Call GF_setFontColor("#000000")
end Function
'------------------------------------------------------------------------------------------------------
Function armadofindepagina(p_dsEntrego, p_dsSolicitante, p_observaciones, p_modo)
	dim y_aux, nrValeRelacionado
	Dim strSQL, conn, rsFirmas,cdGerente,dsGerente,cdCoordAudit,dsCoordAudit,firmaGerente,firmaCoordAudit
	Dim cdResponsable,dsResponsable, firmaResponsable,cdDirector,firmaDirector,dsDirector
	
	y_aux = 617
	if (not esPedido) then nrValeRelacionado = getValeRelacionado(idVale)
	Call armadofindepaginaBox()
	Call GF_setFontColor("#FF0000")
	if ((VS_estado = ESTADO_BAJA) and (nrValeRelacionado <> ""))then
		call GF_setFont(oPDF,"ARIAL",60,8)
		Call GF_writeTextAlign(oPDF, 20, 280, GF_TRADUCIR("ANULADO  POR"), 570	, PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(oPDF, 20, 380, GF_TRADUCIR("VALE  Nº:")	, 570	, PDF_ALIGN_LEFT)
		Call GF_writeTextAlign(oPDF, 20, 480, nrValeRelacionado			, 570	, PDF_ALIGN_CENTER)
	elseif (VS_estado = ESTADO_ANULACION) then
		call GF_setFont(oPDF,"ARIAL",10,0)
		Call GF_writeTextAlign(oPDF, 15, 617,GF_TRADUCIR("Anulacion del Vale Nº: ") & nrValeRelacionado,560, PDF_ALIGN_LEFT)
		y_aux = 635
	end if
	Call GF_setFontColor("#000000")
	if (p_modo = TEXTOFIRMAS) then
		call GF_setFont(oPDF,"ARIAL",14,0)	
		Call GF_writeTextAlign(oPDF,200, 740, GF_TRADUCIR("La firma del pedido"), 190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,200, 760, GF_TRADUCIR("se realiza unicamente"), 190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,200, 780, GF_TRADUCIR("en la ultima pagina"), 190, PDF_ALIGN_CENTER)
	else
		Call GF_setFont(oPDF,"ARIAL",10,0)
		if (not isAjsMayor) then
			Call GF_verticalLine(oPDF, 200, 715, 115)
			Call GF_verticalLine(oPDF, 390, 715, 115)
			Call GF_horizontalLine(oPDF, 30,810,150)
			Call GF_horizontalLine(oPDF,220,810,150)
			Call GF_horizontalLine(oPDF,410,810,150)
		else
			Call GF_verticalLine(oPDF, 145, 715, 115)
			Call GF_verticalLine(oPDF, 290, 715, 115)
			Call GF_verticalLine(oPDF, 435, 715, 115)
			Call GF_horizontalLine(oPDF, 15,810,125)
			Call GF_horizontalLine(oPDF,155,810,125)
			Call GF_horizontalLine(oPDF,300,810,125)
			Call GF_horizontalLine(oPDF,445,810,125)
		end if
		if ((VS_cdVale <> CODIGO_VS_AJUSTE_STOCK and VS_cdVale <> CODIGO_VS_AJUSTE_STOCK_X) and (VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK and VS_cdVale <> CODIGO_VS_RECLASIFICACION_STOCK_X)) then
			if (p_dsEntrego <> "")			then	Call GF_writeTextAlign(oPDF, 10, 815, p_dsEntrego,     190, PDF_ALIGN_CENTER)
			if ((p_dsSolicitante <> "") and _
				((VS_cdVale <> CODIGO_VS_RECEPCION) and (VS_cdVale <> CODIGO_VS_RECEPCION_X)) and _
				((VS_cdVale <> CODIGO_VS_AJUSTE_TRANSFERENCIA) and (VS_cdVale <> CODIGO_VS_AJUSTE_TRANSFERENCIA_X)) and _
				((VS_cdVale <> CODIGO_VS_TRANSFERENCIA) and (VS_cdVale <> CODIGO_VS_TRANSFERENCIA_X))) then	
				Call GF_writeTextAlign(oPDF, 200, 815, p_dsSolicitante, 190, PDF_ALIGN_CENTER)			
			end if
		else
			strSQL = "Select * from TBLVALESFIRMAS where IDVALE=" & idVale & " order by SECUENCIA"
			call executeQueryDb(DBSITE_SQL_INTRA, rsFirmas, "OPEN", strSQL)
			if (not rsFirmas.eof) then
				while not rsFirmas.eof
					select case cint(rsFirmas("SECUENCIA"))
					case VS_FIRMA_RESPONSABLE
						cdResponsable = rsFirmas("CDUSUARIO")
						if (rsFirmas("HKEY") <> "") then firmaResponsable = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
					case VS_FIRMA_GERENTE
						cdGerente = rsFirmas("CDUSUARIO")
						if (rsFirmas("HKEY") <> "") then firmaGerente     = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
					case VS_FIRMA_COORD_AUDIT
						cdCoordAudit = rsFirmas("CDUSUARIO")
						if (rsFirmas("HKEY") <> "") then firmaCoordAudit  = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
					case VS_FIRMA_DIRECTOR
						cdDirector = rsFirmas("CDUSUARIO")
						if (rsFirmas("HKEY") <> "") then firmaDirector  = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))	
					end select
					rsFirmas.movenext
				wend
				'------------
				pX_img = 55
				pX_sign = 30
				fontSize = 6
				auxRenglon = SEPARATION_IMAGE_SIGN
				if (isAjsMayor) then 
					auxRenglon = SEPARATION_IMAGE_SIGN_AJS					
					pX_img = 20
					pX_sign = 12
					fontSize = 5
				end if	
				if (firmaResponsable <> "") then
				'if (cdResponsable <> VS_NO_USER) then
					Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(cdResponsable)), pX_img, 745, 96, 48, 0)
					Call GF_setFont(oPDF,"ARIAL",fontSize,0)
					Call GF_writeTextAlign(oPDF, pX_sign, 800, firmaResponsable, auxRenglon, PDF_ALIGN_RIGHT)												
				end if
				dsResponsable = getUserDescription(cdResponsable)
				if (dsResponsable <> "") then
				    Call GF_setFont(oPDF,"ARIAL",10,0)
					Call GF_writeTextAlign(oPDF, pX_sign - 12, 815, dsResponsable,    auxRenglon, PDF_ALIGN_CENTER)
				end if
				'------------
				auxRenglon = SEPARATION_IMAGE_SIGN
				if (isAjsMayor) then auxRenglon = SEPARATION_IMAGE_SIGN_AJS
				pX_img = pX_img + auxRenglon
				pX_sign = pX_sign + auxRenglon				
				
				if (firmaGerente <> "") then
				'if (cdGerente <> VS_NO_USER) then					
					Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(cdGerente)), pX_img, 745, 96, 48, 0)
					Call GF_setFont(oPDF,"ARIAL",fontSize,0)
					Call GF_writeTextAlign(oPDF, pX_sign, 800, firmaGerente,     auxRenglon, PDF_ALIGN_RIGHT)												
				end if
			    dsGerente = getUserDescription(cdGerente)					
				if (dsGerente <> "") then
				    Call GF_setFont(oPDF,"ARIAL",10,0)
					Call GF_writeTextAlign(oPDF, pX_sign - 12, 815, dsGerente,     auxRenglon, PDF_ALIGN_CENTER)
				end if
				'------------
				auxRenglon = SEPARATION_IMAGE_SIGN
				if (isAjsMayor) then auxRenglon = SEPARATION_IMAGE_SIGN_AJS
				pX_img = pX_img + auxRenglon
				pX_sign = pX_sign + auxRenglon				
				
				if (firmaCoordAudit <> "") then
				'if ((cdCoordAudit <> VS_NO_USER) and (cdCoordAudit <> VS_PORT_SUPERVISOR_USER) and (cdCoordAudit <> VS_AUDIT_USER)) then				    				
					Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(cdCoordAudit)), pX_img, 745, 96, 48, 0)
					Call GF_setFont(oPDF,"ARIAL",fontSize,0)
					Call GF_writeTextAlign(oPDF, pX_sign, 800, firmaCoordAudit,     auxRenglon, PDF_ALIGN_RIGHT)																
				end if
				dsCoordAudit = getUserDescription(cdCoordAudit)				
				if (dsCoordAudit <> "") then
				    Call GF_setFont(oPDF,"ARIAL",10,0)					
					Call GF_writeTextAlign(oPDF, pX_sign - 12, 815,dsCoordAudit,     auxRenglon, PDF_ALIGN_CENTER)				
				end if
				if (isAjsMayor) then					
					pX_img = pX_img + SEPARATION_IMAGE_SIGN_AJS
					pX_sign = pX_sign + SEPARATION_IMAGE_SIGN_AJS
					if (firmaDirector <> "") then
					'if ((cdCoordAudit <> VS_NO_USER) and (cdCoordAudit <> VS_PORT_SUPERVISOR_USER) and (cdCoordAudit <> VS_AUDIT_USER)) then				    				
						Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(cdDirector)), 435, 745, 96, 48, 0)
						Call GF_setFont(oPDF,"ARIAL",fontSize,0)
						Call GF_writeTextAlign(oPDF, pX_sign, 800, firmaDirector,     auxRenglon, PDF_ALIGN_RIGHT)																
					end if
					dsDirector = getUserDescription(cdDirector)				
					if (dsDirector <> "") then
					    Call GF_setFont(oPDF,"ARIAL",10,0)					
						Call GF_writeTextAlign(oPDF, pX_sign - 12, 815, dsDirector,     auxRenglon, PDF_ALIGN_CENTER)				
					end if
				end if	
			end if
		end if
		if (p_observaciones <> "")		then	Call GF_writeTextPlus(oPDF, 15, y_aux,p_observaciones,560, 9,PDF_ALIGN_LEFT)
	end if
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,10,835, GF_TRADUCIR("Cargó: " & VS_usuario & " - " & GF_FN2DTE(VS_momento)), 50, PDF_ALIGN_LEFT)

end Function
'------------------------------------------------------------------------------------------------------
Function armadoBoxArticulos()
	'dibuja celda
	Call GF_squareBox(oPDF, 10, 155, 570, PAGE_HEIGHT_SIZE, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'escribe datos
	Call GF_setFontColor("#F4F4F4")
	call GF_setFont(oPDF,"ARIAL",10,0)
	Call GF_writeTextAlign(oPDF,10, 157, GF_TRADUCIR("Código"), 50, PDF_ALIGN_CENTER)
	Call GF_writeText(oPDF,75, 157, GF_TRADUCIR("Descripción"), 0)
	Call GF_writeText(oPDF,430, 157, GF_TRADUCIR("C. Interno"), 0)
	Call GF_writeTextAlign(oPDF,510, 157, GF_TRADUCIR("Cantidad"), 70, PDF_ALIGN_CENTER)
	Call GF_setFontColor("#000000")
end Function
'------------------------------------------------------------------------------------------------------
Function getArticuloDatosArticulo (idAlmacen, idArticulo, ByRef dsArticulo, ByRef abrrArticulo, ByRef cdInterno)
	Dim strSQL, rs, conn
	
	call getArticuloFull (idArticulo, dsArticulo, abrrArticulo)
	'Se trae el codigo interno
	strSQL = "Select * from TBLARTICULOSDATOS where IDALMACEN=" & idAlmacen & " and  idArticulo=" & idArticulo
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then cdInterno = rs("CDINTERNO")
End Function
'------------------------------------------------------------------------------------------------------
Function armadoArticulos()
	dim posiciony, i, totalreg, nPagina, cdInterno, idAlmacen
	nPagina = 1
	posiciony = 175
	totalreg = rsArticulos.RecordCount
	while (not rsArticulos.eof)
		i = i + 1
		idArticulo = rsArticulos("IDARTICULO")
		idAlmacen = VS_idAlmacen
		call getArticuloDatosArticulo(idAlmacen, idArticulo, dsArticulo, abrrArticulo, cdInterno)		
		'escribe los articulos
		Call GF_setFont(oPDF,"COURIER",8,0)
		Call GF_writeTextAlign(oPDF,10, posiciony, idArticulo, 50, PDF_ALIGN_CENTER)
		Call GF_writeText(oPDF,75, posiciony, dsArticulo, 0)
		if (cdInterno<>"") then Call GF_writeText(oPDF,430, posiciony, Left(cdInterno, 15), 0)				
		Call GF_writeTextAlign(oPDF,510, posiciony, rsArticulos("cantidad") & " " & abrrArticulo, 60, PDF_ALIGN_RIGHT)
		posiciony = posiciony + 10
		if ((posiciony >= PAGE_HEIGHT_SIZE) and (i < totalreg)) then
			Call GF_writeText(oPDF,80, posiciony + 5, GF_TRADUCIR("-----  Continua en la siguiente pagina  -----"), 0)
			Call GF_writeTextAlign(oPDF,530,835, GF_TRADUCIR("Pagina Nº " & nPagina), 50, PDF_ALIGN_RIGHT)
			Call nuevaPagina ()
			nPagina = nPagina + 1
			posiciony = 175
		end if
		rsArticulos.movenext
	wend
	Call GF_writeTextAlign(oPDF,530,835, GF_TRADUCIR("Pagina Nº " & nPagina), 50, PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,75, posiciony + 2, GF_TRADUCIR("-----  Fin del " & mainTitle & "  -----"), 300, PDF_ALIGN_LEFT)
end Function
'------------------------------------------------------------------------------------------------------
Function nuevaPagina ()
	Call armadofindepagina(dsEntrego, vs_dsSolicitante, observaciones, TEXTOFIRMAS)
	Call GF_newPage(oPDF)
	Call armadoCabecera(mainTitle, nroOrden, partPresupuestaria, sector, almacen, almacenDestino, VS_PartidaPendiente, vs_FechaSolicitud, vs_FechaRequerido, vs_dsSolicitante, dsEntrego, esPedido)
	Call armadoBoxArticulos()
end Function
'------------------------------------------------------------------------------------------------------
Function armadoPDF()
	Call armadoCabecera(mainTitle, nroOrden, partPresupuestaria, sector, almacen, almacenDestino, VS_PartidaPendiente, vs_FechaSolicitud, vs_FechaRequerido, vs_dsSolicitante, dsEntrego, esPedido)
	Call armadoBoxArticulos()
	Call armadoArticulos()
	Call armadofindepagina(dsEntrego, vs_dsSolicitante, observaciones, FIRMAS)
end Function
'------------------------------------------------------------------------------------------------------
'*********************************************************************************************************
'*****************************************	COMIENZA LA PAGINA	*****************************************
'*********************************************************************************************************
dim idPedido, idVale, mainTitle, oPDF, nroOrden, partPresupuestaria, sector, almacen, almacenDestino
dim idEntrego, dsEntrego, dsArticulo, abrrArticulo, idArticulo, rsBudget, esPedido, observaciones, rsSector, isAjsMayor

idPedido = GF_PARAMETROS7("idPedido",0,6)
idVale = GF_PARAMETROS7("idVale",0,6)
'defino titulo y tipo de orden
esPedido = false

if idPedido <> 0 then
	Call initHeaderPMDB(idPedido)
	mainTitle = "Pedido de Material"
	Call PM2VS
	nroOrden = idPedido
	esPedido = true
else
	if idVale <> 0 then
		Call initHeaderValeDB(idVale)
		mainTitle = getLeyendaCdVale(VS_cdVale)
		nroOrden = VS_nrVale
	else
		response.redirect "comprasaccesodenegado.asp"
	end if
end if
if (VS_hayCabecera or PM_hayCabecera) then

	partPresupuestaria = ""
	sector=""
	if esPedido then
		'partPresupuestaria
		call loadDatosObra(PM_idObra, PM_cdObra, PM_dsObra, 0, "", 0, "", 0, "", "", "", "", "")
		if (PM_idObra > 0) then
			partPresupuestaria = PM_cdObra & " - " & left(PM_dsObra,40)
			partPresupuestaria = partPresupuestaria & " / " & PM_idBudgetArea & " - " & PM_idBudgetDetalle
		else
			'Sector
			Set rsSector = obtenerSectores(PM_idSector)
			if (not rsSector.eof) then sector = rsSector("IDSECTOR") & "-" & rsSector("DSSECTOR")			
		end if
		'almacen
		Set rsAlmacenes = obtenerListaAlmacenes(PM_idAlmacen)
		if (not rsAlmacenes.eof) then 
			almacen = rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
		else
			almacen = GF_TRADUCIR("ERROR")
		end if
		if PM_idAlmacenDest > 0 then
			Set rsAlmacenes = obtenerListaAlmacenes(PM_idAlmacenDest)
			if (not rsAlmacenes.eof) then 
				almacenDestino = rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
			else
				almacenDestino = GF_TRADUCIR("ERROR")
			end if
		end if
		observaciones = PM_comentario
		'registros de articulos
		strSql = "select * from tblpmdetalle where idpedido = " & idPedido
	else
		'partPresupuestaria
		call loadDatosObra(vs_idObra, vs_cdObra, vs_dsObra, 0, "", 0, "", 0, "", "", "", "", "")
		if (vs_idObra > 0) then
			partPresupuestaria = vs_cdObra & " - " & left(vs_dsObra,40)
			partPresupuestaria = partPresupuestaria & " / " & VS_idBudgetArea & " - " & VS_idBudgetDetalle
		else
			'Sector
			Set rsSector = obtenerSectores(VS_idSector)
			if (not rsSector.eof) then sector = rsSector("IDSECTOR") & "-" & rsSector("DSSECTOR")			
		end if
		'almacen
		Set rsAlmacenes = obtenerListaAlmacenes(vs_idAlmacen)
		if (not rsAlmacenes.eof) then 
			almacen = rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
		else
			almacen = GF_TRADUCIR("ERROR")
		end if
		rsAlmacenes.close
		set rsAlmacenes = nothing
		
		if (isTransferencia(VS_cdVale)) then
			call initHeaderPMDB(VS_PartidaPendiente)
			if ((VS_cdVale <> CODIGO_VS_RECEPCION) and (VS_cdVale <> CODIGO_VS_RECEPCION_X))  then
				Set rsAlmacenes = obtenerListaAlmacenes(PM_idAlmacenDest)
			else
				Set rsAlmacenes = obtenerListaAlmacenes(PM_idAlmacen)
			end if	
			if (not rsAlmacenes.eof) then
				almacenDestino = rsAlmacenes("CDALMACEN") & " - " & rsAlmacenes("DSALMACEN")
			else
				almacenDestino = GF_TRADUCIR("ERROR")
			end if
		end if
		dsEntrego = getUserDescription(VS_usuario)		
		observaciones = getComentarioVale(idVale)
		'registros de articulos
		strSql = "select * from tblvalesdetalle where idvale = " & idVale
	end if
	isAjsMayor = false
	if VS_cdVale = CODIGO_VS_AJUSTE_STOCK then isAjsMayor = checkValuacionValeAJS()
	call executeQueryDb(DBSITE_SQL_INTRA, rsArticulos, "OPEN", strSQL)
	Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)	
	call armadoPDF()
	Call GF_closePDF(oPDF)
else
	response.redirect "comprasaccesodenegado.asp"
end if
%>