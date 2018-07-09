<%
dim afe_idAFE, afe_CdAFE, afe_IdObra, afe_IdPedido, afe_NroAFEComplID, afe_IdDivision, afe_Categoria, afe_CatOtros, afe_Tipo, afe_TipoOtros, afe_TipoCC, afe_Descripcion,afe_ImportePesos, afe_ImporteDolares, afe_TipoCambio, afe_NPV, afe_IRR, afe_ROIC, afe_PAYBACK, afe_PreparedBy, afe_RequestedBy, afe_EngReview, afe_Officer, afe_VicePresident, afe_President, afe_PreparedByCD, afe_RequestedByCD, afe_EngReviewCD, afe_OfficerCD, afe_VicePresidentCD, afe_PresidentCD
dim afe_Preg1, afe_Preg2, afe_Preg3, afe_Preg4, afe_Preg4_Text, afe_Preg5, afe_Preg6, afe_Preg7, afe_Preg8, afe_Preg9, afe_Preg10, afe_Preg11
dim afe_IdProveedor, afe_Titulo, afe_Departamento, afe_DsProveedor, afe_ObraDS, afe_ObraDivID, afe_ObraDivDS, afe_ObraCuentaDS, afe_ObraImporte, afe_ObraMoneda, afe_ObraFechaInicio, afe_ObraFechaFin, afe_ObraFechaAjustada, afe_ObraRespCD, afe_ObraRespDS, afe_Confirmado
dim afe_ObraCD, afe_RespSectorID, afe_RespSectorDS, afe_cdUsuario, afe_Momento, afe_FechaBudget,afe_IDArea,afe_IDDetalle
dim afe_ControllerCD,afe_Controller,afe_AuditorCD,afe_Auditor, afe_cfo, afe_cfoCD, afe_cfoHkey, afe_cfoHkeyDate, afe_BDT
dim afe_PreparedByHkey,afe_RequestedByHkey, afe_EngReviewHkey,afe_OfficerHkey,afe_VicePresidentHkey,afe_ControllerHkey,afe_AuditorHkey
dim afe_PreparedByHkeyDate,afe_RequestedByHkeyDate, afe_EngReviewHkeyDate,afe_OfficerHkeyDate,afe_VicePresidentHkeyDate,afe_ControllerHkeyDate,afe_AuditorHkeyDate,afe_PresidentHkey,afe_PresidentHkeyDate,afe_isCFO
'---------------------------------------------------------------------------------------------
'**** CONSTANTES
'---------------------------------------------------------------------------------------------
Const PATH_AFE_TEMP =  "Temp"
Const PATH_AFE_FINAL =  "Documentos\ArchivosCompras"

Const AFE_TODOS = 0
Const AFE_RAIZ = 1

Const AFE_AREA_TODAS = 0
Const AFE_DETALLE_TODOS = 0

Const HKEY_RECHAZO = "RECHAZO"

'Codigos de categorias y tipos
Const AFE_CATEGORIA_CAPITAL			= "C"
Const AFE_CATEGORIA_GASTOS			= "G"
Const AFE_CATEGORIA_INVERSIONES		= "I"
Const AFE_CATEGORIA_COMPLEMENTARIO	= "A"
Const AFE_CATEGORIA_ALQUILER		= "Q"
Const AFE_CATEGORIA_SERVICIOS		= "S"
Const AFE_CATEGORIA_OTROS			= "O"

Const AFE_TIPO_MEJORA				= "E"
Const AFE_TIPO_REPUESTOS			= "R"
Const AFE_TIPO_CAPACIDAD			= "I"
Const AFE_TIPO_MANTENIMIENTO		= "M"
Const AFE_TIPO_VEHICULOS			= "V"
Const AFE_TIPO_CUMPIMIENTO			= "C"
Const AFE_TIPO_CUMPLIMIENTO_SEG		= "S"
Const AFE_TIPO_CUMPLIMIENTO_MA		= "A"
Const AFE_TIPO_CUMPLIMIENTO_NC		= "N"
Const AFE_TIPO_CAMBIO_OBJETIVO		= "D"
Const AFE_TIPO_DESVIO				= "Y"
Const AFE_TIPO_COMUNICACIONES		= "T"
Const AFE_TIPO_OTROS				= "O"

Const AFE_NO_CONFIRMADO = "N"
'SI SU ESTADO ES NUMERICO, SIGNIFICA QUE ESTA EN FIRMA
Const AFE_ESPERA_HAMBURGO = "H"
Const AFE_APROBADO = "A"
Const AFE_ANULADO = "R"
Const AFE_ANULACION = "X"

'Margen de desvió autorizado.
Const AFE_MAX_DESVIO_PCN = 0.1		'Maximo desvio porcentual.
Const AFE_MAX_DESVIO_ABS = 5000000	'Maximo desvio absoluto en dolares. (Centavos)
'Monto minimo del Afe para tener autorizacion de Hamburgo 
Const DIVSION_EXPORTACION = 1
'---------------------------------------------------------------------------------------------
function addAFE(byref pIdAFE, byref pCdAFE, pIdObra, pIdPedido, pCuenta, pNroAFECompl, pTitulo, pIdDivision, pDepartamento, pCategoria, pCatOtros, pTipo, pTipoOtros, pTipoCC, pDescripcion, pImportePesos, pImporteDolares, pTipoCambio, pArr, pIrr, pRONA, pPAYBACK, pPreparedBy, pRequestedBy, pEngReview, pChkFinanzas, pIdArea, pIdDetalle, pNroAFEAnula)
dim strSQL, rs, conn, rsIns, connIns, esta

if (CLng(pIdAFE) <> 0) then
	strSQL="Select * from TBLDATOSAFE where IDAFE = " & pIdAFE
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then esta = true	
end if
 
if (not esta) then
	pCdAFE = generateCdAFE(pIdDivision, pNroAFECompl, pNroAFEAnula)
	strSQL = "Insert Into TBLDATOSAFE(CDAFE,IDOBRA,IDPEDIDO,CUENTA,NROAFECOMPL,IDDIVISION,DEPARTAMENTO,TITULO,CATEGORIA,CATOTROS,TIPO,TIPOOTROS,TIPOCC,DESCRIPCION,IMPORTEPESOS,IMPORTEDOLARES,TIPOCAMBIO,ARR,IRR,RONA, PAYBACK, CONFIRMADO, CDUSUARIO, MOMENTO, IDAREA, IDDETALLE,FINANZAS) VALUES('" & pCdAFE & "'," & pIdObra & "," & pIdPedido & ",'" & pCuenta & "'," & pNroAFECompl & "," & pIdDivision & ",'" & pDepartamento & "','" & pTitulo & "','" & pCategoria & "','" & pCatOtros & "', '" & pTipo & "', '" & pTipoOtros & "', '" & pTipoCC & "', '" & pDescripcion & "', " & pImportePesos & ", " & pImporteDolares & ", " & pTipoCambio & ", " & pArr & ", " & pIrr & "," & pRONA & "," & pPAYBACK & ", '" & AFE_NO_CONFIRMADO & "', '" & session("Usuario") & "', " & session("MmtoDato") & "," & pIdArea & "," & pIdDetalle & ",'" & pChkFinanzas & "')"	
	Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
	strSQL = "Select IDAFE from TBLDATOSAFE where CDAFE='" & pCdAFE & "'"		
	Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "OPEN", strSQL)
	pIdAFE = rsIns("IDAFE")
 	Call UpdateAfeSignatories(pIdAFE,pPreparedBy, pRequestedBy, pEngReview )
else
	Call UpdateAfeSignatories(pIdAFE,pPreparedBy, pRequestedBy, pEngReview )
	strSQL = "Update TBLDATOSAFE set CUENTA='" & pCuenta & "', NROAFECOMPL=" & pNroAFECompl & ", IDDIVISION=" & pIdDivision & ", DEPARTAMENTO='" & pDepartamento & "', TITULO='" & pTitulo & "',CATEGORIA='" & pCategoria & "', CATOTROS='" & pCatOtros & "', TIPO='" & pTipo & "', TIPOOTROS='" & pTipoOtros & "', TIPOCC='" & pTipoCC & "', DESCRIPCION='" & pDescripcion & "', IMPORTEPESOS=" & pImportePesos & ", IMPORTEDOLARES=" & pImporteDolares & ", TIPOCAMBIO=" & pTipoCambio & ", ARR=" & pArr & ", IRR=" & pIrr & ", RONA=" & pRONA & ", PAYBACK=" & pPAYBACK & ", CONFIRMADO='" & AFE_NO_CONFIRMADO & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoDato") & ", IDAREA=" & pIdArea & ", IDDETALLE= " & pIdDetalle & ", FINANZAS = '" & pChkFinanzas & "' where IDAFE = " & pIdAFE    
	Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "UPDATE", strSQL)
end if
addAFE = true
end function
'---------------------------------------------------------------------------------------------
Function UpdateAfeSignatories(pIdAFE,pPreparedBy, pRequestedBy, pEngReview )

	Dim strSQL,conn,rs,i,Signature
    redim Signature(2) 	
	Signature(0) = pPreparedBy
    Signature(1) = pRequestedBy
    Signature(2) = pEngReview

	strSQL = "SELECT * FROM TBLAFEFIRMAS WHERE IDAFE = " & pidAfe
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.eof) then
        '1) SI EXISTEN LAS FIRMA DEL AFE LAS BORRO A TODAS        
        strSQL = "DELETE FROM TBLAFEFIRMAS WHERE IDAFE = "& pIdAFE
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if
    '2) AGREGO SOLO LOS FIRMANTES POR USUARIO (EL QUE PREPARA EL AFE, EL QUE REQUIERE EL AFE Y EL DE REVISION TECNICA)
	strSQL = "INSERT INTO TBLAFEFIRMAS (IDAFE, SECUENCIA, CDUSUARIO, FECHAFIRMA , HKEY) VALUES "
    for i = LBound(Signature) to UBound(Signature)
        strSQL = strSQL & " (" & pIdAFE & "," & i + 1 & ",'" & Signature(i) & "', null,''),"
    next
    strSQL = left(strSQL,len(strSQL)-1)
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	
end Function
'---------------------------------------------------------------------------------------------
sub readAFE(pIdAFE, pIdObra, pIdPedido)
'on error resume next
dim strSQL, rs, conn
if (not isFormSubmit()) then
	strSQL="Select IDAFE, CDAFE , IDOBRA , IDPEDIDO, NROAFECOMPL ,IDDIVISION, DEPARTAMENTO, TITULO, CATEGORIA, CATOTROS, CUENTA, TIPO, TIPOOTROS, TIPOCC, DESCRIPCION, IMPORTEPESOS, IMPORTEDOLARES, TIPOCAMBIO, ARR, IRR, RONA, PAYBACK, PREGUNTA1, PREGUNTA2, PREGUNTA3, PREGUNTA4, PREGUNTA4_TEXT, PREGUNTA5, PREGUNTA6, PREGUNTA7, PREGUNTA8, PREGUNTA9, PREGUNTA10, PREGUNTA11, CONFIRMADO, CDUSUARIO, MOMENTO, IDAREA, IDDETALLE, CASE WHEN FINANZAS IS NULL THEN 'N' ELSE FINANZAS END AS FINANZAS from TBLDATOSAFE where IDAFE = " & pIdAFE
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		afe_IdAFE			= CDbl(rs("IDAFE"))
		afe_CdAFE			= rs("CDAFE")
		afe_IdObra			= CDbl(rs("IDOBRA")) 
		afe_IdPedido		= CDbl(rs("IDPEDIDO")) 
		afe_NroAFEComplID	= CDbl(rs("NROAFECOMPL"))
		afe_IdDivision		= CDbl(rs("IDDIVISION"))
		afe_Categoria		= Trim(rs("CATEGORIA"))
		afe_CatOtros		= Trim(rs("CATOTROS"))
		afe_Tipo			= Trim(rs("TIPO"))
		afe_TipoOtros		= Trim(rs("TIPOOTROS"))
		afe_TipoCC			= Trim(rs("TIPOCC"))
		afe_Descripcion		= Trim(rs("DESCRIPCION"))
		afe_ImportePesos	= CDbl(rs("IMPORTEPESOS"))
		afe_ImporteDolares	= CDbl(rs("IMPORTEDOLARES"))
		afe_TipoCambio		= CDbl(rs("TIPOCAMBIO"))
		afe_NPV				= rs("ARR")
		afe_Irr				= rs("IRR")
		afe_Confirmado		= rs("CONFIRMADO")
		afe_cdUsuario		= rs("CDUSUARIO")
		afe_Momento			= rs("MOMENTO")
		afe_IDArea			= rs("IDAREA")
		afe_IDDetalle		= rs("IDDETALLE")
		afe_Titulo			= rs("TITULO")
		if (afe_IDArea <> "") then afe_IDArea = CDbl(afe_IDArea)
		if (afe_IDDetalle <> "") then afe_IDDetalle = CDbl(afe_IDDetalle)
		afe_ROIC			= rs("RONA")
		afe_PAYBACK			= rs("PAYBACK")
		afe_isCFO           = rs("FINANZAS")        
		Call LoadSignatories(pIdAFE)
		
		afe_IdProveedor = 0
		afe_DsProveedor = ""
		if afe_IdPedido > 0 then
			Call initHeader(afe_IdPedido)					
			afe_IdProveedor = pct_idProveedorElegido
			afe_DsProveedor = pct_dsProveedorElegido			
		    if ((pct_idObra > 0) and (afe_IdObra = 0)) then afe_IdObra = CDbl(pct_idObra)
		    if ((pct_idArea > 0) and (afe_IDArea = 0)) then afe_IDArea = CDbl(pct_idArea)
		    if ((pct_idDetalle > 0) and (afe_IDDetalle = 0)) then afe_IDDetalle = CDbl(pct_idDetalle)
		end if
		if (afe_IdObra > 0) then
			Call loadDatosObra(afe_IdObra, afe_ObraCD, afe_ObraDS, afe_ObraDivID, afe_ObraDivDS, afe_ObraImporte, afe_FechaBudget, afe_ObraMoneda, afe_ObraFechaInicio, afe_ObraFechaFin, afe_ObraFechaAjustada, afe_ObraRespCD, afe_ObraRespDS)				
			'JAS --> Valores leidos por respeto de herencia para AFEs viejos.
			Call loadSectorEmpleado(afe_ObraRespCD, afe_RespSectorID, afe_RespSectorDS)			
			'<-- JAS
		else
			'No hay obra
			afe_ObraDivDS = getDivisionDS(afe_IdDivision)
			afe_ObraDivID = afe_IdDivision
			afe_ObraFechaFin = 0
		end if
		if (CLng(afe_ObraFechaFin) = 0) then 			
			if (pct_idProveedorElegido > 0) then
				if (afe_IdProveedor = 0) then afe_IdProveedor = pct_idProveedorElegido
			end if	
		end if							
		
		
	
		'JAS --> Valores leidos por respeto de herencia para AFEs viejos.
		afe_ObraCuentaDS	= rs("CUENTA")	
		afe_Departamento	= rs("DEPARTAMENTO")
		afe_Preg1			= rs("PREGUNTA1")
		afe_Preg2			= rs("PREGUNTA2")
		afe_Preg3			= rs("PREGUNTA3")
		afe_Preg4			= rs("PREGUNTA4")
		afe_Preg4_Text		= rs("PREGUNTA4_TEXT")
		afe_Preg5			= rs("PREGUNTA5")
		afe_Preg6			= rs("PREGUNTA6")
		afe_Preg7			= rs("PREGUNTA7")
		afe_Preg8			= rs("PREGUNTA8")
		afe_Preg9			= rs("PREGUNTA9")
		afe_Preg10			= rs("PREGUNTA10")
		afe_Preg11			= rs("PREGUNTA11")
		'<-- JAS
		
	else	
		call initHeader(pIdPedido)				
		if (pct_idObra > 0) then pIdObra = pct_idObra
		if (pct_idProveedorElegido > 0) then 
			afe_IdProveedor = pct_idProveedorElegido
			afe_DsProveedor = pct_dsProveedorElegido
		end if		
		afe_TipoCambio = cdbl(getTipoCambio(MONEDA_DOLAR, ""))			
		call loadDatosObra(pIdObra, afe_ObraCD, afe_ObraDS, afe_ObraDivID, afe_ObraDivDS, afe_ObraImporte, afe_FechaBudget, afe_ObraMoneda, afe_ObraFechaInicio, afe_ObraFechaFin, afe_ObraFechaAjustada, afe_ObraRespCD, afe_ObraRespDS)
		'if (not isInversion(pIdObra)) then afe_ObraDS = ""
		if (pIdObra = 0) then
			'No hay obra!
			afe_ObraDivDS = pct_dsDivision
			afe_IdDivision = pct_idDivision
			afe_ObraDivID = afe_IdDivision			
		end if	
		
		afe_IdAFE = 0
		afe_CdAFE = "PENDIENTE"
		afe_IdObra = pIdObra
		afe_IdPedido = 0
		afe_IdDivision = afe_ObraDivID
		afe_Categoria = ""		
		afe_CatOtros = ""
		afe_Tipo = ""
		afe_TipoOtros = ""
		afe_TipoCC = ""
		afe_Descripcion = ""
		if (afe_ObraImporte = "") then afe_ObraImporte = 0
				
		'Detemino el importe sugeridom del AFE.
		afe_ImportePesos = 0
		afe_ImporteDolares = 0
		
		if (pIdObra > 0) then
			'Hay obra
			if (afe_ObraMoneda = MONEDA_PESO) then
				afe_ImportePesos = CDbl(afe_ObraImporte)
				afe_ImporteDolares = CDbl(afe_ImportePesos / cdbl(afe_TipoCambio))
			else	
				afe_ImporteDolares = CDbl(afe_ObraImporte)
				afe_ImportePesos = CDbl(afe_ImporteDolares) * cdbl(afe_TipoCambio)
			end if
		else
			'No hay obra
			afe_ImportePesos = obtenerImporteCotizacionElegida(pIdPedido, MONEDA_PESO,afe_TipoCambio)
			afe_ImporteDolares = obtenerImporteCotizacionElegida(pIdPedido, MONEDA_DOLAR,afe_TipoCambio)
		end if
				
		afe_TipoCambio = replace(afe_TipoCambio,".",",")
		afe_NPV = 0
		afe_Irr = 0
		
		afe_Confirmado = AFE_NO_CONFIRMADO
		afe_IDArea	 = 0
		afe_IDDetalle= 0
		
		afe_ROIC = 0
		afe_PAYBACK = 0
		afe_isCFO = TIPO_NEGACION
		Call Loadsignatories(0)
		
	end if
else
	
		afe_IdAFE = GF_Parametros7("idAFE",0,6)
		afe_CdAFE = GF_Parametros7("cdAFE","",6)
		afe_IdObra = GF_Parametros7("idObra",0,6)
		afe_ObraCD = GF_Parametros7("cdObra","",6)
		'JAS --> Valores leidos por respeto de herencia para AFEs viejos.
		'afe_ObraDS = GF_Parametros7("dsObra","",6)
		'<-- JAS
		afe_IdPedido = GF_Parametros7("idPedido",0,6)
		afe_NroAFEComplID = GF_Parametros7("nroAFEComplID",0,6)				
		afe_IdDivision = GF_Parametros7("idDivision",0,6)
		afe_ObraDivDS = GF_Parametros7("obraDivDS","",6)
		afe_Categoria = GF_Parametros7("categoria","",6)
		afe_CatOtros = GF_Parametros7("catOtros","",6)
		afe_Tipo = GF_Parametros7("tipo","",6)
		afe_TipoOtros = GF_Parametros7("tipoOtros","",6)
		afe_TipoCC = GF_Parametros7("cumplimientos","",6)
		afe_Descripcion = GF_Parametros7("descripcion","",6)
		afe_Departamento    = GF_Parametros7("departamento"    ,"",6)
		afe_Titulo			= GF_Parametros7("afe_titulo"    ,"",6)
		
		afe_ImportePesos = GF_Parametros7("importePesos","",6)
		if afe_ImportePesos = "" then afe_ImportePesos = 0
		afe_ImportePesos = replace(afe_ImportePesos,",",".")
		afe_ImportePesos = afe_ImportePesos * 100

		afe_ImporteDolares = GF_Parametros7("importeDolares","",6)
		if afe_ImporteDolares = "" then afe_ImporteDolares = 0
		afe_ImporteDolares = replace(afe_ImporteDolares,",",".")		
		afe_ImporteDolares = afe_ImporteDolares * 100
		
		afe_TipoCambio = GF_Parametros7("tipoCambio","",6)
		if afe_TipoCambio = "" then afe_TipoCambio = 0
		afe_TipoCambio = replace(afe_TipoCambio,",",".")		
		afe_TipoCambio = afe_TipoCambio

		afe_NPV = GF_Parametros7("Arr","",6)
		if afe_NPV = "" then afe_NPV = 0
		afe_NPV = replace(afe_NPV,",",".")		
		afe_NPV = afe_NPV * 100

		afe_Irr = GF_Parametros7("Irr","",6)
		if afe_Irr = "" then afe_Irr = 0
		afe_Irr = replace(afe_Irr,",",".")		
		afe_Irr = afe_Irr * 100

		afe_Confirmado	= GF_Parametros7("confirmado","",6)
		afe_ObraFechaFin= GF_Parametros7("obraFechaFin","",6)
		afe_IdProveedor	= GF_Parametros7("idProveedor",0,6)
		afe_DsProveedor	= GF_Parametros7("dsProveedor","",6)
		afe_ObraFechaFin= GF_Parametros7("obraFechaFin","",6)
		afe_ObraRespCD	= GF_Parametros7("obraRespCD","",6)
		afe_IDArea		= GF_Parametros7("idArea",0,6)		
		afe_IDDetalle	= GF_Parametros7("idDetalle",0,6)
		
		afe_ROIC = GF_Parametros7("RONA",0,6)		
		afe_ROIC = replace(afe_ROIC,",",".")		
		afe_ROIC = afe_ROIC * 100
				
		afe_PAYBACK = GF_Parametros7("PAYBACK",0,6)		
		afe_PAYBACK = replace(afe_PAYBACK,",",".")		
		afe_PAYBACK = afe_PAYBACK * 100
		
        chkFinanzas = GF_Parametros7("chkFinanzas","",6)
        if (UCase(chkFinanzas)  = "ON") then 
            afe_isCFO = TIPO_AFIRMACION
        else
            afe_isCFO = TIPO_NEGACION 
        end if
		Call Loadsignatories(0)		
		
end if	
end sub
'---------------------------------------------------------------------------------------------
Function Loadsignatories(pIdAFE)
	Dim rs,conn,strSQL
	
	afe_PreparedBy      = ""
	afe_RequestedBy     = ""
	afe_EngReview       = ""
	afe_Officer         = ""
	afe_VicePresident   = ""
	afe_President       = ""
	afe_Controller      = "" 'Se inicializa con Sergio Eguivar debido a que es el controller actual.
	afe_Auditor         = "" 'Auditoria no firma mas desde el 08/08/2012. Se conserva su por compatibilidad con AFEs viejos.
	afe_cfo				= "" 
		
	afe_PreparedByCD    = ""
	afe_RequestedByCD   = ""
	afe_EngReviewCD     = ""
	afe_OfficerCD       = ""
	afe_VicePresidentCD = ""
	afe_PresidentCD     = ""	
	afe_ControllerCD    = ""	'Se inicializa con Sergio Eguivar debido a que es el controller actual.
	afe_AuditorCD       = "" 'Auditoria no firma mas desde el 08/08/2012. Se conserva su por compatibilidad con AFEs viejos.
	afe_cfoCD	        = ""

	if (not isFormSubmit()) then		
		strSQL = "SELECT IDAFE,CDUSUARIO,CASE WHEN HKEY IS NULL THEN '' ELSE HKEY END AS HKEY,CASE WHEN FECHAFIRMA IS NULL THEN '' ELSE FECHAFIRMA END AS FECHAFIRMA FROM TBLAFEFIRMAS WHERE IDAFE = " & pIdAFE & " ORDER BY SECUENCIA "
        
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
        'PREPARA AFE
        if not rs.Eof then    
            afe_PreparedByCD    	= rs("CDUSUARIO")
		    afe_PreparedByHkey  	= rs("HKEY")
		    afe_PreparedByHkeyDate  = rs("FECHAFIRMA")
		    afe_PreparedBy = getUserDescription(afe_PreparedByCD)	              
            rs.MoveNext()
        end if
        'REQUIERE AFE
        if not rs.Eof then
            afe_RequestedByCD    	= rs("CDUSUARIO")
		    afe_RequestedByHkey  	= rs("HKEY")
		    afe_RequestedByHkeyDate = rs("FECHAFIRMA")
		    afe_RequestedBy = getUserDescription(afe_RequestedByCD)		
            rs.MoveNext()
        end if
        'REVISION TECNICA
        if not rs.Eof then
            afe_EngReviewCD     	= rs("CDUSUARIO")
		    afe_EngReviewHkey   	= rs("HKEY")
            afe_EngReviewHkeyDate	= rs("FECHAFIRMA")
		    afe_EngReview = getUserDescription(afe_EngReviewCD)		
            rs.MoveNext()
        end if
        'GERENTE DE PUERTOS
        if (afe_IdDivision <> DIVSION_EXPORTACION) then
            'SOLO SE NECESITA LA FIRMA DEL GERENTE DE PUERTO CUANDO LA DIVISION NO SEA DE EXPORTACION
            if not rs.Eof then
                afe_OfficerCD        	= rs("CDUSUARIO")
			    afe_OfficerHkey    	 	= rs("HKEY")
			    afe_OfficerHkeyDate  	= rs("FECHAFIRMA")
			    afe_Officer = getUserDescription(afe_OfficerCD)		
                rs.MoveNext()
            end if
        end if
        'COORDINADOR DE PUERTOS
        if not rs.Eof then
            afe_VicePresidentCD      = rs("CDUSUARIO")
		    afe_VicePresidentHkey  	 = rs("HKEY")
		    afe_VicePresidentHkeyDate= rs("FECHAFIRMA")
		    afe_VicePresident = getUserDescription(afe_VicePresidentCD)				
            rs.MoveNext()
        end if
        'CONTROLLER
        if not rs.Eof then
            afe_ControllerCD    	= rs("CDUSUARIO")
		    afe_ControllerHkey  	= rs("HKEY")
		    afe_ControllerHkeyDate  = rs("FECHAFIRMA")
		    afe_Controller = getUserDescription(afe_ControllerCD)	
            rs.MoveNext()
        end if
        'FINANZAS         
        if (afe_isCFO = TIPO_AFIRMACION) then
            'SI AL CARGAR EL AFE TILDO PARA QUE FIRMEN FINANZAS
            if not rs.Eof then
                afe_cfoCD       = rs("CDUSUARIO")
			    afe_cfoHkey  	= rs("HKEY")
			    afe_cfoHkeyDate = rs("FECHAFIRMA")
			    afe_cfo = getUserDescription(afe_cfoCD)	
                rs.MoveNext()
            end if
        end if
        'PRESIDENTE
        if not rs.Eof then
            afe_PresidentCD     	= rs("CDUSUARIO")
		    afe_PresidentHkey  		= rs("HKEY")
		    afe_PresidentHkeyDate   = rs("FECHAFIRMA")
		    afe_President = getUserDescription(afe_PresidentCD)		
            rs.MoveNext()
        end if
	else
		afe_PreparedBy      = GF_Parametros7("preparedBy"      ,"",6)		
		afe_RequestedBy     = GF_Parametros7("requestedBy"     ,"",6)
		afe_EngReview       = GF_Parametros7("engReview"       ,"",6)
		afe_Officer         = GF_Parametros7("officer"         ,"",6)
		afe_VicePresident   = GF_Parametros7("vicePresident"   ,"",6)
		afe_President       = GF_Parametros7("president"       ,"",6)
		afe_Controller		= GF_Parametros7("controller"	   ,"",6)
		afe_cfo				= GF_Parametros7("cfo"			   ,"",6)		
		
		afe_PreparedByCD    = GF_Parametros7("preparedByCD"    ,"",6)
		afe_RequestedByCD   = GF_Parametros7("requestedByCD"   ,"",6)
		afe_EngReviewCD     = GF_Parametros7("engReviewCD"     ,"",6)
		afe_OfficerCD       = GF_Parametros7("officerCD"       ,"",6)
		afe_VicePresidentCD = GF_Parametros7("vicePresidentCD" ,"",6)
		afe_PresidentCD     = GF_Parametros7("presidentCD"     ,"",6)
		afe_ControllerCD	= GF_Parametros7("controllerCD"	   ,"",6)
		afe_cfoCD		    = GF_Parametros7("cfoCD"		   ,"",6)
	end if
	
end function
'---------------------------------------------------------------------------------------------
Function readAFEPedido(idPedido, level)
	dim strSQL, rs, conn, cond
	
	if (level <> AFE_TODOS) then cond = " and NROAFECOMPL = 0"
	strSQL="Select * from TBLDATOSAFE where IDPEDIDO=" & idPedido & cond & " order by CDAFE"
	'response.write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readAFEPedido = rs
End Function
'---------------------------------------------------------------------------------------------
Function readAFEObra(idObra, level)
	dim strSQL, rs, conn
	if (level <> AFE_TODOS) then cond = " and NROAFECOMPL = 0"
	strSQL="Select * from TBLDATOSAFE where IDOBRA=" & idObra & cond & " order by CDAFE"
	'response.write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readAFEObra = rs
End Function
'---------------------------------------------------------------------------------------------
'Lee todos los AFE que son complementario al AFE indicado por parametro.
Function listaAFESComplementarios(idAFE)
	Dim rs, strSQL, conn
	
	strSQL="Select * from TBLDATOSAFE where NROAFECOMPL=" & idAFE & " order by CDAFE"
	'response.write strSQL
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set listaAFESComplementarios = rs
End Function
'---------------------------------------------------------------------------------------------
'Controlar que la partida presupuestaria exista.
Function controlPartidaPresupuestariaAFE(pIdObra, pIdArea, pIdDetalle)
	Dim strSQL, conn, rsBudget,ret
	
	ret = true	
	if (pIdObra <> 0) then		
		'Si se eligió una obra, si es de inversión no requiere area y detalle.
		if ((pIdArea = 0) or (pIdDetalle = 0)) then
			'Area detalle requerida para obra de mantenimiento	
			if (isMantenimiento(pIdObra)) then 
				ret = false
				Call setError(MANT_AREA_DET_OBLIGATORIO)
			end if
		else
			'Si esta especificando un area, es obligatorio un detalle.
			strSQL="Select * from TBLBUDGETOBRAS where IDOBRA=" & pIdObra & " and IDAREA=" & pIdArea & " and IDDETALLE=" & pIdDetalle
			Call executeQueryDb(DBSITE_SQL_INTRA, rsBudget, "OPEN", strSQL)
			if (rsBudget.eof) then 
				Call setError(BUDGET_NO_EXISTE)
				ret = false
			end if
		end if			
	end if
	controlPartidaPresupuestariaAFE = ret
	
End Function
'-----------------------------------------------------------------------------------------------
'- - - - - - - - - - - - - - 
' Modifico	: 	Javier A. Scalisi
' Fecha		:	18/09/2012
' Cambios	:	Se agregó la firma del CFO
'---------------------------------------------------------------------------------------------
function ControlAFE(pAFEComplID, pIdDivision, pCategoria, pCatOtros, pTipo, pTipoOtros, pTipoCC, pDescripcion, pTitulo,pArea,pDetalle,pIdObra, pPreparedBy, pRequestedBy, pEngReview)

ControlAFE = true
'Seccion 1
if pCategoria = "" then
	Call setError(FALTA_CATAGORIA_AFE)
	ControlAFE = false
elseif pCategoria = "A" and CLng(pAFEComplID) = 0 then
	Call setError(FALTA_AFE_COMPLEMENTARIO)	
	ControlAFE = false
elseif pCategoria = "O" and pCatOtros = "" then
	Call setError(FALTA_TEXTO_OTROS)		
	ControlAFE = false
end if
'Seccion 2
if pTipo = "" then
	Call setError(FALTA_TIPO_GASTO)			
	ControlAFE = false
elseif instr(pTipo,"C")>0 and pTipoCC = "" then
	Call setError(FALTA_TIPO_CUMPLIMIENTO)			
	ControlAFE = false
elseif instr(pTipo,"O")>0 and pTipoOtros = "" then
	Call setError(FALTA_TEXTO_OTROS)			
	ControlAFE = false
end if	
'Seccion 3
if len(pDescripcion) > 4000 then
	Call setError(DESCRIPCION_DEMASIADO_LARGA)	
	ControlAFE = false
end if	
if pTitulo = "" then
	Call setError(FALTA_TITULO_AFE)	
	ControlAFE = false
end if

ControlAFE = controlPartidaPresupuestariaAFE(pIdObra, pArea, pDetalle)

if (not controlFirmasAFE(pPreparedBy, pRequestedBy, pEngReview)) then
	Call setError(FALTA_MIEMBROS_AFE)
	ControlAFE = false
end if

end function
'---------------------------------------------------------------------------------------------
' Autor: 	Santi Juan Pablo
' Fecha: 	07/11/2010
' Objetivo:	Controlar que existan todos los miembros del comite.


'Parametros =	cdusuario de los miembros que aprueban el afe (ej: 'JPS')

' Devuelve:		True/False
'- - - - - - - - - - - - - - 
' Modifico	: 	Javier A. Scalisi
' Fecha		:	18/09/2012
' Cambios	:	Se agregó la firma del CFO
'- - - - - - - - - - - - - - 
' Modifico	: 	Nahuel Ajaya
' Fecha		:	21/10/2014
' Cambios	:	Se modifico para que solo se validen los que firman por Usuario.
'---------------------------------------------------------------------------------------------
Function controlFirmasAFE(pPreparedBy, pRequestedBy, pEngReview)
	Dim i, rtrn
	rtrn = false
	if((pPreparedBy <> "")and(pRequestedBy <> "")and(pEngReview <> "")) then rtrn = true
	controlFirmasAFE = rtrn
End Function
'---------------------------------------------------------------------------------------------
'Parametros =	nroAFECompl	[INT]		corresponde al id del afe que se complementa
'				nroAFEAnula	[INT]		corresponde al id del afe que se anula
'---------------------------------------------------------------------------------------------
Function generateCdAFE(pIdDivision, nroAFECompl, nroAFEAnula)
	if (nroAFEAnula > 0) then
		generateCdAFE = generateCdAFEAnulado(nroAFEAnula)
	elseif (nroAFECompl > 0) then
		generateCdAFE = generateCdAFECompl(nroAFECompl)
	else
		generateCdAFE = generateCdAFENew(pIdDivision)
	end if
End Function
'---------------------------------------------------------------------------------------------
Function getCdAFE(pIdAFE) 
	Dim cdAFE, conn , rsAFE, strSQL
	cdAFE = "NA"
	if (pIdAFE <> "") then
		strSQL="Select CDAFE from TBLDATOSAFE where IDAFE=" & pIdAFE
        Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE, "OPEN", strSQL)
			if (not rsAFE.eof) then cdAFE = rsAFE("CDAFE")
	end if
	getCdAFE = cdAFE
End Function
'---------------------------------------------------------------------------------------------
'Funcion solo para uso interno por las funciones esEditable, getEfitAFEIcon, getRejectAFEIcon!!! 
Function readEditableAfe(pIdAFE)
	Dim rtrn,rsAFE,conn,strSQL, gastosPedidoPesos, gastosPedidoDolares, gastosObra
		
	strSQL="Select CDUSUARIO, IDDIVISION, IDOBRA, IDPEDIDO, IDAREA, IDDETALLE from TBLDATOSAFE where IDAFE=" & pIdAFE &_
				" and CONFIRMADO not in ('" & AFE_ANULADO & "', '" & AFE_ANULACION & "')"
	Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE, "OPEN", strSQL)
	'Si el AFE esta un estado valido, calculo la cantidad de dinero gastado asociado al AFE.
	if (not rsAFE.eof) then	    
	    if (session("AFE_EDITABLE_" & pIdAFE) <> "") then
	        rsAFE.MoveNext()
        else	        
	        if (CLng(rsAFE("IDPEDIDO")) <> 0) then	            
                Call loadImporteAcumuladoPIC(rsAFE("IDPEDIDO"), 0, 0, true, gastosPedidoPesos, gastosPedidoDolares)                
                if (gastosPedidoDolares > 0) then
                    'Hubo gastos, no puede editarse.
                    session("AFE_EDITABLE_" & pIdAFE) = "X"
                    rsAFE.MoveNext()
                end if	            
	        end if
	        if (CLng(rsAFE("IDOBRA")) <> 0) then
                gastosObra = calcularGastosObra(MONEDA_DOLAR, rsAFE("IDOBRA"), rsAFE("IDAREA"), rsAFE("IDDETALLE"), true)            
                if (gastosObra > 0) then
                    'Hubo gastos, no puede editarse.
                    session("AFE_EDITABLE_" & pIdAFE) = "X" 
                    rsAFE.MoveNext()
                end if
	        end if        	        
	    end if	    
	end if	
	Set readEditableAfe = rsAFE
End Function
'---------------------------------------------------------------------------------------------
Function esEditable(pIdAFE)
	Dim rtrn,rsAFE,conn,strSQL
	rtrn = false	
	if (pIdAFE = 0) then 
		rtrn = true	
	else	
		Set rsAFE = readEditableAfe(pIdAFE)
		if (not rsAFE.eof) then rtrn = true	
	end if
	esEditable = rtrn
End Function
'---------------------------------------------------------------------------------------------
'Devuelve el ícono y su correspondiente codigo para permitir la edición del AFE.
'Para funcionar la pagina que utilice esta función debe tener una función en javascript que se llame editAFE que reciba el ID de afe a editar.
Function getEditAFEIcon(pIdAFE)
	Dim rtrn,rsAFE,conn,strSQL

	rtrn = ""	
	Set rsAFE = readEditableAfe(pIdAFE)
	if (not rsAFE.eof) then 
		'Se va a mostrar el ícono de edición, solo permito realmente editar si es el usuario que creó el AFE o un administrador.
		rtrn =	"<span style='cursor:pointer;' onclick='"		
		if ((rsAFE("CDUSUARIO") = session("usuario")) or (isAdmin(rsAFE("IDDIVISION")))) then
			rtrn =  rtrn & "editAFE(" & pIdAFE & ")"
		else
			dsUsuario = getUserDescription(rsAFE("CDUSUARIO"))
			rtrn = rtrn & "alert('Solo el usuario " & dsUsuario & " o un Administrador pueden modificar el AFE seleccionado.')"
			
		end if
		rtrn = rtrn & "'><img src='images\edit-16.png' alt='Afe16x16'  title='" & GF_TRADUCIR("Modificar AFE") & "'></span>"	
	end if
	getEditAFEIcon = rtrn
End Function
'----------------------------------------------------------------
'Devuelve el ícono y su correspondiente codigo para permitir la edición del AFE.
'Para funcionar la pagina que utilice esta función debe tener una función en javascript que se llame anularAFE que reciba el ID de afe a anular.
Function getRejectAFEIcon(pIdAFE)
	Dim rtrn,rsAFE,conn,strSQL, dsUsuario

	rtrn = ""	
	Set rsAFE = readEditableAfe(pIdAFE)
	if (not rsAFE.eof) then 
		'Se va a mostrar el ícono de edición, solo permito realmente editar si es el usuario que creó el AFE o un administrador.
		rtrn =	"<span style='cursor:pointer;' onclick='"		
		if ((rsAFE("CDUSUARIO") = session("usuario")) or (isAdmin(rsAFE("IDDIVISION")))) then
			rtrn =  rtrn & "anularAFE(" & pIdAFE & ")"
		else
			dsUsuario = getUserDescription(rsAFE("CDUSUARIO"))
			rtrn = rtrn & "alert('Solo el usuario " & dsUsuario & " o un Administrador pueden anular el AFE seleccionado.')"
			
		end if
		rtrn = rtrn & "'><img src='images/cross-16.png' title='" & GF_TRADUCIR("Anular AFE") & "'></span>"	
	end if
	getRejectAFEIcon = rtrn
End Function
'----------------------------------------------------------------
'Suma los AFE que se aprobaron para el pedido.
'Se permite calcular el total teniendo en cuenta o no, el desvio autorizado para los gatos respecto al valor real del AFE.
Function totalizarAFESPedido(idMoneda, idPedido, pContarDesvio)
	Dim strSQL, rs, conn, importe, margen
	importe = 0
	
	if (idPedido <> 0) then
		strSQl = " Select sum(IMPORTEPESOS) IMPORTEPESOS, sum(IMPORTEDOLARES) IMPORTEDOLARES "
		strSQL = strSQL & " FROM tbldatosafe"
		strSQL = strSQL & " WHERE idPedido = " & idPedido
		strSQL = strSQL & " AND confirmado ='" & AFE_APROBADO &"'"
		'response.write strSQL	
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			if (rs("IMPORTEPESOS") <> "") then
				importe = cdbl(rs("IMPORTEPESOS"))
				if (idMoneda = MONEDA_DOLAR) then importe = cdbl(rs("IMPORTEDOLARES"))
			end if
		end if
		if (not IsNumeric(importe)) then importe = 0			
		if (pContarDesvio) then
			'Calculo el margen autorizado de exceso del AFE: Porcentual o absoluto, el menor de ambos (EN DOLARES!!).
			margen = importe*AFE_MAX_DESVIO_PCN
			if (margen > AFE_MAX_DESVIO_ABS) then margen=AFE_MAX_DESVIO_ABS
			importe = importe + margen
		end if	
	end if
	totalizarAFESPedido = importe
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	EAB - Ezequiel Bcarini
' Fecha: 	23/07/13
' Objetivo:	
'			Informa si un determinado pedido tiene AFEs asociados
' Parametros:
'			idPedido	[int]   Id del pedido de cotizacion a analizar.
' Devuelve:
'			True/False
function tieneAFEPedido(pIdPedido)
tieneAFEPedido = false
if (pIdPedido <> 0) then
		strSQl = " SELECT COUNT(*) AS QUANTITY FROM TBLDATOSAFE"
		strSQL = strSQL & " WHERE IDPEDIDO = " & pIdPedido
		strSQL = strSQL & " AND CONFIRMADO ='" & AFE_APROBADO & "'"
	    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if cint(rs("QUANTITY"))>0 then tieneAFEPedido = true
end if		
end function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	--/--/10
' Objetivo:	
'			Controla si se debe generar un AFE para una compra.
' Parametros:
'			idObra 		[int]	Id de la partida presupuestaria
'			idPedido	[int]   Id del pedido de cotizacion (por si no tiene obra.
'			pImporte 	[int] 	Importe del nuevo gasto a incluir en el AFE.
'			pArea		[int]	Area de la obra a comprobar
'			pDetalle	[int]	Detalle de la obra a comprobar
' Devuelve:
'			True/False
' Modificaciones:
'			20/10/2010 - GFG
'			16/11/2010 - GFG - Optimizacion
'			28/03/2011 - JAS - Se modificó para ignorar las obra de mantenimiento hasta que se resuelva el correcto manejo de las mismas.
'			23/09/2011 - JAS - Se modificó para controlar PCT sin partida presupuestaria.
'			25/11/2011 - JAS - Se modificó para no controlar los PICs relacionados con contratos.
'			08/04/2013 - CNA - Se modificó para que se controle desde la tabla ctzCabecera si el Pic tiene asociado un contrato
'			23/07/2013 - EAB - Se altero el orden de las prioridades a la hora de controlar
'--------------------------------------------------------------------------------------------------
Function necesitaAFE(idObra,idPedido, idPIC, pImporte,pArea,pDetalle)
	Dim seguir, strSQL, rs, importeCompra, tipoCambio, importeViejo
	
	necesitaAFE = false
    importeCompra = pImporte				
	'Si hago esta pregunta al modificar un PIC (idPIC <> 0), 
	'si el PIC pertenece a un pago de un contrato no hay que controlar 
	'dado que el control se realizó al autorizar el contrato.
	seguir = true
	if (idPIC <> 0) then	
		strSQL="Select * from TBLCTZCABECERA where IDCOTIZACION = " & idPIC
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (rs("IDCONTRATO") > 0) then  
		    seguir = false
        else
            'Como se especifica un PIC como parametro implica que es una modificacion.
            'Al realizar el c{alculo se va a tomar el valor del PIC ya guardado, se debe verificar si hay saldo disponible solo para la diferencia.
            'Si la modificacion desminuye el valor del PIC, no controlar ya que es menos plata que antes y seguro alcanza el saldo del AFE.            
            importeViejo = CDbl(rs("IMPORTEDOLARES"))
            if (rs("CDMONEDA") = MONEDA_PESO) then 
                tipoCambio = CDbl(getTipoCambio(MONEDA_DOLAR, ""))
                importeViejo = CDbl(rs("IMPORTEPESOS")) / tipoCambio
            end if                
            importeCompra = importeCompra - importeViejo
            if (importeCompra <=0) then  seguir = false
        end if		    
	end if
	if (seguir) then
		'El pedido tiene prioridad
		if (idPedido > 0 and tieneAFEPedido(idPedido)) then	
			necesitaAFE = necesitaAFEPedido(idPedido, importeCompra)	
		else
			if (idObra = 0) then
				'No tiene ni obra ni pedido => Compra Directa!
				necesitaAFE = necesitaAFEPIC(idPic, importeCompra)							
			else
				'No tiene pedido, tiene obra
				necesitaAFE = necesitaAFEObra(idObra, pArea, pDetalle, importeCompra)
			end if
		end if			
	end if	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	23/09/2011
' Objetivo:	
'			Controla si se debe generar un AFE para una compra asociada a una compra directa.
' Parametros:
'			idPedido	[int]   Id del pedido de cotizacion (por si no tiene obra.
'			pImporte 	[int] 	Importe del nuevo gasto a incluir en el AFE.
' Devuelve:
'			True/False
Function necesitaAFEPIC(idPIC, pImporte)

	Dim strSQL, rs, importeCompra, limite, unidad, tipoCambio
	
	necesitaAFEPIC = false	
	'Determino la norma de auditoria que se debe cumplir, esto depende de la obra.	
	importeCompra = cdbl(pImporte)
	tipoCambio = CDbl(getTipoCambio(MONEDA_DOLAR, ""))		
	limite = CDbl(getValorNorma("VLAFEM"))*100								
	unidad = getUnidadNorma("VLAFEM")
	'El limite debe estar en dolares.
	if (unidad = MONEDA_PESO) then	limite = limite/tipoCambio	
	if (importeCompra > limite) then  necesitaAFEPIC = TRUE
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	23/09/2011
' Objetivo:	
'			Controla si se debe generar un AFE para una compra asociada a un pedido de precio.
' Parametros:
'			idPedido	[int]   Id del pedido de cotizacion (por si no tiene obra.
'			pImporte 	[int] 	Importe del nuevo gasto a incluir en el AFE.
' Devuelve:
'			True/False
Function necesitaAFEPedido(idPedido, pImporte)
	Dim limite, importeCompra, unidad,  totalAFE, saldoPedido, gastosPedidoPesos, gastoPedidoDolares, tipoCambio
	Dim cdMoneda, importePCP, cdIdProveedor
	
	necesitaAFEPedido = false
	
	importeCompra = cdbl(pImporte)	
		
	'Se calculan los AFEs.
	totalAFE = totalizarAFESPedido(MONEDA_DOLAR, idPedido, true)	
	if (totalAFE > 0) then
		'Si ya tiene AFE asumo que es por que lo necesita. Valido que el pedido tenga saldo disponible.
		Call loadImporteAcumuladoPIC(idPedido, 0, 0, true, gastosPedidoPesos, gastosPedidoDolares)
		saldoPedido = totalAFE - gastosPedidoDolares				
		if ( (saldoPedido < 0) or (importeCompra > saldoPedido) ) then necesitaAFEPedido = TRUE			
	else
		'No hay ningún AFE cargado, verifico si debería haberlo.
		tipoCambio = CDbl(getTipoCambio(MONEDA_DOLAR, ""))		
		
		'Determino la norma de auditoria que se debe cumplir, esto depende de la obra.	
		limite = CDbl(getValorNorma("VLAFEM"))*100								
		unidad = getUnidadNorma("VLAFEM")		
		'El limite debe estar en dolares.
		if (unidad = MONEDA_PESO) then	limite = limite/tipoCambio	

		pct_idPedido = idPedido
		Call obtenerGanadorPlanilla(cdMoneda, importePCP, cdIdProveedor)
		budgetPedido = importePCP
		if (cdMoneda = MONEDA_PESO) then budgetPedido = budgetPedido/tipoCambio
		if (budgetPedido > limite) then necesitaAFEPedido = true
	end if
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	23/09/2011
' Objetivo:	
'			Controla si se debe generar un AFE para una compra asociada a una partida presupuestaria.
' Parametros:
'			idObra 		[int]	Id de la partida presupuestaria
'			pImporte 	[int] 	Importe del nuevo gasto a incluir en el AFE.
'			pArea		[int]	Area de la obra a comprobar
'			pDetalle	[int]	Detalle de la obra a comprobar
' Devuelve:
'			True/False
Function necesitaAFEObra(idObra, pArea, pDetalle, pImporte)

	Dim limite, importeCompra, unidad,  totalAFE, saldoObra, gastosObra, tipoCambio
	Dim myIdArea, myIdDetalle

	necesitaAFEObra = false
	
	importeCompra = cdbl(pImporte)	
	tipoCambio = CDbl(getTipoCambio(MONEDA_DOLAR, ""))
	
	'Determino la norma de auditoria que se debe cumplir, esto depende de la obra.	
	limite = CDbl(getValorNorma("VLAFEM"))*100								
	unidad = getUnidadNorma("VLAFEM")		
	'El limite debe estar en dolares.
	if (unidad = MONEDA_PESO) then	limite = limite/tipoCambio	
	 					
	if (isInversion(idObra)) then
		'Response.Write "<BR>ES Obra"
		limite = CDbl(getValorNorma("VLAFE"))*100						
		unidad = getUnidadNorma("VLAFE")		
		'El limite debe estar en dolares.
		if (unidad = MONEDA_PESO) then	limite = limite/tipoCambio	
		'Se toma para el cálculo toda la obra y no solo un item de la misma.
		myIdArea = 0
		myIdDetalle = 0
	'JAS -->
	'else
		'Es mantenimiento						
	'	myIdArea = pArea
	'	myIdDetalle = pDetalle
	'end if
	'<-- JAS
		'Se calculan los AFEs.
		totalAFE= totalizarAFESObra(MONEDA_DOLAR, idObra, myIdArea, myIdDetalle, TRUE)
		if (totalAFE > 0) then
			'Valido que la obra tenga saldo disponible.
			gastosObra = calcularGastosObra(MONEDA_DOLAR, idObra, myIdArea, myIdDetalle, true)				
			saldoObra = totalAFE - gastosObra				
			if ( (saldoObra < 0) or (importeCompra > saldoObra) ) then necesitaAFEObra = TRUE
		else
			'obtengo el presupuesto del budget
			budgetObra = calcularCostoEstimadoObra(MONEDA_DOLAR, idObra, myIdArea, myIdDetalle)			
			if (importeCompra > budgetObra) then budgetObra = importeCompra
			if (budgetObra > limite) then necesitaAFEObra = true
		end if
	else
		'JAS -	Si es una obra de mantenimiento no se solicita AFE, esto será así hasta que se pueda saber si
		'		una partida incluye varios trabajos o es una único tarea, solo así se puede saber si se necesita o no un AFE.
		necesitaAFEObra = false
	end if
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Santi Juan Pablo
' Fecha: 	28/10/10
' Parametros: IdAfe
' Devuelve: Proximo Usuario a Firmar AFE
'-------------------------------------------------------------------------------------------------
Function getUsuarioAFirmar(idAFE)
	Dim aux, rs, conn, strSQL
    aux="ERROR"
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLAFEFIRMAS_GET_NEXT_USER", idAFE)
	if (not rs.eof) then aux = rs("CDUSUARIO")
	getUsuarioAFirmar = aux
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	17/11/2014
' Parametros: IdAfe
' Devuelve: Proximo descripcion Usuario o Rol a Firmar AFE
'-------------------------------------------------------------------------------------------------
Function getDSUsuarioAFirmar(idAFE)
	Dim aux, rs, conn, strSQL
    aux="ERROR"
    
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLAFEFIRMAS_GET_NEXT_USER", idAFE)
	if (not rs.eof) then 
	    if (rs("IDROL") = FIRMA_ROL_NINGUNO) then
	        aux = getUserDescription(rs("CDUSUARIO"))
	    else
	        aux = rs("DSROL")
	    end if
	end if
	getDSUsuarioAFirmar = aux
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Santi Juan Pablo
' Fecha: 	28/10/10
' Parametros: pIdDivision
' Devuelve: Nuevo cd afe
'-------------------------------------------------------------------------------------------------
Function generateCdAFENew(pIdDivision)
	Dim cdDivision, rsObra, rsDivision, strSQL, dte, mes

	cdDivision = "NA"
	strSQL="Select * from TBLDIVISIONES where IDDIVISION=" & pIdDivision
    Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
	if (not rsDivision.eof) then cdDivision = rsDivision("CDDIVISION")
	Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "CLOSE", strSQL)
				
	dte = Right(year(date),2)	
	strSQL="Select * from TBLNUMERACION where CLAVE='" & cdDivision & "_" & dte & "' and PREFIJO='" & PREFIX_AFE & "'"
    Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
	if not rsDivision.eof then
		afe_idAFE = clng(rsDivision("Valor")) + 1
		strsql = "Update TBLNUMERACION set VALOR=" & afe_idAFE & " where CLAVE = '" & cdDivision & "_" & dte & "' and PREFIJO='" & PREFIX_AFE & "'"	
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strsql)
	else
		afe_idAFE = 1
		strSQL = "Insert into TBLNUMERACION values('" & PREFIX_AFE & "','" & cdDivision & "_" & dte & "',1)"
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strsql)
	end if
	generateCdAFENew = dte & "-ARG-" & cdDivision & "-" & GF_nDigits(afe_idAFE, 3)
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Santi Juan Pablo
' Fecha: 	28/10/10
' Parametros: nroAFECompl [numero de afe al que complementa]
' Devuelve: Nuevo cd afe complementario
'-------------------------------------------------------------------------------------------------
Function generateCdAFECompl(nroAFECompl)
	Dim codigoCompl, rsAFECompl, cdAFECompl, rs, contRs

	contRs = 0

	strSQL = "Select count(*) as CANTIDAD from TBLDATOSAFE where NROAFECOMPL = " & nroAFECompl
	Call executeQueryDb(DBSITE_SQL_INTRA, rsAFECompl, "OPEN", strSQL)
	if (not rsAFECompl.eof) then contRs = cDbl(rsAFECompl("CANTIDAD"))
	cdAFECompl = getCdAFE(nroAFECompl)
	codigoCompl = contRs + 1
	generateCdAFECompl = cdAFECompl & "/" & codigoCompl
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Santi Juan Pablo
' Fecha: 	28/10/10
' Parametros: nroAFEAnula [numero de afe que se anula]
' Devuelve: Nuevo cd afe anulado
'-------------------------------------------------------------------------------------------------
Function generateCdAFEAnulado(nroAFEAnula)
	Dim cdAFEAnulado
	cdAFEAnulado = getCdAFE(nroAFEAnula)
	generateCdAFEAnulado = cdAFEAnulado & "-A"
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier A. Scalisi
' Fecha: 	18/09/2012
' Parametros: pTipo		Codigo de tipo de AFE
' Devuelve: Descripción del tipo indicado.
'-------------------------------------------------------------------------------------------------
Function getDescripcionTipoAFE(pTipo)
			
	Select case(pTipo)
		Case AFE_TIPO_MEJORA
			ret = "MEJORA DE EFICIENCIA"
		Case AFE_TIPO_REPUESTOS
			ret = "REPUESTOS"
		Case AFE_TIPO_COMUNICACIONES
			ret = "IT/TELECOMUNICACIONES"
		Case AFE_TIPO_DESVIO
			ret = "DESVIO"
		Case AFE_TIPO_CAPACIDAD
			ret = "INCREMENTO DE CAPACIDAD"
		Case AFE_TIPO_MANTENIMIENTO
			ret = "MANTENIMIENTO"
		Case AFE_TIPO_VEHICULOS
			ret = "VEHICULOS"
		Case AFE_TIPO_CAMBIO_OBJETIVO
			ret = "CAMBIO DE OBJETIVO"
		Case AFE_TIPO_CUMPIMIENTO
			ret = "CUMPLIMIENTO CON (SELECCIONE UNO)"
		Case AFE_TIPO_CUMPLIMIENTO_NC
			ret = "NORMAS DE CALIDAD"
		Case AFE_TIPO_CUMPLIMIENTO_MA
			ret = "MEDIO AMBIENTE"
		Case AFE_TIPO_CUMPLIMIENTO_SEG
			ret = "SEGURIDAD Y SALUD"
		Case AFE_TIPO_OTROS
			ret = "OTROS:"
		Case Else
			ret = "ERROR TIPO"
	End Select		
	getDescripcionTipoAFE = GF_TRADUCIR(ret)
			
End Function
'-------------------------------------------------------------------------------------------------
' Autor: 	Javier A. Scalisi
' Fecha: 	18/09/2012
' Parametros: pCategoria		Codigo de categoria de AFE
' Devuelve: Descripción de la categoria indicada.
'-------------------------------------------------------------------------------------------------
Function getDescripcionCategoriaAFE(pCategoria)
			
	Select case(pCategoria)
		Case AFE_CATEGORIA_CAPITAL
			ret = "CAPITAL"
		Case AFE_CATEGORIA_GASTOS
			ret = "GASTOS"
		Case AFE_CATEGORIA_INVERSIONES
			ret = "INVERSIONES FINANCIERAS"
		Case AFE_CATEGORIA_SERVICIOS
			ret = "SERVICIOS DE CONSULTORIA"
		Case AFE_CATEGORIA_ALQUILER
			ret = "ALQUILER"
		Case AFE_CATEGORIA_COMPLEMENTARIO
			ret = "AFE COMPLEMENTARIO NO.:"				
		Case AFE_CATEGORIA_OTROS
			ret = "OTROS:"
		Case Else
			ret = "ERROR CATEGORIA"
	End Select
	getDescripcionCategoriaAFE = GF_TRADUCIR(ret)
			
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   necesitaAFEaprobacionHamburgo
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   06/06/2014
' Objetivo:
'			   Se encarga de verificar si el AFE debe ser aprobado por Hamburgo, dependiendo de su importe.
'				|___Si ya firmó el último usuario y el AFE es menor USD 50.000, no se necesita la aprobacion de Hamburgo 
'				|___Si ya firmó el último usuario y el AFE es mayor o igual a USD 50.000, se necesita la aprobación de Hamburgo.
' Parametros:
'			   pIdAfe			[int]	ID Afe
' Devuelve:	   
'			  True  : necesita aprobacion de Hamburgo
'			  False : NO necesita aprobacion de Hamburgo
'-------------------------------------------------------------------------------------------- 
Function necesitaAFEaprobacionHamburgo(pIdAfe)	
	Dim strSQL,rtrn	
	rtrn = false
	strSQL = "SELECT IMPORTEDOLARES FROM TBLDATOSAFE WHERE IDAFE = " &pIdAfe
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.Eof then
		'Consulta si se espera la firma de Hamburgo.
		if (Cdbl(rs("IMPORTEDOLARES")) >= (CDbl(getValorNorma("VLAFEHAM"))*100)) then rtrn = true
	end if
	necesitaAFEaprobacionHamburgo = rtrn
End Function
%>