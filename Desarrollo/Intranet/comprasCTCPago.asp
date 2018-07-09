<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<%
Const PAGOS_ITEM_CANTIDAD = 1

Call initAccessInfo(RES_CC)

'--------------------------------------------------------------------------
'controla los unicos datos que ingresa el usuario a la hora de crear el contrato
'los demas datos son asignados por el pedido
Function controlPago()
	Dim rtrn
	rtrn = false
	if (CTC_descripcion <> "") then	
		if ((CTC_idObra <> 0 and CTC_areaObra <> 0 and  CTC_detalleObra <> 0) or (CTC_idObra = OBRA_GEID)) then
			if (CTC_APagarDolares <> 0) then
				if (controlImporte()) then
				    if (controlFecha()) then
				        if (member2Cd <> "") then		
		                    rtrn = true
	                    else
					        Call setError(AUTORIZANTE_NO_EXISTE)
                        end if					        
					end if
				else
					setError(CTC_SALDO_INSUFICIENTE)
				end if
			else
				setError(IMPORTE_NO_EXISTE)
			end if
		else
			Call setError(BUDGET_NO_EXISTE)
		end if
	else
		Call setError(DESCRIPCION_VACIA)
	end if
	controlPago = rtrn
End Function
'-----------------------------------------------
Function controlFecha()
    Dim fHoy, rtrn
    
    'Para el control de la fecha se distingue entre 2 casos:
    '   1.- Emision de CEC con fecha de autorizacion = a la fecah del dia. En estos casos no se hace control. Mientras haya saldo permite la emision.
    '   2.- Emision de un CEC para una fecha futura, aqui si se debe controlar que el CEC no exceda de la fecha de vencimiento del contrato o de la partida, la que vence primero.
    '   No se aceptan emisiones de CECs con fechas anteriores al dia actual.
    
    rtrn = true
    fHoy = Left(session("MmtoDato"), 8)
    
    '1ro valido que la fecha no sea anterior al dia de hoy.
    if (CLng(fHoy) > CLng(CTC_FechaEntrega)) then 
        setError(FECHA_ENTREGA_INCORRECTA)
        rtrn = false
    else                   
        if (CLng(fHoy) < CLng(CTC_FechaEntrega)) then
            'Es un CEC a futuro    
             'Tomo la fecha de vto del contrato y la de vencimiento de la partida.
             'Me quedo con la menor de ambas para compararla con la fecha de autorizacion del CEC, ningun CEC se puede autorizar
             'luego de finalizado un contrato
             'Ver cuando se hace el loadObra al comienzo de la pagina.
            if (CLng(CTC_fechaVto) < CLng(CTC_FechaEntrega)) then                     
                setError(FECHA_ENTREGA_INCORRECTA)
                rtrn = false
            end if
        end if            
    end if
    controlFecha = rtrn
End Function
'-----------------------------------------------
Function getItemObra(tipo)
    Dim item
    
    item = ITEM_OBRAS_EN_CURSO
	if (tipo <> CTC_TIPO_OBRA) then item = ITEM_SERVICIOS_GENERALES
	
    getItemObra = item
    
End Function
'-----------------------------------------------
Function getImporteMovimiento(cdMoneda, rsPagos)
	if (cdMoneda = MONEDA_PESO) then
		getImporteMovimiento = cDbl(rsPagos("IMPORTEPESOS"))		
	else
		getImporteMovimiento = cDbl(rsPagos("IMPORTEDOLARES"))
	end if
End Function
'-----------------------------------------------
'* Funcion: totalizarConceptosPagados
'* 
'* Totaliza los conceptos pagados para un determinado contrato.
'* Si el PIC informado es mayor a cero se asume que es una modificación de un pago existente por lo cual se ignora el mismo al momento de leer los datos registrados.
'*
'* Parametros:
'*              idContrato      [IN]    ID del contrato
'*              idPIC           [IN]    ID del PIC que se quiere ignorar de la DB, sino cero.
'*              pagadoObra      [OUT]   Total abonado por concepto de obra (+)   
'*              pagadoAnticipo  [OUT]   Total abonado por concepto Adelanto de obra (+)
'*              pagadoFReparo   [OUT]   Total retenido por concepto de Fondo de Reparo (-)   
'*
'* Autor: Javier A. Scalisi
'* Fecha: 15/05/2013
'**
Function totalizarConceptosPagados(idContrato, idObra, idArea, idDetalle, idPIC, ByRef pagadoObra, ByRef pagadoAnticipo, ByRef pagadoFReparo)
    Dim strSQL, rs
    
    pagadoObra = 0
    pagadoAnticipo = 0
    pagadoFReparo = 0

    strSQL = "          Select IDARTICULO, Sum(IMPORTEPESOS) as IMPORTEPESOS, Sum(IMPORTEDOLARES) as IMPORTEDOLARES"     
    strSQL = strSQL & " from" 
    strSQL = strSQL & "(" & readCTCPagosSQL(idContrato) & ") TABLA"    
    strSQL = strSQL & " where IDPIC <> " & idPIC & " and IDFAC = 0" 
    strSQL = strSQL & " and (IDARTICULO in (" & ITEM_ANTICIPO_OBRAS_EN_CURSO & ", " & ITEM_FONDO_REPARO_ARS & ", " & ITEM_FONDO_REPARO_USD & ", " & ITEM_FONDO_REPARO_ARS_IVA & ", " & ITEM_FONDO_REPARO_USD_IVA & ")"
    strSQL = strSQL & " or IDOBRA = " & idObra 
    strSQL = strSQL & " and IDAREA = " & idArea 
    strSQL = strSQL & " and IDDETALLE = " & idDetalle & ")"
    strSQL = strSQL & " group by IDARTICULO"      
    strSQL = strSQL & " order by idarticulo"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	'Se recorren los resultados y se asigna cada importe en la moneda que corresponde a una variable según el artículo.		
	while (not rs.eof)
		    Select case CLng(rs("IDARTICULO"))			
			Case ITEM_ANTICIPO_OBRAS_EN_CURSO
				pagadoAnticipo = getImporteMovimiento(CTC_cdMoneda, rs)
			Case ITEM_FONDO_REPARO_ARS
				pagadoFReparo = getImporteMovimiento(MONEDA_PESO, rs)
			Case ITEM_FONDO_REPARO_USD
				pagadoFReparo = getImporteMovimiento(MONEDA_DOLAR, rs)					
			Case ITEM_FONDO_REPARO_ARS_IVA	'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano en el PIC.
				pagadoFReparo = getImporteMovimiento(MONEDA_PESO, rs)								
			Case ITEM_FONDO_REPARO_USD_IVA	'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano en el PIC.
				pagadoFReparo = getImporteMovimiento(MONEDA_DOLAR, rs)
			Case else 'Cualquiera de los tipos de pago de obra
				pagadoObra= getImporteMovimiento(CTC_cdMoneda, rs)
		End Select
		rs.MoveNext()
	wend	 
End Function
'-----------------------------------------------
'Controla que el importe total de los pagos no exceda el total del contrato.
Function controlImporte() 

	Dim importeTotal, rtrn
	Dim pagoObra, pagoAnticipo, pagoFReparo
	Dim pagadoObra, pagadoAnticipo, pagadoFReparo
	Dim totalPagado, totalPagoActual
	
	rtrn = false	
	Call totalizarConceptosPagados(CTC_idContrato, CTC_IdObra, CTC_areaObra, CTC_detalleObra, CTC_idPIC, pagadoObra, pagadoAnticipo, pagadoFReparo)
	'Se toman los valores del pago actual (CTC_idPIC tiene el ID si se está modificando un pago)
	if (CTC_cdMoneda = MONEDA_DOLAR) then
	    pagoObra = CDbl(CTC_ImporteObraDolares)
	    pagoAnticipo = CDbl(CTC_AnticipoDolares)
	    pagoFReparo = CDbl(CTC_FReparoDolares) 
	    importeTotal = CDbl(CTC_ContratoDolares)
	else
	    pagoObra = CDbl(CTC_ImporteObraPesos)
	    pagoAnticipo = CDbl(CTC_AnticipoPesos)
	    pagoFReparo = CDbl(CTC_FReparoPesos) 
	    importeTotal = CDbl(CTC_ContratoPesos)	    
	end if
	totalPagado = pagadoObra + pagadoAnticipo + pagadoFReparo
	totalPagoActual = pagoObra + pagoAnticipo + pagoFReparo		
	'Se controla que el nuevo pago no exceda el total del contrato, ni el total general ni los totales parciales.	
	if (importeTotal => (totalPagado + totalPagoActual)) then		
	    if (importeTotal => (pagadoObra + pagoObra)) and (importeTotal => (pagadoAnticipo + pagoAnticipo)) then
	        'El fondo de reparo es algo que se retiene de las facturas, siempre debe ser negativo y cuando se devuelve como máximo debe quedar en cero. 	        
	        if ((pagadoFReparo + pagoFReparo) <= 0) then
	            rtrn = true
	        end if
	    end if
	end if		
	controlImporte = rtrn
	
End Function
'--------------------------------------------------------------------------
Function getDsBudget(idObra, idArea, idDetalle)
	Dim rsBudget
	getDsBudget = ""
	Set rsBudget = obtenerListaBudgetObra(idObra, idArea, idDetalle)
	if (not rsBudget.eof) then getDsBudget = rsBudget("DSBUDGET")
End Function
'--------------------------------------------------------------------------------------------
' Función:	getTipoPago
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	10/04/2013
' Objetivo:	
'			Devolver el tipo de pago de un contrato por medio de su articulo
' Parametros:
'			pIdArticulo			[int]	Id Articulo
' Devuelve:
'			tipo pago [int]
'--------------------------------------------------------------------------------------------
Function getTipoPago(pIdArticulo)
	Dim rtrn 	
	rtrn  = PAGO_OBRA		
	select case pIdArticulo
		case ITEM_ANTICIPO_OBRAS_EN_CURSO
			rtrn = PAGO_ANTICIPO
		case ITEM_FONDO_REPARO_ARS,	ITEM_FONDO_REPARO_USD, ITEM_FONDO_REPARO_ARS_IVA, ITEM_FONDO_REPARO_USD_IVA
			rtrn = PAGO_RECUPERO_FR
	end select 
	getTipoPago = rtrn
End Function
'--------------------------------------------------------------------------
'Funcion responsable por tomar un importe y calcular los pesos y dolares segun corresponda.
Function determinaImportes(pImporte, pMoneda, pTipoCambio, ByRef pImportePesos, ByRef pImporteDolares)

    pImportePesos = pImporte
	pImporteDolares = round(pImporte/pTipoCambio, 0)
	if (pMoneda = MONEDA_DOLAR) then
	    pImporteDolares = pImporte
	    pImportePesos = round(pImporte * pTipoCambio, 0)
	end if
	
End Function	
'--------------------------------------------------------------------------	
'CARGAR NUEVOS PAGOS DENTRO DE UN CONTRATO. -------------------------------
'LA PAGINA TRABAJA EN CONJUNTO CON OTRA A TRAVEZ DE AJAX, QUE CALCULA LOS -
'VALORES Y PORCENTAJES CORRESPONDIENTES. ----------------------------------
' LA DESCRIPCION DEL PAGO ES OBLIGATORIA ----------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'***********************************************
'*************  COMIENZO DE PAGINA  ************
'***********************************************
Dim myaux, controlOK, flagGrabar, flagSubmit, myArticulo, optStat, item , importePIC_old, member2Cd
Dim input_Importe, auxDsBudget, flagValorUnitario, obraFechaFin, obraFechaFinAjustada, rsBgt

Call GP_CONFIGURARMOMENTOS

accion = GF_PARAMETROS7("accion","",6)
CTC_idContrato = GF_PARAMETROS7("idContrato",0,6)
CTC_idPIC = GF_PARAMETROS7("idPIC",0,6)
CTC_cdMoneda = GF_PARAMETROS7("CTC_cdMoneda","",6)

Set rsCTC = readCTC(CTC_idContrato)

if (rsCTC.eof) then Response.Redirect "comprasAccesoDenegado.asp"

'Si quiere modificar el PIC, toma la obra del PIC sino busca la que esta activa.
if (CTC_idPIC = 0) then
	Set rsBgt = leerBudgetActivosCTC(CTC_idContrato)
	if (not rsBgt.eof) then CTC_idObra = rsBgt("IDOBRA")
else	
	Call readCTZ(CTC_idPIC)
	CTC_idObra = ctz_idObra
end if

CTC_fechaVto = rsCTC("FECHAVTO")
CTC_idDivision = rsCTC("IDDIVISION")
CTC_cdResponsable = rsCTC("CDRESPONSABLE")
if (CTC_idObra <> OBRA_GEID) then 
    Call loadDatosObra(CTC_idObra, CTC_obraCD, CTC_obraDS, "", "", 0, "", 0, "", obraFechaFin, obraFechaFinAjustada, "", "")    
    'Tomo la fecha de vto del contrato y la de vencimiento de la partida.
    'Me quedo con la menor de ambas para compararla con la fecha de autorizacion del CEC, ningun CEC se puede autorizar
    'luego de finalizado un contrato
    if (CLng(obraFechaFinAjustada) > 0) then obraFechaFin = obraFechaFinAjustada
    if (CLng(CTC_fechaVto) > CLng(obraFechaFin)) then CTC_fechaVto = obraFechaFin    
else
    CTC_obraCD = OBRA_GECD     
    CTC_obraDS = OBRA_GEDS    
end if
if (CTC_cdMoneda = "") then CTC_cdMoneda = rsCTC("CDMONEDA")
	
CTC_estado = CInt(rsCTC("ESTADO"))
CTC_tipo = rsCTC("TIPO")
flagValorUnitario = tieneValorUnitario(CTC_tipo)
if (isFormSubmit()) then
	flagSubmit = true
	myaux = GF_PARAMETROS7("AreaDetalle","",6)
	if (myaux = "") then myaux = "0,0"
	'compruebo los filtros area y detalle.	
	if (myaux <> "" and myaux <> ",") then	    
		myaux  = split(myaux,",")		
		CTC_areaObra = CInt(myaux(0))
		if (CTC_areaObra <> BUDGET_SIN_AREA) then 
			CTC_detalleObra = CInt(myaux(1))
		else
			CTC_detalleObra = BUDGET_SIN_DETALLES
		end if
	end if	
	CTC_idPedido = rsCTC("IDPEDIDO")
	CTC_cdContrato = rsCTC("CDCONTRATO")
	CTC_tipoCambio =  CDbl(GF_PARAMETROS7("CTC_tipoCambio", 3,6))
	'Obtengo los limites disponibles de pago
	CTC_ContratoPesos=0
	CTC_ContratoDolares = 0
	strSQL = "Select * FROM TBLCTCPARTIDAS WHERE IDCONTRATO = "& CTC_idContrato &" AND IDOBRA = "& CTC_idObra &_
			 " AND IDAREA = "&CTC_areaObra& " AND IDDETALLE= " & CTC_detalleObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rsImporte, "OPEN", strSQL)
	if (not rsImporte.eof) then Call determinaImportes(CDbl(rsImporte("IMPORTEASIGNADO")), rsImporte("CDMONEDA"), CTC_tipoCambio, CTC_ContratoPesos, CTC_ContratoDolares)    
	CTC_Total_ImporteSaldo = rsCTC("SALDO")
	CTC_idProveedor = rsCTC("IDPROVEEDOR")			
	'se toman los paramtros desde la pagina
	CTC_FechaEntrega = GF_PARAMETROS7("CTC_FechaEntrega","",6)
	if (CTC_FechaEntrega = "") then 
	    CTC_FechaEntrega = left(session("MmtoSistema"), 8)
    else
        CTC_FechaEntrega = GF_DTE2FN(CTC_FechaEntrega)
    end if        	    	
	CTC_tipoPago = GF_PARAMETROS7("CTC_tipoPago",0,6)
	CTC_Importe = GF_PARAMETROS7("CTC_Importe",0,6)	
	Call determinaImportes(CTC_Importe, CTC_cdMoneda, CTC_tipoCambio, CTC_ImportePesos, CTC_ImporteDolares)
	CTC_ImporteObra = GF_PARAMETROS7("CTC_ImporteObra",0,6)	
	Call determinaImportes(CTC_ImporteObra, CTC_cdMoneda, CTC_tipoCambio, CTC_ImporteObraPesos, CTC_ImporteObraDolares)
	CTC_Anticipo = GF_PARAMETROS7("CTC_Anticipo",0,6)
	Call determinaImportes(CTC_Anticipo, CTC_cdMoneda, CTC_tipoCambio, CTC_AnticipoPesos, CTC_AnticipoDolares)
	CTC_FReparo = GF_PARAMETROS7("CTC_FReparo",0,6)	
	Call determinaImportes(CTC_FReparo, CTC_cdMoneda, CTC_tipoCambio, CTC_FReparoPesos, CTC_FReparoDolares)
	'El importe total de los PICs surge de la suma de sus items.
	if (CTC_tipoPago = PAGO_OBRA) then
	    CTC_APagarPesos = CTC_ImportePesos + CTC_AnticipoPesos + CTC_FReparoPesos	
	    CTC_APagarDolares = CTC_ImporteDolares + CTC_AnticipoDolares + CTC_FReparoDolares	    
	else
	    CTC_APagarPesos = CTC_AnticipoPesos + CTC_FReparoPesos	
	    CTC_APagarDolares = CTC_AnticipoDolares + CTC_FReparoDolares
	end if
	'CTC_APagar = GF_PARAMETROS7("CTC_APagar",0,6)	
	'Call determinaImportes(CTC_APagar, CTC_cdMoneda, CTC_tipoCambio, CTC_APagarPesos, CTC_APagarDolares)	
	CTC_descripcion = GF_PARAMETROS7("CTC_descripcion","",6)
	CTC_aplicaAnticipo = GF_PARAMETROS7("CTC_aplicaAnticipo",0,6)
	CTC_aplicaFReparo = GF_PARAMETROS7("CTC_aplicaFReparo",0,6)
	CTC_ItemCantidad = GF_PARAMETROS7("CTC_ItemCantidad",0,6)
	if (CTC_ItemCantidad = 0) then CTC_ItemCantidad = PAGOS_ITEM_CANTIDAD
	CTC_valorUnitario = GF_PARAMETROS7("CTC_valorUnitario",0,6) 	
	Call determinaImportes(CTC_valorUnitario, CTC_cdMoneda, CTC_tipoCambio, CTC_valorUnitarioPesos, CTC_valorUnitarioDolares)
    importePIC_old = GF_PARAMETROS7("importePIC_old",0,6) 'ESTA VARIABLE SE UTILIZA PARA GUARDAR EL IMpORTE ANTERIOR AL MOMENTO DE SER EDITADO
	input_Importe = CTC_Importe/100 
	member2Cd = GF_PARAMETROS7("member2Cd", "",6)
	'se controlan los datos del pago
	controlOK = controlPago()	
	if ((accion = ACCION_GRABAR) and (controlOK)) then	
		if (CTC_APagarDolares <> 0) then
			'------------ CREA O ACTUALIZA EL PIC EN LA CTZCABECERA, GUARDA LAS FIRMAS Y EN CASO DE QUE   ----------------
			'------------ SE ENCUENTRE EN LA CTZDETALLE SE BORRA LOS ITEMS DEL PIC						  -----------------
			falgNewPIC = false 'variable que vamos a utilizar para saber si es un nuevo pic o una edicion de pic
            if (CTC_idPIC = 0) then falgNewPIC = true
            'se crea el PIC y se guardan las firmas
            Call addCTZCabecera(CTC_idPIC, CTC_idObra, CTC_idPedido, CTC_idProveedor, CTC_FechaEntrega, CTC_descripcion, CTC_APagarPesos, CTC_APagarDolares, CTC_tipoCambio, CTC_idDivision, CTC_cdMoneda, CTC_idContrato)
                        

			Call delCTZItems(CTC_idPIC)												
			'-------------PAGO OBRA:  SE GUARDA O ACTUALIZA EN CTZDETALLE EL PIC	  ---------------
			'se guarda el item siempre que no sea 0 y que sea un pago normal
			if ((CTC_ImporteDolares <> 0) and (CTC_tipoPago = PAGO_OBRA)) then
				Call getUnidadArticulo(ITEM_OBRAS_EN_CURSO, CTC_ItemUnidad, "", "")				
				Call addCTZItems(CTC_idPIC, getItemObra(rsCTC("TIPO")), CTC_ItemCantidad, CTC_ItemUnidad, CTC_areaObra, CTC_detalleObra, CTC_ImportePesos, CTC_ImporteDolares, CTC_tipoCambio)
			end if			
			'se guarda el item siempre que no sea 0, mismo con F. reparo
			if (CTC_AnticipoDolares <> 0) then
				Call getUnidadArticulo(ITEM_ANTICIPO_OBRAS_EN_CURSO, CTC_ItemUnidad, "", "")
				Call addCTZItems(CTC_idPIC, ITEM_ANTICIPO_OBRAS_EN_CURSO, CTC_ItemCantidad, CTC_ItemUnidad, CTC_areaObra, CTC_detalleObra, CTC_AnticipoPesos, CTC_AnticipoDolares, CTC_tipoCambio)
			end if
			
			'-------------PAGO FONDO DE REPARO: SE GUARDA O ACTUALIZA EN CTZDETALLE EL PIC   ---------------
			if (CTC_FReparoDolares <> 0) then				
				if (rsCTC("CDMONEDA") = MONEDA_PESO) then 
				    myArticulo = ITEM_FONDO_REPARO_ARS										
				else
					myArticulo = ITEM_FONDO_REPARO_USD
				end if				
				Call getUnidadArticulo(myArticulo, CTC_ItemUnidad, "", "")
			    Call addCTZItems(CTC_idPIC, myArticulo, CTC_ItemCantidad, CTC_ItemUnidad, CTC_areaObra, CTC_detalleObra, CTC_FReparoPesos, CTC_FReparoDolares, CTC_tipoCambio)			    
			end if
			'------------- Se graban las firmas
			Call addCTZFirmas(CTC_idPIC, CTC_cdResponsable, member2Cd)
			'------------- Se actuaiza el Saldo del Cto.
			Call ajusteSaldoPendiente(CTC_idContrato)
            'Call actualizarSaldoPendiente(CTC_APagarPesos,CTC_APagarDolares,CTC_cdMoneda,CTC_idContrato, CTC_idObra, CTC_areaObra, CTC_detalleObra,importePIC_old,falgNewPIC)
			flagGrabar = true
		end if
	end if
else	
	flagSubmit = false	
	CTC_areaObra = BUDGET_SIN_AREA
	CTC_detalleObra = BUDGET_SIN_DETALLES
	CTC_cdContrato = rsCTC("CDCONTRATO")	
	CTC_valorUnitario = rsCTC("IMPORTEUNITARIOPESOS")	
	if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_valorUnitario = rsCTC("IMPORTEUNITARIODOLARES")
	if (CTC_idPIC = 0) then		
		CTC_tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
		CTC_tipoPago = PAGO_OBRA
		CTC_FechaEntrega = left(session("MmtoSistema"), 8)
		'Si el contrato ya finalizo solo se puede devolver el fondo de reparo.
		'if (CTC_estado = ESTADO_CTC_FINALIZADO) then CTC_tipoPago = PAGO_RECUPERO_FR
		CTC_aplicaAnticipo = 0
		CTC_aplicaFReparo = 0
		CTC_ItemCantidad = 0				
		if (not flagValorUnitario) then
		    CTC_ItemCantidad = PAGOS_ITEM_CANTIDAD
		    CTC_aplicaAnticipo = 1
		    CTC_aplicaFReparo = 1
		end if				
	else				
		Set rsPago = readCTCPago(CTC_idPIC)		
		
		'Se esta modificando un pago.
		CTC_idObra = rsPago("IDOBRA")
		CTC_areaObra = rsPago("IDAREA")
		CTC_detalleObra = rsPago("IDDETALLE")
		CTC_tipoPago = getTipoPago(rsPago("IDARTICULO"))
		CTC_tipoCambio = rsPago("TIPOCAMBIO")
		CTC_ItemCantidad = rsPago("CANTIDAD")
		CTC_FechaEntrega = rsPago("FECHAENTREGA")
		CTC_APagar = CDbl(rsPago("IMPORTEPESOS"))
		if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_APagar = CDbl(rsPago("IMPORTEDOLARES"))						
		importePIC_old = CTC_APagar
		CTC_descripcion = rsPago("OBSERVACIONES")
		'Tomo el autorizante del CEC.		
	    strSqlFirmas = "SELECT CDUSUARIO FROM TBLCTZFIRMAS WHERE IDCOTIZACION = " & CTC_idPIC & " and SECUENCIA = " & PIC_FIRMA_GTE_SECTOR 
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSqlFirmas)
		if (not rs.eof) then member2Cd=rs("CDUSUARIO")			                    
            
		idArea = CLng(rsPago("IDAREA"))
		idArea_old = idArea
		idDetalle = CLng(rsPago("IDDetalle"))
		idDetalle_old = idDetalle
		While ((not rsPago.eof) and (idArea = idArea_old) and (idDetalle = idDetalle_old))		
			Select case rsPago("IDARTICULO")																
				Case ITEM_ANTICIPO_OBRAS_EN_CURSO
					CTC_Anticipo =  cdbl(rsPago("IMPORTEPESOS"))
					if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_Anticipo =   cdbl(rsPago("IMPORTEDOLARES"))					
				Case ITEM_FONDO_REPARO_ARS
					CTC_FReparo =  cdbl(rsPago("IMPORTEPESOS"))
				Case ITEM_FONDO_REPARO_USD
					CTC_FReparo = cdbl(rsPago("IMPORTEDOLARES"))
				Case ITEM_FONDO_REPARO_ARS_IVA	
					CTC_FReparo =  cdbl(rsPago("IMPORTEPESOS"))
				Case ITEM_FONDO_REPARO_USD_IVA	
					CTC_FReparo =  cdbl(rsPago("IMPORTEDOLARES"))
				Case else 'Cualquiera de los tipos de pago de obra
					CTC_Importe = cdbl(rsPago("IMPORTEPESOS"))					
					if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_Importe =  cdbl(rsPago("IMPORTEDOLARES"))					
				End Select
			rsPago.MoveNext()
		wend
		CTC_aplicaAnticipo = 0
		if (CTC_AnticipoPesos <> 0) then	CTC_aplicaAnticipo = 1
			
		CTC_aplicaFReparo = 0
		if (CTC_FReparoPesos <> 0) then CTC_aplicaFReparo = 1
		
		if (CTC_Importe = 0) then
			'NO ES UN PAGO DE OBRA 
			if(CTC_Anticipo <> 0)then 
				'ES UN ANTICIPO 
				CTC_Importe = CTC_Importe + CTC_Anticipo
			end if
			if (CTC_FReparo <> 0) then 
				'ES UN F.REPARO
				CTC_Importe = CTC_Importe + CTC_FReparo
			end if
		end if			
		input_Importe = CTC_Importe / 100
	end if
end if

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title><% =GF_TRADUCIR("Sistema de Compras - Contratos") %></title>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" type="text/css" href="css/calendar-win2k-2.css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
	<script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/formato.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
	<script type="text/javascript" src="scripts/controles.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
  	<script type="text/javascript" src="scripts/calendar-<% =session("UsuarioIdiomaCodigo") %>.js"></script>
	<script type="text/javascript">
		var chImp = new channel();
		var aplicaAnticipo;
		var aplicaFReparo;
		var tipoPago;
		var refpopupPago;
        var ctrlSTOPSubmit = 0;
        
		function bodyOnLoad(){
			refpopupPago = getObjPopUp('popUpPago');
			<% if (flagGrabar) then %>
				cerrar();
			<% end if%>
			var	tb = new Toolbar('toolbar', 4, "images/compras/");
			idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
			idBtnControl = tb.addButtonCONFIRM("Controlar", "submitInfo('<% =ACCION_CONTROLAR %>')");			
			tb.addButton("Close-16x16.png", "Cerrar", "cerrar()");
			tb.draw();
			aplicaAnticipo = <% =CTC_aplicaAnticipo %>;
			aplicaFReparo = <% =CTC_aplicaFReparo %>;
			actualizarPago();
		}

        function actualizarPago(){
            <%  if (flagValorUnitario) then %>
			calcularValorUnidades();
			<%  else %>			
			calcularMoneda();
			<%  end if %>			
			calcularImportes();
        }
        
        function seleccionAutorizante() {
	        if (document.getElementById("cmbUsrAut")) {
		        var e = document.getElementById("cmbUsrAut");
                document.getElementById("member2Cd").value = e.options[e.selectedIndex].value;
            }            
	    }
	    
		function submitInfo(acc){
		    if (ctrlSTOPSubmit == 0) {
			    document.getElementById("accion").value = acc;
			    document.getElementById("frmSel").submit();
			}
		}

		function cerrar() {
			parent.recargar();
			refpopupPago.hide();
		}
		
		function calcularValorUnidades() {
		    var uImporte = document.getElementById("CTC_valorUnitario").value;		    
		    var uCantidad = document.getElementById("CTC_ItemCantidad").value;
		    var hiddenP = document.getElementById("CTC_Importe");						
            hiddenP.value = uCantidad*uImporte;
		}
		
		function calcularMoneda() {						
		    var hiddenP = document.getElementById("CTC_Importe");			
		    var objP = document.getElementById("input_Importe");
			var objPValue = objP.value.replace(/,/,".");
			hiddenP.value = objP.value * 100;		    		    			
		}

        function tomarMoneda() {
            var selMoneda = document.getElementById("CTC_cdMonedaSel");
			document.getElementById("CTC_cdMoneda").value = selMoneda.options[selMoneda.selectedIndex].value;
        }
        
		function calcularImportes() {
			tipoPago = getTipoPago();			
			var importe = document.getElementById("CTC_Importe").value;			
			var cdMoneda = document.getElementById("CTC_cdMoneda").value;
			document.getElementById("importes").innerHTML="<table align='center'><tr><td><img src='images/compras/loading_big.gif'></td></tr></table>";
			var params = '?idContrato=<% =CTC_idContrato %>';
			params = params + '&CTC_tipoPago=' + tipoPago + '&CTC_Importe=' + importe + '&CTC_cdMoneda=' + cdMoneda;
			params = params + '&CTC_aplicaAnticipo=' + aplicaAnticipo + '&CTC_aplicaFReparo=' + aplicaFReparo;
			params = params + '&CTC_idPIC=<%=CTC_idPIC%>';
			chImp.bind('comprasCTCPagoAjax.asp' + params,'calcularImportes_callback()');
			ctrlSTOPSubmit = 1;
			chImp.send();			
		}
	
		function calcularImportes_callback() {
		    ctrlSTOPSubmit = 0;
			var resp = chImp.response();
			document.getElementById("importes").innerHTML=resp;
			if (tipoPago != <% =PAGO_OBRA %>) {
				document.getElementById("CTC_aplicaAnticipo").disabled = true;
				document.getElementById("CTC_aplicaFReparo").disabled = true;
			}
		}
        		
		function getTipoPago() {
			var i
			for (i=0;i < document.frmSel.CTC_tipoPago.length;i++){
                if (document.frmSel.CTC_tipoPago[i].checked) break;
			}
			if (i < document.frmSel.CTC_tipoPago.length) {
			    return document.frmSel.CTC_tipoPago[i].value;
			} else {
			    return <% =PAGO_OBRA %>;
			}
		}

		function aplicarAnticipo() {
			if (aplicaAnticipo == 0) {
			//if (document.frmSel.CTC_aplicaAnticipo.checked) {
				aplicaAnticipo = 1
			} else {
				aplicaAnticipo = 0
			}
			calcularImportes();
		}

		function aplicarFReparo() {
			if (document.frmSel.CTC_aplicaFReparo.checked) {
				aplicaFReparo = 1			
				refpopupPago.resize(530, 590);
			} else {
				aplicaFReparo = 0				
				refpopupPago.resize(530, 570);
			}
			calcularImportes();
		}	
		function CerrarCal(cal) {
			cal.hide();
		}		
		function MostrarCalendario(p_objID, funcSel) {
			var dte= new Date();		    	    
			var elem= document.getElementById(p_objID);
			if (calendar != null) calendar.hide();		
			var cal = new Calendar(false, dte, funcSel, CerrarCal);
			cal.weekNumbers = false;
			cal.setRange(1993, 2045);
			cal.create();
			calendar = cal;		
			calendar.setDateFormat("dd/mm/y");
			calendar.showAtElement(elem);
		}
		function SeleccionarCal(cal, date) {
			var str= new String(date);		
			document.getElementById("CTC_FechaEntrega").value = str;
			if (cal) cal.hide();
		}		
	</script>
</head>
<body onload="bodyOnLoad()">
	<div id="toolbar"></div>
	<br>
	<form method="post" id="frmSel" name="frmSel">
		<table class="reg_header" width="95%" align="center" border="0">
			<tr><td><% call showErrors() %></td></tr>
			<tr>
				<td align="right" style="font-weight: bold;font-size: 14px;">
					<% =GF_TRADUCIR("Id Contrato:") %>&nbsp;<% =CTC_cdContrato %>
				</td>
				<input type="hidden" name="CTC_idContrato" id="CTC_idContrato" value="<%=CTC_idContrato%>">
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Tipo de Pago") %></td>
			</tr>
			<tr>
				<td>
					<table width="100%" border="0">
						<tr>
							<td>
							    <%  optStat = ""							    
							        if (CTC_tipoPago = PAGO_OBRA) then optStat = "checked"
							    %>
								<input type="radio" id="CTC_tipoPago" name="CTC_tipoPago" value="<% =PAGO_OBRA %>" onChange="calcularImportes();" <% =optStat %>>
								<% 
								    if (CTC_tipo <> CTC_TIPO_OBRA) then 
								        response.Write GF_TRADUCIR("Servicio") 
								    else
								        response.Write GF_TRADUCIR("Obra")
								    end if								
								%>
							</td>
							<td>								
							</td>
						</tr>
						<tr>
							<td>
							    <%  optStat = ""
							        if (CTC_tipoPago = PAGO_ANTICIPO) then optStat = "checked" 							    
							    %>
								<input type="radio" id="CTC_tipoPago" name="CTC_tipoPago" value="<% =PAGO_ANTICIPO %>" onChange="calcularImportes();" <% =optStat %>>
								<% =GF_TRADUCIR("Anticipo") %>
							</td>
							<td>
							    <%  optStat = ""
							        if (CTC_tipoPago = PAGO_RECUPERO_FR) then optStat = "checked" 							        
							    %>
								<input type="radio" id="CTC_tipoPago" name="CTC_tipoPago" value="<% =PAGO_RECUPERO_FR %>" onChange="calcularImportes();" <% =optStat %>>
								<% =GF_TRADUCIR("Recupero Fondo de Reparo") %>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Partida Presupuestaria") %></td>
			</tr>
			<tr>
				<td>
				    &nbsp;<% =CTC_obraCD & "-"  %>
        <%  if (CTC_idObra = OBRA_GEID) then 
                response.Write CTC_obraDS
            else 
                if (CTC_idPIC = 0) then		
                    idAreaOLD = 0    
            %>				
				    <select name="AreaDetalle" id="AreaDetalle" size="1">		
				        <option value="<%=BUDGET_SIN_AREA%>"><%=GF_TRADUCIR("Ninguno...")%></option>
    <%                  while not rsBgt.eof 
						    myValue = rsBgt("IDAREA") & "," & rsBgt("IDDETALLE")						
						    if (CLng(rsBgt("IDAREA")) <> CLng(idAreaOLD)) then      %>
						        <optgroup label="<%	=rsBgt("IDAREA") & " - " & rsBgt("DSCABECERA") %>">    
    <%                      end if
                            mySelect = ""
						    if (cint(CTC_areaObra)=rsBgt("IDAREA")) and (cint(CTC_detalleObra) = rsBgt("IDDETALLE")) then mySelect = "selected='selected'"
    %>
						        <option value="<%=myValue %>" <%=myClass%> <%=mySelect%>><% ="&nbsp;&nbsp;&nbsp;&nbsp;" & rsBgt("IDDETALLE") & " - " &  rsBgt("DSDETALLE") %></option>
    <%                      if (CLng(rsBgt("IDAREA")) <> CLng(idAreaOLD)) then      %>
						        </optgroup>
    <%						    idAreaOLD = rsBgt("IDAREA")
                            end if						                                    			        
    						rsBgt.movenext
					    wend	%>
				    </select>		
		<%	    else
		            auxDsBudget = "&nbsp;" & CTC_areaObra & "-" & CTC_detalleObra & "&nbsp;-"
					auxDsBudget = auxDsBudget & "&nbsp;" & getDsBudget(CTC_idObra, CTC_areaObra, CTC_detalleObra)
					Response.Write auxDsBudget
		%>			
		            <input type="hidden" name="AreaDetalle" id="Hidden1" value="<% =CTC_areaObra & "," & CTC_detalleObra %>">
        <%      end if
		    end if
		%>
				</td>
			</tr>
        <%  if (CTC_idObra = OBRA_GEID) then    %>
            <tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Division") %></td>
			</tr>
            <tr>
				<td>
		            <%= getDescripcionDivision(CTC_idDivision) %>
				</td>
			</tr>            
        <%  end if        %>
            <tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Fecha de Pasaje a Agenda de Autorizaciones") %></td>				
			</tr>			
			<tr>
			    <td><input type="text" name="CTC_FechaEntrega" id="CTC_FechaEntrega" size="15" readonly onClick="javascript:MostrarCalendario('CTC_FechaEntrega', SeleccionarCal)" value="<% =GF_FN2DTE(CTC_FechaEntrega) %>"></td>
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Pago a Realizar") %></td>
			</tr>			
			<tr>				
				<td>
					<table width="100%" border="0">
						<tr>
							<td>
							<% if (flagValorUnitario) then %>
							    Cantidad Unidades
							<% else %>
							    Importe Factura
							<% end if %>
							</td>							
							<% if (not flagValorUnitario) then %>
							<td>							
							    <select id="CTC_cdMonedaSel" name="CTC_cdMonedaSel" onchange="tomarMoneda()">
					                <option value="<% =MONEDA_DOLAR %>" <% if(CTC_cdMoneda = MONEDA_DOLAR) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_DOLAR) %>
					                <option value="<% =MONEDA_PESO %>"  <% if(CTC_cdMoneda = MONEDA_PESO) then response.write "selected='true'" %>><% =getSimboloMoneda(MONEDA_PESO) %>							
				                </select>								
							</td>
							<td>			
							    <input type="text" id="input_Importe" name="input_Importe" value="<% =input_Importe %>" style="text-align:right;" onKeyPress="return controlIngreso(this, event, 'I')" onBlur="actualizarPago()">							
                            </td>							    		
                            <td>T.C.</td>
							<td><input type="input" name="CTC_tipoCambio" id="CTC_tipoCambio" size="5" value="<% =CTC_tipoCambio %>" onBlur="actualizarPago()"></td>
							<% else %>				
							<td colspan="4">			
							    <input type="text" id="CTC_ItemCantidad" name="CTC_ItemCantidad" value="<% =CTC_ItemCantidad %>" size="5" style="text-align:right;" onKeyPress="return controlIngreso(this, event, 'I')" onBlur="actualizarPago()">
							    <input type="hidden" id="CTC_valorUnitario" name="CTC_valorUnitario" value="<% =CTC_valorUnitario %>">							    							    
							    <input type="hidden" name="CTC_tipoCambio" id="CTC_tipoCambio" value="<% =CTC_tipoCambio %>">
							</td>
							<% end if %>							
						</tr>
					</table>
					<input type="hidden" id="CTC_cdMoneda" name="CTC_cdMoneda" value="<% =CTC_cdMoneda %>">							    
					<input type="hidden" id="CTC_Importe" name="CTC_Importe" value="<% =CTC_Importe %>">
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Resumen") %></td>
			</tr>
			<tr>
				<td><div id="importes"></div><br/></td>
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Descripción") %></td>
			</tr>
			<tr>
				<td align="center">
					<textarea id="CTC_descripcion" name="CTC_descripcion" value="<% =CTC_descripcion %>" maxlength="3500" cols="50" rows="5"><% =CTC_descripcion %></textarea>
				</td>
			</tr>
			<tr>
				<td class="reg_header_nav"><% =GF_TRADUCIR("Firmantes") %></td>
			</tr>
			<tr>
				<td align="center">
				    <table border="0" cellpadding="0" cellspacing="0" width="100%">
				        <tr>
				            <td width="50%" align="center">Responsable</td>
				            <td width="50%" align="center">Autorizante</td>
				        </tr>
				        <tr>
				            <td width="50%" align="center"><b><% =getUserDescription(CTC_cdResponsable) %></b></td>
				            <td width="50%" align="center"><% Call dibujarComboGte(CTC_cdResponsable, member2Cd) %></td>
				        </tr>
				    </table>					
				</td>
			</tr>
		</table>
		<input type="hidden" id="accion" name="accion" value="">		
		<input type="hidden" id="idPIC" name="idPIC" value="<% =CTC_idPIC %>">
        <input type="hidden" id="importePIC_old" name="importePIC_old" value="<% =importePIC_old%>">
        <input type="hidden" id="member2Cd" name="member2Cd" value="<%=member2Cd%>">
	</form>
</body>
</html>