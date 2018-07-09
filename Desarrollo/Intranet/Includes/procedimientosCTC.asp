<%
'Fondo de reparo sugerido en %
Const F_REPARO_INICIAL = 5

'tipos de contratos
Const CTC_TIPO_OBRA		  = "O" 'Pago Contra Certificados de Obra
Const CTC_TIPO_CUOTA_FIJA = "C" 'Servicio Repetitivo - Cuotas Fijas 
Const CTC_TIPO_UNITARIO   = "U" 'Servicio Repetitivo - Valor Unidad Fijo
Const CTC_TIPO_GENERAL	  = "G" 'Contrato General

Const CTC_AJUSTE_UNITARIO   = "U" 'Ajuste del vaor de unidad del contrato Solo para (Servicio Repetitivo - Valor Unidad Fijo)
Const CTC_AJUSTE_GENERAL    = "G" 'Ajuste del presupuesto total del contrato. 
 
'Tipos de pago
Const PAGO_OBRA			= 1
Const PAGO_ANTICIPO		= 2
Const PAGO_RECUPERO_FR	= 3

'Tipo items
Const ITEM_OBRAS_EN_CURSO			= 8835  'Se utilizara para pagos de obra
Const ITEM_SERVICIOS_GENERALES		= 11548 'Se utilizara para pagos de Servicios Generales
Const ITEM_ANTICIPO_OBRAS_EN_CURSO	= 10536 'Se utilizara para pagos de anticipo (+) y aplicacion de anticipo a pagos de obra (-)
Const ITEM_FONDO_REPARO_ARS			= 9664  'Se utilizará para retención (-) y devolucion (+) de Fondo de reparo en Pesos
Const ITEM_FONDO_REPARO_USD			= 10650 'Se utilizará para retención (-) y devolucion (+) de Fondo de reparo en Dolares
Const ITEM_FONDO_REPARO_ARS_IVA		= 11321  'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano por personal de IT en el PIC. Se utilizará para devolucion (+) de Fondo de reparo en Pesos.
Const ITEM_FONDO_REPARO_USD_IVA		= 11322 'Item especial, solamente se usa en contratos viejos cuando el fondo de reparo incluia IVA. Siempre que se usa se reemplazo a mano por personal de IT en el PIC. Se utilizará para devolucion (+) de Fondo de reparo en Dolares

'Estado contratos
Const ESTADO_CTC_PENDIENTE	=10	'contrato cargado en el sistema
Const ESTADO_CTC_AUTORIZADO	=20	'contrato autorizado por legales
Const ESTADO_CTC_FINALIZADO	=30	'se completaron todos los pagos del contrato
Const ESTADO_CTC_CANCELADO	=40	'contrato cancelado
Const ESTADO_CTC_EN_AJUSTE	=50	'contrato en ajuste

'Firmantes de ajustes
'Const SECUENCIA_RESPONSABLE = 0
Const CTC_FIRMA_RESPONSABLE = 0
Const CTC_FIRMA_GTE_PUERTO  = 1
Const CTC_FIRMA_GTE_SECTOR  = 2
Const CTC_FIRMA_GTE_COMPRAS = 3

'cdContrato por defecto

Const CONTRATO_A_CONFIRMAR = "A CONFIRMAR" 'texto por defecto a la hora de crear el contrato, se coloca en la columna CDCONTRATO
Const CONTRATO_TIPO_SERVICIO = "SERVICIO" 'texto por defecto cuando se crea un contrato SIN CONTRATO FISICO, se coloca en la columna CDCONTRATO

Dim CTC_idContrato, CTC_cdMoneda, CTC_tipoPago
Dim CTC_APagarPesos, CTC_APagarDolares, CTC_AnticipoPesos, CTC_AnticipoDolares
Dim CTC_FReparoPesos, CTC_FReparoDolares, CTC_ImportePesos, CTC_ImporteDolares
Dim CTC_APagar, CTC_Anticipo, CTC_FReparo, CTC_Importe, CTC_ImporteObra
Dim CTC_ContratoPesos, CTC_ContratoDolares, CTC_ImporteAPje, CTC_aplicaFReparo
Dim CTC_PjeAnticipo, CTC_PjeFReparo, CTC_PjePago, CTC_aplicaAnticipo
Dim CTC_ImporteObraPesos, CTC_ImporteObraDolares, CTC_ImportePesosFacturado, CTC_ImporteDolaresFacturado
Dim CTC_idObra, CTC_idPedido, CTC_idProveedor, CTC_tipoCambio, CTC_observaciones
Dim CTC_cdResponsable, CTC_idPIC, CTC_Titulo
Dim CTC_obraCD, CTC_obraDS, CTC_detalleObra, CTC_areaObra, CTC_descripcion
Dim CTC_idDivision, CTC_FechaEntrega, CTC_ItemCantidad, CTC_ItemUnidad, CTC_TotalImporte
Dim CTC_cdObra, CTC_dsObra, CTC_cdContrato, CTC_dsProveedor, CTC_dsResponsable, CTC_estado, CTC_fechaVto
Dim CTC_valorUnitarioPesos, CTC_valorUnitarioDolares, CTC_valorUnitario
Dim CTC_PIC_ImporteObra, CTC_PIC_ImporteAnticipo, CTC_PIC_ImporteFReparo
Dim CTC_tipo, CTC_Total_ImporteSaldo
'--------------------------------------------------------------------------------------------
' Función:	grabarCTC
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	05/01/2011
' Modifico: JAS - Javier A. Scalisi
' Fecha:	05/09/2011
' Objetivo:	
'			Grabar contrato en la base de datos
' Parametros:
'			pIdContrato 	[int] 	ID CONTRATO
'			pIdObra 		[int] 	ID obra
'			pIdPedido		[int]	ID pedido
'			pIdProveedor	[int] 	ID proveedor asignado
'			pImportePesos 	[int] 	Importe total del contrato en pesos
'			pImporteDolar	[int]	Idem en dolares
'			pImporteUPesos 	[int] 	Importe unitario del contrato en pesos
'			pImporteUDolar	[int]	Idem en dolares
'			pFondoReparo	[int] 	Fondo de reparo que el contrato desigana
'			pCdResponsable 	[char] 	responsable del contrato
'			pIdArea			[int] 	ID area
'			pIdDetalle		[int] 	ID detalle
'			pCdContrato		[char]	Numero Contrato
'			pCdMoneda		[char]	Codigo Moneda
'			pFechaVto		[string]	Fecha Vencimiento
'			pTipo			[char]	Tipo Contrato
' Devuelve:
'			valor booleano
'--------------------------------------------------------------------------------------------
Function grabarCTC(ByRef pIdContrato, pTitulo, pIdPedido, pIdProveedor, pImportePesos, pImporteDolar, pImporteUPesos, pImporteUDolar, pFondoReparo, pCdResponsable, pCdContrato, pCdMoneda, pFechaVto, pTipo, pIdDivision,pEstado)
	Dim esta, strSQL, conn, rs, connIns, rsIns, auxSaldo,rsPCT
	grabarCTC = false

	if (pIdContrato <> 0) then
		Set rs = readCTC(pIdContrato)
		if (not rs.eof) then esta = true
	end if
    auxSaldo = pImportePesos 'VARIABLE QUE SE USA PARA CARGAR COLUMNA SALDO SEGUN MONEDA DEL CONTRATO
    if (pCdMoneda = MONEDA_DOLAR) then auxSaldo = pImporteDolar
	if (not esta) then
		strSQL = "Insert Into TBLOBRACONTRATOS(IDPEDIDO,IDPROVEEDOR,IMPORTEPESOS,IMPORTEDOLARES,FONDOREPARO,CDRESPONSABLE,ESTADO,CDCONTRATO, CDMONEDA, FECHAVTO, TIPO, IMPORTEUNITARIOPESOS, IMPORTEUNITARIODOLARES, TITULO,IDDIVISION,SALDO) VALUES(" & pIdPedido & "," & pIdProveedor & "," & pImportePesos & "," & pImporteDolar & "," & pFondoReparo & ",'" & pCdResponsable & "'," & pEstado & ",'" & pCdContrato & "', '"& pCdMoneda &"', " & pFechaVto & ", '" & pTipo &"', " & pImporteUPesos & ", " & pImporteUDolar & ", '" & pTitulo & "', "& pIdDivision &","& auxSaldo &")"
		Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
		
		strSQL = "Select Max(IDCONTRATO) IDCONTRATO from TBLOBRACONTRATOS where IDPEDIDO=" & pIdPedido
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			pIdContrato = rs("IDCONTRATO")
		end if				
				
        if (pEstado = ESTADO_CTC_PENDIENTE) then Call sendMailToLegales(pIdContrato, pTitulo, pIdPedido, pIdProveedor, pCdResponsable)
        if (pEstado = ESTADO_CTC_AUTORIZADO) then
            if pIdPedido > 0 then
			    strSQL = "Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_APROBADO & " WHERE IDPEDIDO = "& pIdPedido
                Call executeQueryDb(DBSITE_SQL_INTRA, rsPCT, "UPDATE", strSQL)
		    end if
        end if
	else
		strSQL = "Update TBLOBRACONTRATOS set IMPORTEPESOS=" & pImportePesos & ", IMPORTEDOLARES=" & pImporteDolar & ", IMPORTEUNITARIOPESOS=" & pImporteUPesos & ", IMPORTEUNITARIODOLARES=" & pImporteUDolar & ", FONDOREPARO=" & pFondoReparo & ", CDRESPONSABLE='" & pCdResponsable & "', ESTADO=" & ESTADO_CTC_PENDIENTE & ", TIPO = '" & pTipo & "', FECHAVTO = " & pFechaVto & ", IDDIVISION = "& pIdDivision &",SALDO = "& auxSaldo &" where IDCONTRATO = " & pIdContrato
		Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "UPDATE", strSQL)
	end if

	grabarCTC = true
End Function
'--------------------------------------------------------------------------------------------
' Función:	readCTC
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	06/01/11
' Objetivo:	
'			Leer contrato en la base de datos
' Parametros:
'			pIdContrato 	[int] 	ID CONTRATO
' Devuelve:
'			rs del contrato con todos sus campos
'--------------------------------------------------------------------------------------------
Function readCTC(pIdContrato)
	Dim strSQL, conn, rs, tipoCambio

    tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
    
	strSQL = "		    SELECT IDCONTRATO,     "
	strSQL = strSQL & "		   TITULO,	       "  	
	strSQL = strSQL & "		   IDPEDIDO,       " 
	strSQL = strSQL & "		   IDPROVEEDOR,    " 
	strSQL = strSQL & "        case when CDMONEDA='" & MONEDA_PESO & "' then IMPORTEPESOS else IMPORTEDOLARES * " & tipoCambio & " end	IMPORTEPESOS, " 
	strSQL = strSQL & "        case when CDMONEDA='" & MONEDA_PESO & "' then IMPORTEPESOS/" & tipoCambio & " else IMPORTEDOLARES end	IMPORTEDOLARES, "
	strSQL = strSQL & "		   IMPORTEDOLARES, " 
	strSQL = strSQL & "		   FONDOREPARO,    " 
	strSQL = strSQL & "		   CDRESPONSABLE,  " 
	strSQL = strSQL & "		   CDCONTRATO,     " 
	strSQL = strSQL & "	 	   ARCHIVO,        " 
	strSQL = strSQL & "	 	   ESTADO,         " 
	strSQL = strSQL & "		   ARCHIVO_EXT,    " 
	strSQL = strSQL & "	 	   CDMONEDA,       " 
	strSQL = strSQL & "		   MMTOCONF,       " 
	strSQL = strSQL & "		   CDUSERCONF,     " 
	strSQL = strSQL & "		   CASE WHEN TIPO IS NULL THEN '' ELSE TIPO END AS TIPO, "
	strSQL = strSQL & "		   CASE WHEN FECHAVTO IS NULL THEN 0 ELSE FECHAVTO END AS FECHAVTO, "
	strSQL = strSQL & "		   IMPORTEUNITARIOPESOS,   " 
	strSQL = strSQL & "		   IMPORTEUNITARIODOLARES, " 
    strSQL = strSQL & "		   CASE WHEN IDDIVISION IS NULL THEN 0 ELSE IDDIVISION  END AS IDDIVISION, " 
    strSQL = strSQL & "		   IDDIVISION,    " 
    strSQL = strSQL & "		   SALDO     " 
	strSQL = strSQL & "	FROM   TBLOBRACONTRATOS WHERE IDCONTRATO =  " & pIdContrato		
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set readCTC = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	readCTCPagos
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	12/01/11
' Objetivo:	
'			Lee todos los pagos de un contrato especifico.
' Parametros:
'			pIdContrato 	[int] 	ID CONTRATO
' Devuelve:
'			rs con la lista de pagos
'--------------------------------------------------------------------------------------------
Function readCTCPagos(pIdContrato)
	Dim strSQL, conn, rs, myImporte, tipoCambio
    
    strSQL = readCTCPagosSQL(pIdContrato) & " ORDER BY IDPIC, IDAREA, IDDETALLE, IDFAC, IDARTICULO"
   Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readCTCPagos = rs
End Function
'--------------------------------------------------------------------------------------------
Function readCTCPagosSQL(pIdContrato)
	Dim strSQL, tipoCambio

    tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
    
	strSQL=	"Select * from " & _
			"((SELECT CTZ.IDCOTIZACION  as IDPIC," & _
			" CTZ.IDOBRA, " &_
			" CASE when CTZ.IDOBRA = " & OBRA_GEID & " then '" & OBRA_GECD & "' else OBR.CDOBRA end CDOBRA, " & _
			" DET.IDAREA, " & _
			" DET.IDDETALLE, " & _
			" CTZ.estado," & _
			" CTZ.observaciones," & _
			" DET.IDARTICULO," & _						
			" case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then CTZ.IMPORTEPESOS else CTZ.IMPORTEDOLARES * " & tipoCambio & " end	IMPORTEPESOSPIC, " & _
			" case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then CTZ.IMPORTEPESOS/" & tipoCambio & " else CTZ.IMPORTEDOLARES end	IMPORTEDOLARESPIC, " & _
			" 0						AS IDFAC," & _
			" case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then DET.IMPORTEPESOS else DET.IMPORTEDOLARES * " & tipoCambio & " end	IMPORTEPESOS, " & _
            " case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then DET.IMPORTEPESOS/" & tipoCambio & " else DET.IMPORTEDOLARES end	IMPORTEDOLARES, " & _			
			" CTZ.FECHAENTREGA      AS FECHAAGENDA, " &_
	    	" 'CEC'					AS TIPOFAC " & _
			" FROM (Select * from tblctzcabecera where idcontrato = " & pIdContrato & " and estado <> '" & CTZ_ANULADA & "') CTZ  " & _ 			
			" INNER JOIN tblctzdetalle DET ON CTZ.idcotizacion = DET.idcotizacion " & _	
			" LEFT JOIN TBLDATOSOBRAS OBR on OBR.IDOBRA=CTZ.IDOBRA " &_
			" ) UNION (" & _
			"SELECT CTZ.IDCOTIZACION  as IDPIC," & _
			" CTZ.IDOBRA, " &_
			" CASE when CTZ.IDOBRA = " & OBRA_GEID & " then '" & OBRA_GECD & "' else OBR.CDOBRA end CDOBRA, " & _			
			" DET.IDAREA, " & _
			" DET.IDDETALLE, " & _
			" CTZ.estado," & _
			" CTZ.observaciones," & _
			" DET.IDARTICULO," & _
			" case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then CTZ.IMPORTEPESOS else CTZ.IMPORTEDOLARES * " & tipoCambio & " end	IMPORTEPESOSPIC, " & _
			" case when CTZ.CDMONEDA='" & MONEDA_PESO & "' then CTZ.IMPORTEPESOS/" & tipoCambio & " else CTZ.IMPORTEDOLARES end	IMPORTEDOLARESPIC, " & _
			" FAC.nroInt       AS IDFAC," & _
			" FAC.ImportePesos*100   AS IMPORTEPESOS, " & _
			" FAC.ImporteDolares*100   AS IMPORTEDOLARES, " & _
			" 0                AS FECHAAGENDA, " & _ 
			" CASE WHEN FACCAB.tipcbt = "& CBTE_PROVEEDORES_FAC &" then '" & PREFIX_FAC & "' when FACCAB.tipcbt = "& CBTE_PROVEEDORES_NDB &" then '" & PREFIX_NDB & "' else '" & PREFIX_NCR & "' end	AS TIPOFAC "  &_          
			" FROM (Select * from tblctzcabecera where idcontrato = " & pIdContrato & " and estado <> '" & CTZ_ANULADA & "') CTZ  " & _			
			" INNER JOIN tblctzdetalle DET ON CTZ.idcotizacion = DET.idcotizacion " & _
			" INNER JOIN (Select anio, mes, nroInt, IDPIC, IDArticulo, IDArea, IDDetalle, SUM(ImportePesos) ImportePesos, SUM(ImporteDolares) ImporteDolares from VWMEP001C group by anio, mes, nroInt, IDPIC, IDArticulo, IDArea, IDDetalle) FAC ON CTZ.idcotizacion = FAC.IDPIC and DET.IDARTICULO=FAC.IDArticulo and DET.IDAREA=FAC.IDArea and DET.IDDETALLE=FAC.IDDetalle " & _  			
			" INNER JOIN (SELECT tipcbt,nroint from [Database].[dbo].MEP001A where anulado <> 'S') FACCAB ON FACCAB.nroInt = FAC.nroInt " & _    
			" LEFT JOIN TBLDATOSOBRAS OBR on OBR.IDOBRA=CTZ.IDOBRA )) AS TABLA "
	'Response.Write strSQL	
	readCTCPagosSQL = strSQL
End Function
'--------------------------------------------------------------------------------------------
' Función:	readCTCPago
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	20/12/11
' Modificacion:	CNA - Ajaya Nahuel
' Fecha:	10/04/2013
' Objetivo:	
'			Lee un pago de un contrato especifico.
' Parametros:
'			pIdPago 	[int] 	ID PIC
' Devuelve:
'			rs con los datos del pago
'--------------------------------------------------------------------------------------------
Function readCTCPago(pIdPago)
	Dim strSQL, conn, rs
	strSQL = "		    SELECT DISTINCT CTZ.*,	"
	strSQL = strSQL & "		   DET.IDAREA,		"
	strSQL = strSQL & "		   DET.IDDETALLE,	"
	strSQL = strSQL & "		   CTC.IMPORTEPESOS AS PAGOPESOS,	    " 
	strSQL = strSQL & "		   CTC.IMPORTEDOLARES AS PAGODOLARES,   "  		
	strSQL = strSQL & "		   CTC.FONDOREPARO,					    "  		
	strSQL = strSQL & "		   DET.IDARTICULO,					    "  		
	strSQL = strSQL & "		   DET.CANTIDAD,					    "  		
	strSQL = strSQL & "		   DET.IMPORTEPESOS,				    "  		
	strSQL = strSQL & "		   DET.IMPORTEDOLARES				    "  		
	strSQL = strSQL & " FROM TBLCTZCABECERA CTZ		    " 
	strSQL = strSQL & "	  INNER JOIN (Select Distinct IDCOTIZACION, "
	strSQL = strSQL & "						 IDAREA,				"
	strSQL = strSQL & "				   	     IDDETALLE,				"
	strSQL = strSQL & "				   	     IDARTICULO,			"
	strSQL = strSQL & "				   	     CANTIDAD,			"
	strSQL = strSQL & "						 IMPORTEPESOS,		    "
	strSQL = strSQL & "						 IMPORTEDOLARES		    "			
	strSQL = strSQL & "				FROM TBLCTZDETALLE	"
	strSQL = strSQL & "				WHERE IDCOTIZACION = " & pIdPago & ") DET " 
	strSQL = strSQL & "			ON CTZ.IDCOTIZACION = DET.IDCOTIZACION		  " 
	strSQL = strSQL & " INNER JOIN tblobracontratos CTC on CTC.idcontrato = CTZ.idcontrato "
	strSQL = strSQL & "	WHERE CTZ.IDCOTIZACION = " & pIdPago
		
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readCTCPago = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	getPedidoCTC
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	11/01/11
' Objetivo:	
'			devolver el id del pedido de un contrato especifico.
' Parametros:
'			pIdContrato 	[int] 	ID CONTRATO
' Devuelve:
'			IDPEDIDO
'--------------------------------------------------------------------------------------------
Function getPedidoCTC(pIdContrato)
	Dim strSQL, conn, rs, rtrn
	rtrn = 0
	
	strSQL="Select IDPEDIDO from TBLOBRACONTRATOS where IDCONTRATO = " & pIdContrato
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then rtrn = cLng(rs("IDPEDIDO"))

	getPedidoCTC = rtrn
End Function
'--------------------------------------------------------------------------------------------
' Función:	sendMailToLegales
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	31/01/11
' Objetivo:	Envia Mail a Legales para Aprobar el contrato.
' Parametros:
'			pIdContrato 	[int] 	ID contrato
'			pIdPedido		[int]	ID pedido
'			pIdProveedor	[int] 	ID proveedor asignado
' Devuelve:	valor booleano
'--------------------------------------------------------------------------------------------
Function sendMailToLegales(pIdContrato, pTitulo, pIdPedido, pIdProveedor, pCdResponsable)
	Dim asunto, strMsg, mailFrom, mailTo, rs, conn, cdPedido
	sendMailToLegales = false

	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", "Select CDPEDIDO from TBLPCTCABECERA where IDPEDIDO="&pIdPedido)
	
	if (not rs.eof) then cdPedido = rs("CDPEDIDO")

	strMsg = strMsg & "Se ha publicado un nuevo contrato en el Sistema. Se solicita que ingrese al sistema para confirmarlo." & vbCrLf
	strMsg = strMsg & vbCrLf & vbCrLf
	strMsg = strMsg & "Datos del contrato" & vbCrLf
	strMsg = strMsg & "------------------" & vbCrLf
	strMsg = strMsg & "Titulo................: " & pTitulo & vbCrLf
	strMsg = strMsg & "Asignado a Pedido.....: " & cdPedido & vbCrLf
	strMsg = strMsg & "Responsable...........: " & pCdResponsable & " - " & getUserDescription(pCdResponsable) & vbCrLf
	strMsg = strMsg & "Proveedor Elegido.....: " & pIdProveedor & " - " & getDescripcionProveedor(pIdProveedor) & vbCrLf
	strMsg = strMsg & "Cuit Proveedor........: " & GF_STR2CUIT(getCUITProveedor(pIdProveedor)) & vbCrLf

	mailFrom = getUserMail(pCdResponsable)
	if (mailFrom = "") then mailFrom = obtenerMail("99999997")
	mailTo = SENDER_LEGALES
	asunto = GF_TRADUCIR("Sistema de Compras Web - Alerta Contrato")
	Call GP_ENVIAR_MAIL(asunto, strMsg, mailFrom, mailTo)

	sendMailToLegales = true
End Function
'--------------------------------------------------------------------------------------------
' Función:	canConfirmCTC
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	09/02/11
' Objetivo:	Controla si el usuario puede confirmar contratos.
' Parametros:
'			cdUsuario 	[str]
' Devuelve:	valor booleano
'--------------------------------------------------------------------------------------------
Function canConfirmCTC(cdUsuario, idContrato)
	Dim rs, conn, strSQL
	canConfirmCTC = false
	
	strSQL = "Select CONFIRMACONTRATOS from TBLREGISTROFIRMAS where CDUSUARIO = '" & cdUsuario & "'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		if (rs("CONFIRMACONTRATOS") = 1) then canConfirmCTC = true
	end if
End Function
'--------------------------------------------------------------------------------------------
' Función:	confirmarContrato
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	08/02/11
' Modificacion:
'			CNA - Ajaya Nahuel
' Objetivo:	Controla los datos que fueron asignados por personal de legales
'			para autorizar el contrato. Se agrego la posibilidad de poder guardar el archivo adjunto sin la necesidad
'			de hacerlo cuando se confirma.
' Parametros:
'			pIdContrato 	[int] 	ID contrato
'			pCdContrato		[str]	CD contrato
'			pFile			[str] 	Archivo adjunto al contrato
' Devuelve:	valor booleano
'--------------------------------------------------------------------------------------------
Function confirmarContrato(pIdContrato, pCdContrato, pFile)
	Dim auxCdContrato,auxIdPedido,auxEstado
	auxCdContrato = pCdContrato	
	auxEstado     = ESTADO_CTC_AUTORIZADO
	auxIdPedido	  = 0	
	rtrn = false
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACONTRATOS_GET_BY_IDCONTRATO", pIdContrato)
	if (not rs.Eof) then
		auxIdPedido = CInt(rs("IDPEDIDO"))
		if (Cdbl(rs("ESTADO")) <> ESTADO_CTC_PENDIENTE) Then auxEstado     = CInt(rs("ESTADO"))		
		rtrn = saveDataConfirmCTC(pIdContrato,auxCdContrato,auxIdPedido,auxEstado,pFile)
	End if	
	confirmarContrato = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------
' Función:	
'			saveDataConfirmCTC
' Autor: 	
'			CNA - Ajaya Nahuel
' Fecha: 	
'			07/04/14
' Objetivo:	
'			Guarda los nuevos valores en la base de datos,que fueron asignados por personal de legales
'			para autorizar el contrato.
' Parametros:
'			pIdContrato 	[int] 	ID contrato
'			pCdContrato		[str]	Codigo del contrato
'			pIdPedido		[int] 	Id del Pedido del contrato
'			pEstado			[int] 	Estado del contrato
'			pFile			[str] 	Archivo adjunto al contrato
' Devuelve:	
'			valor booleano - [true = OK] - [false = ERROR]
'-----------------------------------------------------------------------------------------------------------------
Function saveDataConfirmCTC(pIdContrato,pCdContrato,pIdPedido,pEstado,pFile)	
	Dim path, binaryFile, rs, conn, strSQL, rtrn, extencion
	binaryFile = null
	extencion = null	
	'unico control obligatorio de la autorizacion
	if (pCdContrato <> "") then
		' la carga del archivo es opcional
		if (pFile <> "") then
			Set fso = CreateObject("Scripting.FileSystemObject")
			'path del archivo fisico			
			path = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & pFile
			'se convierte el archivo a binario
			binaryFile = readBinaryFile(path)
			'se toma su extencion
			extencion = fso.GetExtensionName(path)
			'se elimina el fisico
			fso.DeleteFile(path)
			Set fso = nothing
			hayArchivo = true
		end if
		'se guardan los nuevos datos, se realiza de esta manera ya que un update convencional
		'no funciona para levantar un valor binario
		strSQL = "SELECT * FROM TBLOBRACONTRATOS WHERE IDCONTRATO = " & pIdContrato
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if not rs.eof then	
		    auxIdPedido = rs("IDPEDIDO")
		    'Se inserta a travez de un store procedure. Para manejar los parametros binarios no se utiliza la funcion standard ya que solo recibe parametros de texto y/o numeros.
		    Set con = server.CreateObject("ADODB.connection")
            con.CursorLocation = 3            
            'con.open session("conn" & CONEXION_AS400 &  "Alias"),  session("conn" & CONEXION_AS400 &  "User"), session("conn" & CONEXION_AS400 &  "Key")
            con.open session("conn" & DBSITE_SQL_INTRA &  "CS")
		    Set cmd = Server.CreateObject("ADODB.Command")
            Set cmd.ActiveConnection = con
            cmd.CommandText = "TBLOBRACONTRATOS_UPD_CONFIRMAR_CONTRATO"
            cmd.CommandType = 4
            cmd.Parameters.Refresh
    		cmd.Parameters(1) = Cstr(pIdContrato)
    		cmd.Parameters(2) = Cstr(pCdContrato)
    		cmd.Parameters(3) = Cstr(ESTADO_CTC_AUTORIZADO)
    		cmd.Parameters(4) = Cstr(session("MmtoDato"))
    		cmd.Parameters(5) = Cstr(session("Usuario"))
            if (hayArchivo) then
    		    cmd.Parameters(6) = Cstr(binaryFile)
    		    cmd.Parameters(7) = Cstr(extencion)
    		else    		
    		    cmd.Parameters(6) = ""
    		    cmd.Parameters(7) = ""
    		end if
            Set rs = cmd.Execute
		end if
		'Una vez aprobado el Contrato por Legales, si tiene Pedido relacionado se lo Aprueba.
		if auxIdPedido > 0 then
			strSQL = "Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_APROBADO & " WHERE IDPEDIDO = "& auxIdPedido
			Call executeQueryDb(DBSITE_SQL_INTRA, rsPCT, "UPDATE", strSQL)
		end if
		rtrn = true
	else
		setError(CODIGO_VACIO)
		rtrn = false
	end if

	saveDataConfirmCTC = rtrn
End Function
'--------------------------------------------------------------------------------------------
' Función:	getCTCDBFile
' Autor: 	JPS - Santi Juan Pablo
' Fecha: 	08/02/11
' Objetivo:	Obtiene el archivo binario de un contrato de la base de datos.
'			tambine devuelve el nombre con su extencion para poder abrirlo.
' Parametros:
'			pIdContrato 	[int] 	ID contrato
'			binaryFile		[BLOB]	archivo adjunto al contrato en binario
'			fileName		[str]	nombre que se usara para abrir el archivo
' Devuelve:	binario del archivo alojado en la DB y el nombre del mismo con su extención
'--------------------------------------------------------------------------------------------
Function getCTCDBFile(pIdContrato, ByRef binaryFile, ByRef fileName)
	dim rs
	getCTCFile= false
	Set rs = readCTC(pIdContrato)
	binaryFile = rs("ARCHIVO")
	fileName = PREFIX_CTC & "-" & rs("CDCONTRATO") & "." & rs("ARCHIVO_EXT")
	getCTCFile= true
End Function
'--------------------------------------------------------------------------------------------
'Funcion que devuelve el importe total facturado o comprometido por PIC del contrato.
Function readCTCTotalPagado(pIdContrato, pIdObra, pIdArea, pIdDetalle, ByRef pImportePesos, ByRef pImporteDolares, includeFR)
	Dim strSQL, conn, rs

	strSQL=	"SELECT " & _
			" sum(CTZD.IMPORTEPESOS)		AS IMPORTEPESOSPIC, " & _
			" sum(CTZD.IMPORTEDOLARES)	AS IMPORTEDOLARESPIC " & _
			" FROM tblctzcabecera CTZ " &_
			"   INNER JOIN tblctzdetalle CTZD on CTZ.IDCOTIZACION=CTZD.IDCOTIZACION " &_
			"WHERE CTZ.estado <> '" & CTZ_ANULADA & "' AND CTZ.IDCONTRATO = " & pIdContrato
    if (CLng(pIdObra) > 0) then
            strSQL= strSQL & " and IDOBRA=" & pIdObra
    end if
    if (CInt(pIdArea) > 0) then
            strSQL= strSQL & " and IDAREA=" & pIdArea
    end if
    if (CInt(pIdDetalle) > 0) then
        strSQL= strSQL & " and IDDETALLE="  & pIdDetalle
    end if        
	if (not includeFR) then	
		strSQL= strSQL & " and IDARTICULO not in ("  & ITEM_FONDO_REPARO_ARS & ", " & ITEM_FONDO_REPARO_USD & ", " & ITEM_FONDO_REPARO_ARS_IVA & ", " & ITEM_FONDO_REPARO_USD_IVA & ")"
    end if        
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

	pImportePesos	= 0
	pImporteDolares	= 0
	if (not isNull(rs("IMPORTEPESOSPIC"))) then
		pImportePesos	= CDbl(rs("IMPORTEPESOSPIC"))
		pImporteDolares	= cDbl(rs("IMPORTEDOLARESPIC"))
	end if
End Function
'--------------------------------------------------------------------------------------------
'Funcion que devuelve el codigo de un determido contrato.
Function getCodigoCTC(pIdContrato)
	Dim strSQL, conn, rs,rtrn
	rtrn= ""
	strSQL="Select CDCONTRATO from TBLOBRACONTRATOS where IDCONTRATO = " & pIdContrato
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.eof)then rtrn = rs("CDCONTRATO")
	getCodigoCTC = rtrn
End Function
'--------------------------------------------------------------------------------------------
' Función:	getDsTipoCTC
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	01/03/2013
' Objetivo:	Obtiene la descripcion de un determinado Tipo de Contrato por medio de su codigo.
' Parametros:
'			pTipo	 	[char] 	Tipo de Contrato
' Devuelve:	Descripcion del Tipo [string]
'--------------------------------------------------------------------------------------------
Function getDsTipoCTC(pTipo)
	Dim strTipo		
	select case pTipo
		case CTC_TIPO_OBRA			
			strTipo = "Pago c/Certificados de Obra"
		case CTC_TIPO_CUOTA_FIJA
			strTipo = "Servicio Repetitivo c/Cuota Fija"			
		case CTC_TIPO_UNITARIO
			strTipo = "Servicio Repetitivo c/Valor Unitario Fijo"			
		case CTC_TIPO_GENERAL
			strTipo = "General"			
	end select	
	getDsTipoCTC = strTipo
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			  buscarFechaVtoCTC
' Autor: 	 
'			  CNA - Ajaya Nahuel
' Autor Modificación: 	 
'			  CNA - Ajaya Nahuel
' Fecha: 	 
'			  01/03/2013
' Fecha Modificación: 	 
'			  04/11/2013
' Objetivo:	  
'  			  Obtiene todos los CTC que esten autorizados por legales y que no hallan finalizado, devolviendo solo 
'			  aquellos que su fecha de vencimiento no supere la que se pasa por parametro, además se podra buscar
'			  por el Tipo de Contrato	
' Parametros: 
'		      pFechaHasta		[date] 	Fecha Fin
'		      pTipoCTC			[string]Tipo de Contrato (Tipo Obra [O],Tipo General [G])
' Devuelve:	  
'			  RecordSet
'--------------------------------------------------------------------------------------------
Function buscarCTCporVencer(pFechaHasta, pTipoCTC)
	Dim strSQL
	strSQL = "		      SELECT CTC.IDCONTRATO,			  "
	strSQL = strSQL & "	  		 CTC.CDCONTRATO,			  "
	strSQL = strSQL & "			 CTC.FECHAVTO,				  "
	strSQL = strSQL & "			 CTC.CDRESPONSABLE,			  "
	strSQL = strSQL & "			 CTC.TITULO,				  "
	strSQL = strSQL & "			 PCT.CDPEDIDO,				  "
	strSQL = strSQL & "			 CASE WHEN CTC.TITULO IS NULL THEN PCT.TITULO ELSE CTC.TITULO END TITULO,				  "
	strSQL = strSQL & "			 CTC.IDPROVEEDOR,			  "
	strSQL = strSQL & "			 CTC.TIPO		        	  "			
	strSQL = strSQL & "	  FROM (SELECT *                      "
	strSQL = strSQL & "         FROM TBLOBRACONTRATOS "
	strSQL = strSQL & "         WHERE ESTADO >= " & ESTADO_CTC_AUTORIZADO & " AND ESTADO < " & ESTADO_CTC_FINALIZADO
	strSQL = strSQL & "			AND FECHAVTO < " & pFechaHasta
	if (pTipoCTC <> "") then strSQL = strSQL & "	AND TIPO = '" & pTipoCTC & "'"
	strSQL = strSQL & "			) CTC "
	strSQL = strSQL & "	  LEFT JOIN TBLPCTCABECERA PCT ON PCT.IDPEDIDO=CTC.IDPEDIDO "
	strSQL = strSQL & "	  ORDER BY CTC.FECHAVTO,			  "
	strSQL = strSQL & "		 	   CTC.IDCONTRATO 			  "	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set buscarCTCporVencer = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   actualizarFechaVto
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   05/03/2013
' Objetivo:
'			   Actualiza solo la fecha de Vencimiento de un Contrato.
' Parametros:
'			   pIdCtc		[int]	ID Contrato
'			   pFechaVto		[date]	Fecha Vencimiento
' Devuelve:
'			    -
'--------------------------------------------------------------------------------------------
Function actualizarFechaVto(pIdCtc, pFechaVto)
	Dim strSQL ,rs
	strSQL = "UPDATE TBLOBRACONTRATOS SET FECHAVTO = '" & pFechaVto & "' WHERE IDCONTRATO = " & pIdCtc
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   readCTCPartida
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Modificacion: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   17/10/2013
' Fecha Modificacion: 	   
'			   28/10/2013
' Objetivo:
'			   Devuelve todas las partidas que estan asociadas a un Contrato.
' Parametros:
'			   pIdContrato		[int]	ID Contrato
' Devuelve:
'			   Recordset
'--------------------------------------------------------------------------------------------
Function readCTCPartida(pIdContrato)
	Dim strSQL, rs
	strSQL = "  SELECT A.*, CASE when A.IDOBRA = " & OBRA_GEID & " then '" & OBRA_GECD & "' else B.CDOBRA end CDOBRA , CASE when A.IDOBRA = " & OBRA_GEID & " then '" & OBRA_GEDS & "' else B.DSOBRA end DSOBRA, C.IDDIVISION "&_
			 "	FROM TBLCTCPARTIDAS A "&_
			 "		INNER JOIN TBLOBRACONTRATOS C ON C.IDCONTRATO = A.IDCONTRATO "&_
			 "		LEFT JOIN TBLDATOSOBRAS B ON B.IDOBRA = A.IDOBRA "&_
			 "	WHERE A.IDCONTRATO = "& pIdContrato &_
			  " ORDER BY A.FECHACIERRE, A.FECHAINICIO,A.IDOBRA, A.IDAREA, A.IDDETALLE "			  
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readCTCPartida = rs	
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   updateEstadoCTCPartida
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   17/09/2013
' Objetivo:
'			   Actualiza el estado de la Partida del Contrato
' Parametros:
'			   pIdContrato		[int]	ID Contrato
'			   idObra			[int]	ID Obra
'			   areaObra			[int]	ID Area
'			   detalleObra		[int]	ID Detalle
'			   estado			[int]	Estado
' Devuelve:    -
'--------------------------------------------------------------------------------------------
Function updateEstadoCTCPartida(idContrato,idobra,areaObra,detalleObra,estado)
	Dim strSQL,rs
	strSQL = "UPDATE TBLCTCPARTIDAS SET ESTADO = "& estado &_
			 " WHERE IDCONTRATO="&idContrato&" AND IDOBRA = "&idobra&" AND IDAREA ="&areaObra&" AND IDDETALLE ="&detalleObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   getPartidaCTCByStatus
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   17/09/2013
' Objetivo:
'			   Devuelve la partida del Contrato que tenga el estado pasado
' Parametros:
'			   pIdContrato		[int]	ID Contrato
'			   pEstado			[int]	ID Estado
' Devuelve:
'			   Recordset
'--------------------------------------------------------------------------------------------
Function getPartidaCTCByStatus(pIdContrato,pEstado)
	Dim strSQL,rs	
	strSQL = " SELECT * FROM TBLCTCPARTIDAS" &_
	         " WHERE    IDCONTRATO=" & pIdContrato &_ 
	         "          AND ESTADO = " & pEstado &_
	         "          AND FECHAINICIO <= " & left(session("MmtoDato"), 8) &_
	         "          AND FECHACIERRE >= " & left(session("MmtoDato"), 8)	         
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getPartidaCTCByStatus = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   grabarPartidaCTC
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Autor Modificacion: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   17/09/2013
' Fecha Modificación: 	   
'			   22/10/2013
' Objetivo:
'			   Graba los datos de la Partida Presupuestaria de un Contrato en la tabla CTCPARTIDAS
' Parametros:
'			   pIdContrato		[int]	ID Contrato
'			   pIdObra			[int]	ID Obra
'			   pIdArea			[int]	ID Area
'			   pIdDetalle		[int]	ID Detalle
'			   pFechaEmision	[int]	Fecha Emision
'			   pFechaVto		[int]	Fecha Vto
' Devuelve:	   -
'--------------------------------------------------------------------------------------------
Function grabarPartidaCTC(pIdContrato, pIdObra, pIdArea, pIdDetalle, pFechaVto, pFechaEmision, pCdMoneda, pImporte, pResponsable)
	Dim strSQL
	strSQL = "SELECT * FROM TBLCTCPARTIDAS WHERE IDCONTRATO="&pIdContrato&" AND IDOBRA="&pIdObra&" AND IDAREA="&pIdArea&" AND IDDETALLE="&pIdDetalle	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	If not rs.Eof then
		strSQL = "UPDATE TBLCTCPARTIDAS SET FECHAINICIO ="&pFechaEmision&",FECHACIERRE ="&pFechaVto&", IMPORTEASIGNADO=" & pImporte & ", CDUSUARIO ='"&pResponsable&"',MOMENTO='"&session("MmtoSistema")&"' "&_
				 "WHERE IDCONTRATO="&pIdContrato&" AND IDOBRA="&pIdObra&" AND IDAREA="&pIdArea&" AND IDDETALLE="&pIdDetalle		
        Call executeQueryDb(DBSITE_SQL_INTRA, rsPartidas, "UPDATE", strSQL)
	else
		strSQL = "INSERT INTO TBLCTCPARTIDAS(IDCONTRATO,IDOBRA,IDAREA,IDDETALLE,FECHAINICIO,FECHACIERRE, CDMONEDA, IMPORTEASIGNADO, IMPORTEGASTADO,CDUSUARIO,MOMENTO) "&_
				 "VALUES(" & pIdContrato & "," & pIdObra & "," & pIdArea & "," & pIdDetalle & ","& pFechaEmision &"," & pFechaVto & ",'" & pCdMoneda & "'," & pImporte & ", 0, '" & pResponsable & "'," & session("MmtoSistema") & ")"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsPartidas, "EXEC", strSQL)
	end if
End Function
'--------------------------------------------------------------------------------------------
' Función:                                                                                  -
'			   actualizarSaldoPendiente                                                     -
' Autor:                                                                                    -
'			   JCZ - Jonathan G. Costilla                                                   -
' Fecha: 	                                                                                -
'			   18/05/2016                                                                   -
' Objetivo:                                                                                 -
'			   Actualizar la tabla TOEPFERDB.TBLOBRACONTRATOS campo SALDO para mejor el     -
'               rendimiento los stored procedures                                           -
' Parametros:                                                                               -
'			   p_APagarPesos	[int]	importe en Pesos                                    -
'			   p_APagarDolares	[int]	importe en Dolares                                  -
'			   p_cdMoneda		[str]	tipo de moneda                                      -
'			   p_idContrato		[int]	ID Contrato                                         -
'			   p_ImporteOld     [int]	Importe Anterior                                    -
'			   p_IsNewPIC		[booleano]	  true/false                                    -
' Devuelve: -                                                                               -
'--------------------------------------------------------------------------------------------
Function actualizarSaldoPendiente(p_APagarPesos,p_APagarDolares,p_cdMoneda,p_idContrato, p_idObra, p_idArea, p_idDetalle,p_ImporteOld,p_IsNewPIC)
    Dim strSQL, ImportePIC, rs
    
    ImportePIC = p_APagarPesos
    if (p_cdMoneda = MONEDA_DOLAR) then ImportePIC = p_APagarDolares
    
    if (p_IsNewPIC) then
        strSQL = " UPDATE TBLOBRACONTRATOS SET SALDO = SALDO - "& ImportePIC  &_
                 " WHERE IDCONTRATO = " & p_idContrato
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
        strSQL= "Update TBLCTCPARTIDAS SET IMPORTEGASTADO=IMPORTEGASTADO + " & ImportePIC & " where IDCONTRATO=" & p_idContrato & " and IDOBRA=" & p_idObra & " and IDAREA=" & p_idArea & " and IDDETALLE=" & p_idDetalle
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
    else
        strSQL = " UPDATE TBLOBRACONTRATOS SET SALDO = SALDO + "& p_ImporteOld &" - "& ImportePIC &_
                 " WHERE IDCONTRATO = " & p_idContrato
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
        strSQL= "Update TBLCTCPARTIDAS SET IMPORTEGASTADO=IMPORTEGASTADO - " & p_ImporteOld & " + " & ImportePIC & " where IDCONTRATO=" & p_idContrato & " and IDOBRA=" & p_idObra & " and IDAREA=" & p_idArea & " and IDDETALLE=" & p_idDetalle
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
    end if
    
End Function
'--------------------------------------------------------------------------------------------
' Función:                                                                                  -
'			   ajusteSaldoPendiente                                                         -
' Autor:                                                                                    -
'			   JCZ - Jonathan G. Costilla                                                   -
' Fecha: 	                                                                                -
'			   19/05/2016                                                                   -
' Objetivo:                                                                                 -
'		       Recalcular saldo pendiente en la base de datos.                              -
'              Tomando todos los PIC hasta la Fecha                                         -
' Parametros:                                                                               -
'			   p_idContrato		[int]	ID Contrato                                         -
' Devuelve: -                                                                               -
'--------------------------------------------------------------------------------------------
Function ajusteSaldoPendiente(p_idContrato)
        Dim strSQL, rs, myImporteCEC, auxPesosCEC, auxDolaresCEC
        
        Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACONTRATOS_UPD_SALDO_BY_IDCONTRATO", p_idContrato)	
        strSQL="Select * from TBLCTCPARTIDAS where idcontrato=" & p_idContrato
	    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	    while not rs.eof
            Call readCTCTotalPagado(p_idContrato, rs("IDOBRA"), rs("IDAREA"), rs("IDDETALLE"), auxPesosCEC, auxDolaresCEC, False)		
            myImporteCEC = auxPesosCEC
            if (rs("CDMONEDA") = MONEDA_DOLAR) then myImporteCEC = auxDolaresCEC
            strSQL="Update TBLCTCPARTIDAS Set IMPORTEGASTADO=" & myImporteCEC & " where IDCONTRATO=" & rs("IDCONTRATO") & " and IDOBRA=" & rs("IDOBRA") & " and IDAREA=" & rs("IDAREA") & " and IDDETALLE=" & rs("IDDETALLE")    
            Call executeQueryDb(DBSITE_SQL_INTRA, rsX, "UPDATE", strSQL)
            rs.MoveNext()
        wend            
End Function
'--------------------------------------------------------------------------------------------
' Función:	 
'			   sendMailCTCPartida
' Autor: 	  
'			   CNA - Ajaya Nahuel
' Fecha: 	   
'			   22/10/2013
' Objetivo:
'			   Envia un mail a los Responsables del Contrato indicando que se cambio la Partida Presupustaria
'--------------------------------------------------------------------------------------------
Function sendMailCTCPartida(idObra,idContrato, titulo, pResponsable)
	Dim mailMsj ,mailCoor, mailTo, rsPartidas
	mailCoor = getMailCoordinadorPto()
	mailTo = mailCoor & getUserMail(pResponsable)
	mailMsj = "Se ha modificado la Partida Presupuestaria del contrato " & getCodigoCTC(idContrato) & " " & titulo & vbCrLf
	mailMsj = mailMsj & "La nueva Partida Presupuestaria es: " & getDescripcionObra(idObra) & vbCrLf
	mailMsj = mailMsj & vbCrLf & vbCrLf
	Set rsPartidas = getPartidaCTCByStatus(idContrato,ESTADO_ACTIVO)
	mailMsj = mailMsj & "Responsable del cambio: " & rsPartidas("CDUSUARIO") & " - " & getUserDescription(rsPartidas("CDUSUARIO"))
	Call GP_ENVIAR_MAIL("Sistema de Compras Web - Modificación Partida en un contrato" , mailMsj, obtenerMail(CD_TOEPFER), mailTo)
End Function
'-------------------------------------------------------------------------------------
'Funcion que devuelve el Tipo de un determinado contrato.
Function getTipoCTC(pIdContrato)
	Dim strSQL, conn, rs,rtrn
	rtrn= ""
	strSQL="Select TIPO from TBLOBRACONTRATOS where IDCONTRATO = " & pIdContrato
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.eof)then rtrn = rs("TIPO")
	getTipoCTC = rtrn
End Function
'-------------------------------------------------------------------------------------
'Especifica si un determinado tipo de contratos tiene valor unitario para trabajar.
Function tieneValorUnitario(pTipo)
    Dim rtrn
    
    rtrn=false
    if (pTipo = CTC_TIPO_UNITARIO) then rtrn = true
        
    tieneValorUnitario = rtrn
End Function
'-------------------------------------------------------------------------------------
'Lee los datos de un contrato y los carga en las variables de trabajo.
Function loadContrato(pIdContrato, pCdMoneda)
    Dim ret, rsCTC, myImporte, myImporteUnitario, tipoCambio, mySaldo
    
    ret = false
    
    CTC_idContrato = pIdContrato
    Set rsCTC = readCTC(CTC_idContrato)    
    if (not rsCTC.eof) then         
        
        CTC_idPedido = cDbl(rsCTC("IDPEDIDO"))
        Call initHeaderDB(CTC_idPedido)

        CTC_cdContrato = rsCTC("CDCONTRATO")
        CTC_Titulo = rsCTC("TITULO")
        CTC_idProveedor = rsCTC("IDPROVEEDOR")
        CTC_dsProveedor = getDescripcionProveedor(rsCTC("IDPROVEEDOR"))
        CTC_FReparo = cInt(rsCTC("FONDOREPARO"))
        CTC_cdResponsable = rsCTC("CDRESPONSABLE")
        CTC_dsResponsable = getUserDescription(rsCTC("CDRESPONSABLE"))
        CTC_cdMoneda = rsCTC("CDMONEDA")
        'Tomo el importe que es válido,  y es el que coincide con la moneda del contrato.
        if (CTC_cdMoneda = MONEDA_PESO) then 
            myImporte = cDbl(rsCTC("IMPORTEPESOS"))
            myImporteUnitario  = cDbl(rsCTC("IMPORTEUNITARIOPESOS"))            
        else
            myImporte = cDbl(rsCTC("IMPORTEDOLARES"))
            myImporteUnitario  = cDbl(rsCTC("IMPORTEUNITARIODOLARES"))            
        end if        
	mySaldo = cDbl(rsCTC("SALDO"))	
        if (pCdMoneda = "") then pCdMoneda = CTC_cdMoneda
        'Convierto a la moneda que se pide en pantalla. 	
        if (pCdMoneda <> CTC_cdMoneda) then             	
            tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
            if (CTC_cdMoneda = MONEDA_PESO) then 
                myImporte = myImporte/tipoCambio
                myImporteUnitario = myImporteUnitario/tipoCambio
                mySaldo = mySaldo/tipoCambio		
            else
                myImporte = myImporte*tipoCambio
                myImporteUnitario = myImporteUnitario*tipoCambio
                mySaldo = mySaldo*tipoCambio
            end if            
            CTC_cdMoneda = pCdMoneda
        end if
        CTC_TotalImporte = myImporte
        CTC_valorUnitario = myImporteUnitario	
        CTC_Total_ImporteSaldo = mySaldo
        CTC_estado = cInt(rsCTC("ESTADO"))
        CTC_fechaVto = rsCTC("FECHAVTO")
        CTC_tipo = rsCTC("TIPO")        
        if(rsCTC("TIPO") = "")then CTC_tipo = CTC_TIPO_OBRA
        CTC_idDivision = rsCTC("IDDIVISION")
		
        ret = true
    end if
    
    loadContrato = ret
End Function
'---------------------------------------------------------------------------------------------
Sub delAJUCTCFirmas(pIdAjuste)
dim strSQL, rs, conn, rsDel, connDel
	strSQL="Select * from TBLOBRACTCAJUSTESFIRMAS where idAjuste = " & pIdAjuste
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		strSQL = "Delete from TBLOBRACTCAJUSTESFIRMAS where idAjuste = " & pIdAjuste		
        Call executeQueryDb(DBSITE_SQL_INTRA, rsDel, "EXEC", strSQL)
	end if
end sub
'---------------------------------------------------------------------------------------------
'Función responsable por dejar el esquema de firmas de un Ajuste de Contrato según las reglas definidas por la empresa.
Function addAJUCTCFirmas(pIdAjuste, idDivision, cdSolicitante, cdAutorizante)

    Dim auxUser    
    
    Call delAJUCTCFirmas(pIdAjuste)
    
    ' Solicitante	
	strSQL = "INSERT INTO TBLOBRACTCAJUSTESFIRMAS (IDAJUSTE, SECUENCIA, CDUSUARIO) values(" & myIdAjuste & ", " & CTC_FIRMA_RESPONSABLE & ",'" & cdSolicitante & "')"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	
	' Solicitante	
	strSQL = "INSERT INTO TBLOBRACTCAJUSTESFIRMAS (IDAJUSTE, SECUENCIA, CDUSUARIO) values(" & myIdAjuste & ", " & CTC_FIRMA_GTE_SECTOR & ",'" & cdAutorizante & "')"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	    
    '-- Gerente de Compras
    strSQL = "INSERT INTO TBLOBRACTCAJUSTESFIRMAS (IDAJUSTE, SECUENCIA, CDUSUARIO) values(" & myIdAjuste & ", " & CTC_FIRMA_GTE_COMPRAS & ",'" & FIRMA_NO_USER & "')"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    
End Function                        
'--------------------------------------------------------------------------------------------
Function actualizarResponsableCTC(p_IdContrato, p_CdResponsable)
	Dim strSQL ,rs
	strSQL = "UPDATE TBLOBRACONTRATOS SET CDRESPONSABLE = '" & p_CdResponsable & "' WHERE IDCONTRATO = " & p_IdContrato
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function  
'--------------------------------------------------------------------------------------------
Function getCTCTotalAsignado(pIdContrato)
    Dim strSQL, rs
    
    strSQL="Select SUM(IMPORTEASIGNADO) IMPORTE from TBLCTCPARTIDAS where IDCONTRATO=" & pIdContrato
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    getCTCTotalAsignado = 0
    if (not isNull(rs("IMPORTE"))) then getCTCTotalAsignado = CDbl(rs("IMPORTE"))
    
End function
'--------------------------------------------------------------------------------------------
Function leerBudgetActivosCTC(p_IdContrato)
    Dim strSQL, conn, rs
    
    'Tomo la obra activa.			
    strSQL ="Select IDOBRA from  TBLCTCPARTIDAS " &_
            " where IDCONTRATO=" & p_IdContrato & " and FECHACIERRE >= " & Left(session("MmtoDato"), 8) & " and FECHAINICIO <= " & Left(session("MmtoDato"), 8)
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    
    if (not rs.eof) then
        'Tomo los detalles disponibles.    			
	    strSQL ="Select CP.*, CA.DSBUDGET DSCABECERA, CB.DSBUDGET DSDETALLE from TBLCTCPARTIDAS CP " &_
	            " left join TBLBUDGETOBRAS CA on CP.IDOBRA=CA.IDOBRA and CP.IDAREA=CA.IDAREA and CA.IDDETALLE=0" &_ 
	            " left join TBLBUDGETOBRAS CB on CP.IDOBRA=CB.IDOBRA and CP.IDAREA=CB.IDAREA and CP.IDDETALLE=CB.IDDETALLE " &_ 
	            " where IDCONTRATO=" & p_IdContrato & " and CP.IDOBRA=" &  rs("IDOBRA")&_ 
	            " Order by CP.IDAREA, CP.IDDETALLE"	        	        
	    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    end if	    
	Set leerBudgetActivosCTC = rs
	
End Function


%>