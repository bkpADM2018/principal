<%
response.Charset="ISO-8859-1"
session.LCID = 2058
%>
<!--#include file="procedimientosMensajes.asp"-->
<!--#include file="procedimientosMath.asp"-->
<!--#include file="globalConfiguration.asp"-->
<!--#include file="procedimientosConexion.asp"-->
<%
Const ESTADO_ACTIVO = 1
Const ESTADO_BAJA = 2
Const ESTADO_ANULACION = 3
Const ESTADO_AUTORIZADO = 4

Const TIPO_AFIRMACION = "S"
Const TIPO_NEGACION = "N"

Const ENTER_SYMBOL = "<br>"

Const ACCION_GRABAR = "grabar"
Const ACCION_PROCESAR = "procesar"
Const ACCION_CONTROLAR = "control"
Const ACCION_SUBMITIR = "submit"
Const ACCION_BACH = "bach" 
Const ACCION_BORRAR = "borrar"
Const ACCION_ACTIVAR = "activar"
Const ACCION_CERRAR = "CERRAR"
Const ACCION_CANCELAR = "cancelar"
Const ACCION_CONFIRMAR = "confirmar"
Const ACCION_VALUAR = "valuar"
Const ACCION_CALCULAR = "calcular"
Const ACCION_EMAIL = "mail"
Const ACCION_VISUALIZAR = "ver"

Const DEFAULT_SIGNATURE = "signature-48x48.png"

'/* TIPO DE BIENES REGISTRABLES */
Const ES_BIEN_DE_CONSUMO = "N"
Const ES_BIEN_DE_USO = "S"

'/* TIPOS DE CATEGORIA DE ARTICULOS */
Const TIPO_CAT_BIENES = "B"
Const TIPO_CAT_SERVICIOS = "S"
Const TIPO_CAT_IMPUESTOS = "I"
Const TIPO_CAT_FONDO_REPARO = "F"
Const TIPO_CAT_ANTICIPO = "A"
Const TIPO_CAT_ESPECIAL_IVA = "R"

'/* TIPOS DE CIERRE CONTABLES  */
Const TIPO_CIERRE_PROVISORIO = "P"
Const TIPO_CIERRE_DEFINITIVO = "D"
Const TIPO_CIERRE_DEBE = "1"
Const TIPO_CIERRE_HABER = "2"
'/* MONEDAS */
Const MONEDA_PESO="P"
Const MONEDA_PESOS="P"
Const MONEDA_DOLAR="D"
Const MONEDA_DOLAR_FACTURACION = "U"
Const MONEDA_DOLAR_NUMERICO = 2
Const MONEDA_PESO_NUMERICO  = 1

Const T_CAMBIO_VENDEDOR = "23"
Const T_CAMBIO_COMPRADOR = "43"

'/* ROLES ESPECIALES PARA FIRMAS */
Const FIRMA_ROL_NINGUNO         = 0
Const FIRMA_ROL_AUDITOR         = 8
Const FIRMA_ROL_SUP_PUERTO      = 9     'COORDINADOR DE PUERTOS
Const FIRMA_ROL_RESP_CONTADURIA = 10
Const FIRMA_ROL_RESP_PUERTO     = 11    'GERENTE DEL PUERTO
Const FIRMA_ROL_CONTROLLER      = 12
Const FIRMA_ROL_DIRECTOR        = 13
Const FIRMA_ROL_LEGALES         = 14
Const FIRMA_ROL_IMPUESTOS       = 15
Const FIRMA_ROL_GTE_SECTOR      = 16
Const FIRMA_ROL_GTE_COMPRAS     = 18
'----------REGISTROS DE USUARIOS  -------'
Const FIRMA_NO_USER     = "NAU"
Const CONTROLLER_USER   = "CPS"
Const DIRECTOR_USER     = "DPS"
Const LEGALES_USER      = "LEG"
'/* DELIMITADOR PARA TEXTO COMPUESTO DE MULTIPLES CAMPOS */
Const STRING_DELIMITER = "|"

Const CD_TOEPFER = "99999997"
Const CUIT_TOEPFER = "30621973173"

Const CUIT_ADM = "30697312028"	

Const CUIT_AS400 = 80

Const A_MANO = "A_MANO"

Const CODIGO_EXPORTACION = "E"
Const CODIGO_ARROYO = "N"
Const CODIGO_PIEDRABUENA = "P"
Const CODIGO_TRANSITO = "T"

Const SIN_DIVISION = 0

'/* ESTADO DE LOS PROVEEDORES PARA LA LISTA NEGRA */
Const PROV_ACTIVO = 0 'proveedor activo, no existe en la lista negra
Const PROV_PROHIBIDO_PEDIDOS = 1 'proveedor no puede recibir pedidos
Const PROV_PROHIBIDO_PAGOS = 2 'proveedor no puede recibir pagos

Const PROV_ID_ADM = 13536

'/* TIPO DE DOCUMENTO DE LAS EMPRESAS */
Const TIPO_CUIT_80		= 80 'empresas nacionales
Const TIPO_CUIT_EX_83	= 83 'empresas extranjeras
Const TIPO_CUIL_86     	= 86 'empresas nacionales

'/* CONTANTES DE SEGURIDAD SISTEMAS */
Const SEC_SYS_UNASIGNED     =  0
Const SEC_SYS_GEMINI        =  2
Const SEC_SYS_HEFESTO       =  2
Const SEC_SYS_COMPRAS       =  3
Const SEC_SYS_ALMACENES     =  4
Const SEC_SYS_PROVEEDORES   =  7
Const SEC_SYS_POSEIDON      = 21
Const SEC_SYS_FACTURACION   =  9
Const SEC_SYS_MERCADERIAS   =  5

'/* CONTANTES DE SEGURIDAD DEL TIPO DE USUARIO */
Const SEC_X = 0	'SIN ACCESO
Const SEC_U = 1	'USUARIO
Const SEC_Y = 2	'AUDITOR
Const SEC_A = 3	'ADMINISTRADOR

'/* CONTANTES DE SEGURIDAD - (CARGO/RECURSO) */
'				/* COMPRAS */
Const RES_CC  = 0	'CONCURSOS Y PEDIDOS DE PRECIO
Const RES_CD  = 1	'COMPRA DIRECTA
Const RES_OBR = 2	'OBRAS
Const RES_AFE = 3	'AFES
Const RES_AUD = 4	'AUDITORIA
Const RES_ADM = 5	'ADMINISTRACIÓN
Const RES_CTC = 6	'CONTRATOS
Const RES_PDC = 13	'POLIZA DE CAUCION

'				/* ALMACENES */
Const RES_ADM_AL = 5	'ADMINISTRACIÓN
Const RES_ACC_AL = 6	'CONTABILIDAD

'				/* MANTENIMIENTO */
Const RES_INV_SM	= 20 'Acceso al inventario
Const RES_OT_SM		= 21 'Acceso a las Ordenes de trabajo
Const RES_PLA_SM	= 22 'Acceso a la planificacion de mantenimiento

'               /* PROVEEDORES  */
Const RES_PRV_MASTER= 23 'Acceso al Maestro de Proveedores

'               /*  FACTURACION */
Const RES_FAC_MER = 12 'Acceso a las facturas de Mercaderias      
Const RES_FAC_EJL = 14 'Acceso a las facturas de Ejecucion Locales
Const RES_FAC_EJE =  4 'Acceso a las facturas de Ejecucion Exportacion
Const RES_FAC_CG =  16 'Acceso a las facturas de Contaduria
Const RES_FAC_TRA = 18 'Acceso a las facturas de Transito
Const RES_FAC_ARR = 19 'Acceso a las facturas de Arroyo
Const RES_FAC_LPB = 20 'Acceso a las facturas de Bahia

'               /*  MERCADERIAS  */
Const RES_MER_ANALISIS  = 1  'Acceso a los Analisis
Const RES_MER_RECEPCOND = 2  'Acceso a la recepcion de garantias
Const RES_MER_GTACONTRATO_BSAS = 3  'Acceso a las garantias de contrato de Buenos Aires
Const RES_MER_GTACONTRATO_ROS  = 4  'Acceso a las garantias de contrato de Rosario
           
'               /*  POSEIDON    */
Const RES_PSD_ANALISIS  = 1  'Acceso a los Analisis

'               /*  GEMINI    */
Const RES_GEM_TAREAS  = 1	'Acceso a las Tareas
'/**************** EVENTOS DE TRANSICION ***************
Const EVENTO_FIRMA = 1

'Tipos de documentos a autorizar.
Const AUTH_TYPE_PCP = "PCP"	'Planillas Comparativas
Const AUTH_TYPE_PIC = "PIC"	'Pedidos Internos de Compras
Const AUTH_TYPE_CEC = "CEC"	'Comprobante Electronico de Cumplimiento
Const AUTH_TYPE_AFE = "AFE" 'Autorizaciones para Gastos
Const AUTH_TYPE_AJS = "AJS" 'Ajuste de stock
Const AUTH_TYPE_XJS = "XJS" 'Anulacion de ajuste de stock
Const AUTH_TYPE_AIC = "AIC" 'Ajuste de PIC
Const AUTH_TYPE_AEC = "AEC" 'Ajuste de CEC
Const AUTH_TYPE_VRS = "VRS" 'Vale reclasificacion de stock
Const AUTH_TYPE_XRS = "XRS" 'Anulacin de vale reclasificacion de stock
Const AUTH_TYPE_CTC = "CTC" 'Ajustes de Contratos'
Const AUTH_TYPE_AJD = "AJD" 'Ajuste Draft Survey
Const AUTH_TYPE_AJC = "AJC" 'Ajuste Calidad
Const AUTH_TYPE_AJM = "AJM" 'Ajuste Manipuleo
Const AUTH_TYPE_CCN = "CCN"	'Cierre de Almacenes 
Const AUTH_TYPE_AJV = "AJV" 'Ajuste Merma Volatil
Const AUTH_TYPE_APP = "APP" 'Ajuste de Partidas Presupuestaria
'/******  CONEXION *******/'
Const CONEXION_AS400 = "AS400"

'/******  STORE PROCEDURES *******/'
Const SP_IDERROR = "idError"
Const SP_DSERROR = "dsError"



'/******  EJECUCION INTERNACIONAL (PROVISIONES) *******/'
Const PROVISCIONES_ESTADO_GENERADO   = "1" 'Estado inicial del lote
Const PROVISCIONES_ESTADO_PENDIENTE  = "2" 'Estado autorizado por el solicitante
Const PROVISCIONES_ESTADO_AUTORIZADO = "3" 'Estado autorizado por el jefe del sector
Const PROVISCIONES_ESTADO_APLICADO   = "A" 'Estado autorizado por el controller
Const PROVISCIONES_ESTADO_ERROR      = "E" 'Estado con error

Dim PATH_CARPETA_FIRMAS
PATH_CARPETA_FIRMAS =  server.MapPath(SITE_ROOT) & "\images\firmas\"

Dim ccValorListaDivision(5)
'--------------------------------------------------------------------------------------
function GF_BD_COMPRAS(byref pRS, byref pCon, pOperacion, pSql)
	GF_BD_COMPRAS = executeQueryDb(DBSITE_SQL_INTRA, pRS, pOperacion, pSql)
end function
'--------------------------------------------------------------------------------------'--------------------------------------------------------------------------------------
Function isFormSubmit()
	Dim acc 
	acc = GF_PARAMETROS7("accion","",6)
	if (	(acc = ACCION_GRABAR) _
		or  (acc = ACCION_CONTROLAR) _
		or  (acc = ACCION_SUBMITIR) _
		or  (acc = ACCION_CONFIRMAR) _
		or  (acc = ACCION_BORRAR) _
		or  (acc = ACCION_ACTIVAR) _
		or  (acc = ACCION_CANCELAR)) then
		isFormSubmit = true
	else
		isFormSubmit = false
	end if
End Function
'--------------------------------------------------------------------------------------
Function getAbreviaturaUnidad(codigo) 
	dim strSQL, rs
	
	getAbreviaturaUnidad = ""
	strSQL = "Select * from TBLUNIDADES where IDUNIDAD=" & codigo
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		getAbreviaturaUnidad = rs("ABREVIATURA")
	end if	
End Function
'--------------------------------------------------------------------------------------
Function getCIADivision(pIdDivision) 
	Dim data, strSQL, rs, conn
	getCIADivision = "000"	
	strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & pIdDivision
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then getCIADivision = rs("CIA")					
End Function
'--------------------------------------------------------------------------------------
'Dado un id de division, devuelve Codigo abreviado (N, E, P o T)
Function getDivisionAbreviada(pIdDivision) 
	Dim data, strSQL, rs, conn
		
	getDivisionAbreviada = "Z"	
	strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & pIdDivision
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then getDivisionAbreviada = rs("CDDIVISIONABR")					
End Function
'--------------------------------------------------------------------------------------
'Dado un id de division, devuelve la descripcion (A, E, P o T)
Function getDivisionDS(pIdDivision) 
	Dim data, strSQL, rs, conn
	
	getDivisionDS = "ERROR"	
	strSQL = "Select * from TBLDIVISIONES where IDDIVISION=" & pIdDivision
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then getDivisionDS = rs("DSDIVISION")				
End Function
'--------------------------------------------------------------------------------------
'Dado un cd de division, devuelve el id
Function getDivisionID(pCdDivision) 
	Dim data, strSQL, rs, conn
	getDivisionID = "ERROR"	
	if pCdDivision<>"" then
		strSQL = "Select * from TBLDIVISIONES where CDDIVISIONABR='" & pCdDivision & "'"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then getDivisionID = rs("IDDIVISION")				
	end if
End Function
'---------------------------------------------------------------------------------------------
sub getArticuloFull(pIdArticulo, byref pDS, byref pAbrev)
dim strSQL, rs, conn, rtrn

strSQL="Select TA.IDARTICULO, TA.DSARTICULO, TU.ABREVIATURA from TBLARTICULOS TA INNER JOIN TBLUNIDADES TU ON TU.IDUNIDAD=TA.IDUNIDAD where TA.IDARTICULO=" & pIdArticulo
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then 
	pIdArticulo = rs("IDARTICULO")
	pDS = rs("DSARTICULO")
	pAbrev = rs("ABREVIATURA")
end if
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
end sub
'---------------------------------------------------------------------------------------------
Function getTipoCambio(pCdMoneda, pDate) 	
	getTipoCambio = getTipoCambioCV(pCdMoneda, pDate, T_CAMBIO_COMPRADOR) 	
End Function
'---------------------------------------------------------------------------------------------
Function getTipoCambioCV(pCdMoneda, pDate, pTipo) 	
	Dim rs, strSQL, rtrn	
	rtrn = 0
	if (pDate = "") then pDate = Left(session("MmtoDato"), 8)
	'cgt0012a
	strSQL="select * from CGT012A where CODMON=" & pTipo & " and FECHA <= '" & left(pDate, 4) & "-" & Mid(pDate, 5, 2) & "-" & Right(pDate, 2) & "' and tcambio > 0 order by FECHA desc"	
	call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
	if (not rs.eof) then
		rtrn = CDbl(rs("TCAMBIO"))		
	end if	
	getTipoCambioCV = rtrn
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un articulo.
Function controlarArticulo(pIdArticulo) 	
	Dim rs, strSQL
	controlarArticulo = false
	'Controlo si el articulo existe
	strSQL="select * from TBLARTICULOS where IDARTICULO=" & pIdArticulo & " and ESTADO=" & ESTADO_ACTIVO
	'response.write strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (not rs.eof) then
		controlarArticulo = true		
	else
		Call setError(ARTICULO_NO_EXISTE)		
	end if
End Function
'---------------------------------------------------------------------------------------------
'Devuelve lista de almacenes disponibles
function obtenerListaAlmacenes(p_idAlmacen)
	Dim strSQL, rtrn
	rtrn = 0
	strSQL="select * from TBLALMACENES "
	if not (p_idAlmacen = 0) then strSQL = strSQL & " where IDALMACEN = " & p_idAlmacen
	'Response.Write strsql
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set obtenerListaAlmacenes = rs
End function
'----------------------------------------------------------------
Function getSimboloMoneda(idMoneda)
	getSimboloMoneda = "$"
	Select case (CStr(idMoneda))
		case MONEDA_DOLAR, MONEDA_DOLAR_FACTURACION, CStr(MONEDA_DOLAR_NUMERICO)
			getSimboloMoneda = "u$s"
	End Select
End Function
'----------------------------------------------------------------
Function getSimboloMonedaLetras(idMoneda)
	getSimboloMonedaLetras = "ARS"
	Select case (CStr(idMoneda))
		case MONEDA_DOLAR, MONEDA_DOLAR_FACTURACION, CStr(MONEDA_DOLAR_NUMERICO)
			getSimboloMonedaLetras = "USD"		
	End Select
End Function
'----------------------------------------------------------------
Function getNombreMoneda(idMoneda)
	getNombreMoneda = "PESOS"
	Select case (CStr(idMoneda))
		case MONEDA_DOLAR, MONEDA_DOLAR_FACTURACION, CStr(MONEDA_DOLAR_NUMERICO)
			getNombreMoneda = "DOLARES"		
	End Select
End Function
'----------------------------------------------------------------
Function puedeFirmarAsientos(pUser, pIdDivision)
dim strSQL, rsPFA, conPFA, cdDivision
puedeFirmarAsientos = false
cdDivision = trim(getDivisionAbreviada(pIdDivision))
if cdDivision = "" then 
	puedeFirmarAsientos = false
else
	if (pUser <> "") then
		strSQL = "Select * from TBLREGISTROFIRMAS WHERE CDUSUARIO = '" & pUser & "' AND HKEY<>'' " 
		if cdDivision = CODIGO_EXPORTACION then
			strSQL = strSQL & " AND ASEXPORTACION=1"
		elseif cdDivision = CODIGO_ARROYO then
			strSQL = strSQL & " AND ASARROYO=1"
		elseif cdDivision = CODIGO_PIEDRABUENA then
			strSQL = strSQL & " AND ASPIEDRABUENA=1"
		elseif cdDivision = CODIGO_TRANSITO then
			strSQL = strSQL & " AND ASTRANSITO=1"
		end if
		call executeQueryDb(DBSITE_SQL_INTRA, rsPFA, "OPEN", strSQL)
		if not rsPFA.eof then puedeFirmarAsientos = true
		call executeQueryDb(DBSITE_SQL_INTRA, rsPFA, "CLOSE", strSQL)
	end if
end if	
End Function
'----------------------------------------------------------------
Function getRolFirma(pUser, pIdSystem)
dim rsRF, rtrn, params
rtrn = 0
if (pUser <> "") then    
	CALL executeProcedureDb(DBSITE_SQL_INTRA,rsRF, "TBLROLESUSUARIOS_GET_BY_CDUSER_IDSISTEMA", pUser & "||" & pIdSystem)
	if (not rsRF.eof) then rtrn = CInt(rsRF("IDROL"))
end if
getRolFirma = rtrn
End Function
'----------------------------------------------------------------
'Setup del recordset para paginar en AS400.
Function setupPaginacion(ByRef rs, pag, regXPag) 
	Dim tot
	'Calculo la cantidad de paginas totales en el recordset
	tot = CLng(Ceil(rs.RecordCount / regXPag))
	'Si la pagina pedida no existe, me posiciono en la primera
	if (tot < pag) then pag = 1	
	if (not rs.eof) then
		rs.PageSize= regXPag
		rs.CacheSize = regXPag
		rs.AbsolutePage = pag
	end if
End Function
'----------------------------------------------------------------
'Devuelve las categorias segun el tipo solicitado
function getCategoriasTipo(pTipo)
dim rtrn, rs, conn, strSQL
strSQL="Select * from TBLARTCATEGORIAS where TIPOCATEGORIA = '" & pTipo & "'"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
While (not rs.eof) 
	if (len(rtrn) > 0) then rtrn = rtrn & ","
	rtrn = rtrn & rs("IDCATEGORIA")
	rs.MoveNext()
wend
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
getCategoriasTipo = rtrn
end function
'-----------------------------------------------------------------
Function getComboMonedas(pMoneda)
	Dim rtrn
	Dim vecMonedas(2,2)
	vecMonedas(0,0) = MONEDA_DOLAR
	vecMonedas(0,1) = "Dolares"
	vecMonedas(1,0) = MONEDA_PESO
	vecMonedas(1,1) = "Pesos"
	
	rtrn = "<select name='comboMoneda' id='comboMoneda'>"
	for i = 0 to ubound(vecMonedas) -1
		if (pMoneda = vecMonedas(i,0)) then
			rtrn = rtrn & "<option value='" & vecMonedas(i,0) & "' selected='selected'>" &  vecMonedas(i,1) & "</option>"
		else
			rtrn = rtrn & "<option value='" & vecMonedas(i,0) & "'>" &  vecMonedas(i,1) & "</option>"
		end if
	next
	rtrn = rtrn & "</select>"
	
	getComboMonedas = rtrn
end Function
'--------------------------------------------------------------------------------------
Function getDescripcionProveedor(codigo) 
	dim strSQL, rs
	
	getDescripcionProveedor = ""	
	codigo= Trim(codigo)
	if (isNumeric(codigo)) then
		strSQL="select NOMEMP from MET001A where NROEMP = " & codigo
		Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
		if (not rs.eof) then
			getDescripcionProveedor = rs("NOMEMP")
		end if
	end if
End Function
'--------------------------------------------------------------------------------------
'Funci´no que devuelve la descripcion de un proveedor a partir de su CUIT.
Function getDescripcionProveedorCUIT(cuit) 
	dim strSQL, rs
	
	getDescripcionProveedorCUIT = ""	
	codigo= Trim(cuit)
	if (isNumeric(cuit)) then
		strSQL="select NOMEMP from MET001A where NRODOC = " & cuit
        Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
		if (not rs.eof) then
			getDescripcionProveedorCUIT = rs("NOMEMP")
		end if
	end if
End Function
'--------------------------------------------------------------------------------------
Function getCUITProveedor(codigo) 
	dim strSQL, rs
	
	getCUITProveedor = ""
	if (isNumeric(codigo)) then
		strSQL="select NRODOC from MET001A where NROEMP=" & codigo
		Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
		if (not rs.eof) then
			getCUITProveedor = rs("NRODOC")
		end if
	end if
End Function
'------------------------------------------------------------------------------------
Function armarTextoFirma(hk, mmto)
	if ((hk <> "") and (mmto <> "")) then
		armarTextoFirma = "<div style='font-size:7px' align='center'>" & GF_TRADUCIR("Firma Electrónica") & STRING_DELIMITER & GF_FN2DTE(mmto) & "|HKEY-" & hk & "</div>"
	else
		armarTextoFirma = ""
	end if
End Function
'--------------------------------------------------------------------------------------------------------
Function obtenerFirma(usr)
	Dim  fileFirma, fso
	
	fileFirma = usr & ".png"
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	if (fso.FileExists(PATH_CARPETA_FIRMAS & fileFirma)) then
		obtenerFirma = fileFirma
	else
		obtenerFirma = DEFAULT_SIGNATURE
	end if
	
	Set fso = nothing
End Function
'--------------------------------------------------------------------------------------------------------
Function armarTextoPlanoFirma(hk, mmto)
	if ((hk <> "") and (mmto <> "")) then
		armarTextoPlanoFirma = GF_TRADUCIR("Firma Electrónica") & "|" & GF_FN2DTE(mmto) & "|HKEY-" & hk
	else
		armarTextoPlanoFirma = ""
	end if
End Function
'-----------------------------------------------------------------------------------------
function getSectorDS(pIdSector)
dim rtrn, rs, conn, strSQL
rtrn = ""
strSQL="Select * from TBLSECTORES where IDSECTOR = " & pIdSector
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then rtrn = rs("DSSECTOR")
getSectorDS = rtrn
end function
'-----------------------------------------------------------------------------------------
function getCategoriaDS(pIdCategoria)
dim rtrn, rs, conn, strSQL
rtrn = ""
strSQL="Select * from TBLARTCATEGORIAS where IDCATEGORIA = " & pIdCategoria
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then rtrn = rs("DSCATEGORIA")
getCategoriaDS = rtrn
end function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	22/10/10
' Objetivo:	
'			Obtener el precio del ultimo cierre contable de los articulos pasados por parametro
' Parametros:
'			pArticulos 	[str] 	Articulos separados por ","
'			pDivision 	[int] 	Id de la division 
'			pMoneda		[char]	Moneda con la que se buscara
'			pFecCierre	[int] 	Fecha limite para el cierre contable a buscar
' Devuelve:
'			Diccionario | KEY:idarticulo - VALUE:importe
' Modificaciones:
'			29/10/10 - GFG
'--------------------------------------------------------------------------------------------------
Function obtenerPreciosArticulos(pArticulos,pDivision,pMoneda,pFecCierre)
	Dim dicArticulos,strSQL,conn,rs,campoMoneda,myFecCierre
	
	Set dicArticulos = Server.CreateObject("Scripting.Dictionary")
	if (pFecCierre = "") then pFecCierre = session("MmtoSistema")
	myFecCierre = cstr(pFecCierre)
	if (len(myFecCierre) < 14) then myFecCierre = myFecCierre & "000000"
	
	campoMoneda = "VLUPESOS"
	if (pMoneda = MONEDA_DOLAR) then campoMoneda = "VLUDOLARES"

	strSQL =          "SELECT * "
	strSQL = strSQL & "FROM   TBLARTICULOSPRECIOS PRE inner join "
	strSQL = strSQL & "       (SELECT  MAX(MMTOPRECIO) fecha   , "
	strSQL = strSQL & "                IDARTICULO "
	strSQL = strSQL & "       FROM     TBLARTICULOSPRECIOS "
	strSQL = strSQL & "       WHERE    MMTOPRECIO <= " & myFecCierre
	strSQL = strSQL & "       AND      IDARTICULO IN("&pArticulos&") "
	strSQL = strSQL & "       AND      IDDIVISION = " & pDivision
	strSQL = strSQL & "       GROUP BY IDARTICULO "
	strSQL = strSQL & "       ) "
	strSQL = strSQL & "       FECHA "
	strSQL = strSQL & "		  ON  PRE.MMTOPRECIO = FECHA.fecha "
	strSQL = strSQL & "		  AND PRE.IDARTICULO = FECHA.IDARTICULO "
	strSQL = strSQL & "WHERE PRE.IDDIVISION = " & pDivision
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	dicarticulos.removeAll
	while not rs.EoF
		if (not dicArticulos.Exists(cstr(rs("IDARTICULO")))) then
			dicArticulos.add trim(rs("IDARTICULO")), trim(rs(campoMoneda))
		end if
		rs.MoveNext
	wend
	
	Set obtenerPreciosArticulos = dicArticulos
	
End Function
'-----------------------------------------------------------------------------------------
'CALCULO DE PRECIOS!!!
'-----------------------------------------------------------------------------------------
'Este funcion es la encargada de actualizar los precios cada vez que se modifica el stock de existencias en algun almacen
'Los casos contemplados son:
'	REM y XEM: 
'		Generan un aumento de las existencias en el caso de no tener partida y de sobrante cuando tienen partida
'	VMR y XMR: 
'		Solo importa el caso de las transferencias interdivisionales las cuales generan un aumento en existencias unicamente
'	XMS: 
'		Puede devolver existencia. La misma pudo haber sufrido cambio de precios entre que se genero el VMS y se anulo.
'	XJU: 
'		Puede devolver existencia. La misma pudo haber sufrido cambio de precios entre que se genero el AJU y se anulo.

'Parametros:
'	pId: id del formulario que se cargo
'	pTipoForm: tipo del formulario que se cargo
'a partir del formulario recibido se analiza el tipo de formulario y se recalcula el precio del articulo segun el formulario
function ActualizarPrecios(pId, pTipoForm)
dim mySQL, myRS, mmtoCompra, vluPesosCompra, vluDolaresCompra, idPic
dim myDivisionDeTrabajo, myFechaDeTrabajo, MyIdFormularioDeTrabajo
dim stockActual, stockEnAlmacenes, stockEnCirculacion
dim stockFormulario, stockAnulacion, tipoCambioHoy, precioDePIC
dim precioActual, precioFormulario, precioAnulacion, precioNuevoPesos, precioNuevoDolares,precioOrigen
precioOrigen = 0
'mySQL contendra la sentencia que seleccionara
'	-la division a la cual corresponde el articulo
'	-la fecha en la cual se cargo
'	-el id original (para el caso de anulaciones-transferencias)
'mySQLDetalle contendra la sentencia que seleccionara el detalle del formulario en cuestion
'precioDePIC indica si el precio lo debe ir a buscar al PIC.

precioDePIC = false
	select case UCase(pTipoForm)
		case CODIGO_VS_SALIDA_X, CODIGO_VS_TRANSFERENCIA_X, CODIGO_VS_RECEPCION, CODIGO_VS_RECLASIFICACION_STOCK
			mySQL = "SELECT AL.IDDIVISION AS DIVISION_TRABAJO, VC.MOMENTO AS FECHA_TRABAJO, VC.IDVALE AS ID_FORM_ORG FROM TBLVALESCABECERA VC INNER JOIN TBLALMACENES AL ON VC.IDALMACEN=AL.IDALMACEN AND VC.IDVALE=" & pId 
			mySqlDetalle = "SELECT * FROM TBLVALESDETALLE WHERE IDVALE=" & pId & " AND EXISTENCIA > 0"
		case CODIGO_VS_RECEPCION_X
			mySQL = "SELECT AL.IDDIVISION AS DIVISION_TRABAJO, VC.MOMENTO AS FECHA_TRABAJO, VC.IDVALE AS ID_FORM_ORG FROM TBLVALESCABECERA VC INNER JOIN TBLALMACENES AL ON VC.IDALMACEN=AL.IDALMACEN AND VC.IDVALE=" & pId 
			mySqlDetalle = "SELECT IDVALE,IDARTICULO,-CANTIDAD AS CANTIDAD,-EXISTENCIA AS EXISTENCIA,-SOBRANTE AS SOBRANTE, VLUPESOS, VLUDOLARES FROM TBLVALESDETALLE WHERE IDVALE=" & pId & " AND EXISTENCIA > 0"
		case CODIGO_VS_RECLASIFICACION_STOCK_X
			mySQL = "SELECT AL.IDDIVISION AS DIVISION_TRABAJO, VC.MOMENTO AS FECHA_TRABAJO, VC.IDVALE AS ID_FORM_ORG FROM TBLVALESCABECERA VC INNER JOIN TBLALMACENES AL ON VC.IDALMACEN=AL.IDALMACEN AND VC.IDVALE=" & pId 
			mySqlDetalle = "SELECT IDVALE,IDARTICULO,CANTIDAD AS CANTIDAD,EXISTENCIA AS EXISTENCIA,SOBRANTE AS SOBRANTE, VLUPESOS, VLUDOLARES FROM TBLVALESDETALLE WHERE IDVALE=" & pId & " AND EXISTENCIA <> 0"
		case CODIGO_REM_REMITO
			mySQL = "SELECT AL.IDDIVISION AS DIVISION_TRABAJO, RC.FECHA AS FECHA_TRABAJO, RC.IDREMITO AS ID_FORM_ORG FROM TBLREMCABECERA RC INNER JOIN TBLALMACENES AL ON RC.IDALMACEN=AL.IDALMACEN AND RC.IDREMITO=" & pId 
			mySqlDetalle = "SELECT * FROM TBLREMDETALLE WHERE IDREMITO=" & pId & " AND EXISTENCIA <> 0"
			precioDePIC = true
		case CODIGO_REM_ANULACION
			mySQL = "SELECT AL.IDDIVISION AS DIVISION_TRABAJO, RC.FECHA AS FECHA_TRABAJO, RC2.IDREMITO AS ID_FORM_ORG " & _
					"    FROM TBLREMCABECERA RC " & _ 
					"        INNER JOIN TBLREMCABECERA RC2 ON RC.IDREMITO<>RC2.IDREMITO AND RC.NROREMITO=RC2.NROREMITO AND RC.IDPROVEEDOR=RC2.IDPROVEEDOR " & _
					"        INNER JOIN TBLALMACENES AL ON RC.IDALMACEN=AL.IDALMACEN " & _ 
					"    AND RC.IDREMITO = " & pId
			mySqlDetalle = "SELECT IDREMITO,IDARTICULO,-CANTIDAD AS CANTIDAD,-EXISTENCIA AS EXISTENCIA,-SOBRANTE AS SOBRANTE FROM TBLREMDETALLE WHERE IDREMITO=" & pId & " AND EXISTENCIA <> 0"
			precioDePIC = true
	    case else 
			exit function
	end select
	'response.Write mySQL & "<br>"
	Call GF_BD_COMPRAS(myRS,conn,"OPEN",mySQL)
		if not myRS.eof then 
			myDivisionDeTrabajo = myRS("DIVISION_TRABAJO")
			myFechaDeTrabajo = myRS("FECHA_TRABAJO")
			MyIdFormularioDeTrabajo = myRS("ID_FORM_ORG")
		else
			exit function
			'ERROR!!!
		end if	
	Call GF_BD_COMPRAS(myRS,conn,"CLOSE",mySQL)
	'Abrir el dettalle
	'response.Write mySqlDetalle & "<br>"
	Call GF_BD_COMPRAS(myRS,conn,"OPEN",mySqlDetalle)
		while not myRS.eof
				'Se toma la existencia que tiene el formulario creado
				stockFormulario = cdbl(myRS("EXISTENCIA"))
				'response.Write "Stock Formulario = " & stockFormulario & "<br>"
				if stockFormulario <> 0 then
					'Se busca el precio actual del producto por division 
					precioActual = getUltimoPrecio(myDivisionDeTrabajo, myRS("IDARTICULO"), MONEDA_PESO, myFechaDeTrabajo)
					'response.Write "Precio Actual = " & precioActual & "<br>"
					if precioActual = 0 then
						stockActual = 0
						stockEnAlmacenes = 0
						stockEnCirculacion = 0
					else	
						'Se obtiene el stock actual del producto en el pañol
						stockEnAlmacenes = getCantidadArticulosEnDivision(myDivisionDeTrabajo, myRS("IDARTICULO"))
						'Se obtiene el stock actual del producto en circulacion
						stockEnCirculacion = getCantidadArticulosEnCirculacion(myDivisionDeTrabajo, myRS("IDARTICULO"))
						'Abajo, el stock ya se actualizo!
						'Le resto lo que trae el formulario al stock actualizado
						stockActual = cdbl(stockEnAlmacenes) + cdbl(stockEnCirculacion) - cdbl(stockFormulario)
					end if	
					'response.Write "Stock Actual = " & stockActual & "<br>"
					'response.Write "Stock Alm = " & stockEnAlmacenes & "<br>"
					'response.Write "Stock Circulante = " & stockEnCirculacion & "<br>"
					if precioDePIC then
						precioFormulario = getPrecioPicsAsociados(MyIdFormularioDeTrabajo, myRS("IDARTICULO"))
						'response.Write "Precio Form 1 = " & precioFormulario & "<br>"
					else
						precioFormulario = myRS("VLUPESOS")
						'response.Write "Precio Form 2 = " & precioFormulario & "<br>"
					end if	
					
					if isNull(precioFormulario) then precioFormulario = 0
					'Si hay stock calcular nuevo precio segun:
					'	((sA*pA)+(sF*pF)) / sA+sF
					'	Donde:
					'	sA= existencia que hay actualmente en la division
					'	sF= existencias que trae el formulario
					'	pA= precio actual del articulo en la division
					'	pF= precio que tiene el articulo en el formulario
					
					if (cdbl(stockActual) + cdbl(stockFormulario)) <> 0 then
						precioNuevoPesos = round(((cdbl(stockActual) * cLng(precioActual)) + (cdbl(stockFormulario) * cLng(precioFormulario))) / (cdbl(stockActual) + cdbl(stockFormulario)),0)
						'response.Write "Precio Nvo 1 = " & precioNuevoPesos & "<br>"
					else
						precioNuevoPesos = cLng(precioActual)
						'response.Write "Precio Nvo 2 = " & precioNuevoPesos & "<br>"
					end if	
					'Calculo del precio en dolares
					tipoCambioHoy = getTipoCambio(MONEDA_DOLAR, "")
					precioNuevoDolares = round(clng(precioNuevoPesos) / cdbl(tipoCambioHoy),0)
					
					if precioNuevoPesos > 0 then
						'Antes de actualizar obtengo el valor unitario de la última compra.
						Set rsCompra = getUltimaCompra(myDivisionDeTrabajo, myRS("IDARTICULO"), 0)
						idPic = 0
						mmtoCompra = 0
						vluPesosCompra = 0
						vluDolaresCompra = 0
						if (not rsCompra.eof) then
							idPic = rsCompra("IDCOTIZACION")
							mmtoCompra = rsCompra("MOMENTO")
							vluPesosCompra = rsCompra("VLUPESOS")
							vluDolaresCompra = rsCompra("VLUDOLARES")
						end if
						Call grabarPrecioNuevo(myDivisionDeTrabajo,myRS("IDARTICULO"),precioNuevoPesos,precioNuevoDolares,tipoCambioHoy, mmtoCompra, vluPesosCompra, vluDolaresCompra, idPic)
					end if
				end if
			myRS.movenext
		wend
	Call GF_BD_COMPRAS(myRS,conn,"CLOSE",mySqlDetalle)	
end function
'-----------------------------------------------------------------------------------------
'Funcion que obtiene la información del ultimo precio unitario registrado de un articulo en la división indicada.
function getUltimaCompra(idDivision, vecIdArticulos, idPic)
	dim strSQL, rs, conn
	
	if (idPIC = "") then idPIC=0
	
	strSQL =          "SELECT PRE.* "
	strSQL = strSQL & "FROM "
	strSQL = strSQL & "       (SELECT  MOMENTO, "
	strSQL = strSQL & "				   IDARTICULO, "
	strSQL = strSQL & "				   C.IDCOTIZACION, "
	strSQL = strSQL & "				   C.IDPROVEEDOR, "
	strSQL = strSQL & "				   C.IMPORTEPESOS IMPORTEPESOSTOTAL, "
	strSQL = strSQL & "				   C.IMPORTEDOLARES IMPORTEDOLARESTOTAL, "
	strSQL = strSQL & "				   D.IMPORTEPESOS IMPORTEPESOS, "
	strSQL = strSQL & "				   D.IMPORTEDOLARES IMPORTEDOLARES, "
	strSQL = strSQL & "				   (D.IMPORTEPESOS/CANTIDAD) VLUPESOS, "
	strSQL = strSQL & "				   (D.IMPORTEDOLARES/CANTIDAD) VLUDOLARES "
	strSQL = strSQL & "		   FROM TBLCTZCABECERA C inner join TBLCTZDETALLE D on C.IDCOTIZACION=D.IDCOTIZACION"
	strSQL = strSQL & "		   AND	 IDDIVISION = " & idDivision
	strSQL = strSQL & "        AND   IDARTICULO IN("& vecIdArticulos &") "
	strSQL = strSQL & "       ) PRE "
	strSQL = strSQL & "		inner join "
	strSQL = strSQL & "       (SELECT  MAX(MOMENTO) fecha   , "
	strSQL = strSQL & "                IDARTICULO "
	strSQL = strSQL & "       FROM     TBLCTZCABECERA C inner join TBLCTZDETALLE D on C.IDCOTIZACION=D.IDCOTIZACION"
	strSQL = strSQL & "       WHERE    C.IDCOTIZACION <> " & idPIC
	strSQL = strSQL & "       AND      CANTIDAD <> 0 "
	strSQL = strSQL & "       AND      IDARTICULO IN("& vecIdArticulos &") "
	strSQL = strSQL & "       AND      IDDIVISION = " & idDivision
	strSQL = strSQL & "       GROUP BY IDARTICULO "
	strSQL = strSQL & "       ) FECHA "
	strSQL = strSQL & "		ON  PRE.MOMENTO = FECHA.fecha "
	strSQL = strSQL & "		AND PRE.IDARTICULO = FECHA.IDARTICULO "	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getUltimaCompra = rs
	
end function
'-----------------------------------------------------------------------------------------
'Funcion que recibe una lista separada por comas de articulos y devuelve una sublista con los artículos a los que se les puede aplicar control de precios.
Function listaControlPrecio(vArticulos)
	Dim strSQL, rs, rtrn
	
	strSQL=			  " SELECT IDARTICULO "
	strSQL = strSQL & " FROM	(Select * from TBLARTICULOS where IDARTICULO in (" & vArticulos & ")) ART"
	strSQL = strSQL & "		INNER JOIN"
	strSQL = strSQL & "			TBLARTCATEGORIAS CAT ON ART.IDCATEGORIA=CAT.IDCATEGORIA"
	strSQL = strSQL & " WHERE CAT.TIPOCATEGORIA = '" & TIPO_CAT_BIENES & "'"	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	
	rtrn = "0"
	if (not rs.eof) then 
		rtrn = rs.GetString(2,,"", ",")
		rtrn = Left(rtrn, Len(rtrn) - 1)
	end if
	
	listaControlPrecio = rtrn
	
End Function
'-----------------------------------------------------------------------------------------
function getCantidadArticulosEnDivision(pIdDivision, pIdArticulo)
dim rtrn, rs, conn, strSQL
rtrn = 0
strSQL = "SELECT SUM(EXISTENCIA) AS TOTAL_PLANTA FROM TBLARTICULOSDATOS WHERE IDARTICULO = " & pIdArticulo & " AND IDALMACEN IN(SELECT IDALMACEN FROM TBLALMACENES WHERE IDDIVISION=" & pIdDivision & ")"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then rtrn = rs("TOTAL_PLANTA")
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
getCantidadArticulosEnDivision = rtrn
end function
'-----------------------------------------------------------------------------------------
Function getCantidadArticulosEnAlmacen(pIdAlmacen,pIdArticulo)
	dim rtrn, rs, conn, strSQL
	rtrn = 0
	strSQL = "SELECT SUM(EXISTENCIA+SOBRANTE) AS TOTAL_ALMACEN FROM TBLARTICULOSDATOS WHERE IDARTICULO = " & pIdArticulo & " AND IDALMACEN = " & pIdAlmacen
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then 
		if (not isnull(rs("TOTAL_ALMACEN"))) then rtrn = rs("TOTAL_ALMACEN")
	end if
	
	getCantidadArticulosEnAlmacen = rtrn
End Function 
'-----------------------------------------------------------------------------------------
function getCantidadArticulosEnCirculacion(pIdDivision, pIdArticulo)
dim rtrn, rs, conn, strSQL
rtrn = 0
strSQL =	"SELECT SUM(SALDO1) AS TOTAL_CIRCULACION FROM " & _
            "	( " & _
            "		SELECT C.CDVALE, CASE(C.CDVALE) WHEN '" & CODIGO_VS_DEVOLUCION & "' THEN SUM(-D.EXISTENCIA) WHEN '" & CODIGO_VS_AJUSTE_VALE & "' THEN SUM(-D.EXISTENCIA) WHEN '" & CODIGO_VS_PRESTAMO & "' THEN SUM(D.EXISTENCIA) END " & chr(34) & "SALDO1" & chr(34) & _ 
            "			FROM TBLVALESDETALLE D " & _ 
            "				INNER JOIN TBLVALESCABECERA C ON C.IDVALE=D.IDVALE " & _ 
			"		    WHERE D.IDARTICULO=" & pIdArticulo & " AND D.EXISTENCIA>0 AND C.CDVALE IN ('" & CODIGO_VS_DEVOLUCION & "','" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_AJUSTE_VALE & "') AND C.IDALMACEN IN (SELECT IDALMACEN FROM TBLALMACENES WHERE IDDIVISION=" & pIdDivision & ") and C.ESTADO= " & ESTADO_ACTIVO & _ 
            "	    GROUP BY C.CDVALE " & _
			"	) T1 " 
			'Response.Write "<br>SQL - " & strSQL & "<hr>"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not IsNull(rs("TOTAL_CIRCULACION")) then rtrn = rs("TOTAL_CIRCULACION")
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
getCantidadArticulosEnCirculacion = rtrn
end function
'-----------------------------------------------------------------------------------------
function getPrecioPicsAsociados(pIdRemito, pIdArticulo)
dim rtrn, rs, conn, strSQL
dim myPrecioUnitario, myCantidadAsignada
dim myAcumuladoImportes, myAcumuladoCantidad
myAcumuladoImportes = 0
myAcumuladoCantidad = 0
rtrn = 0
strSQL =	"SELECT RP.CANTIDAD AS CANTIDAD_ASIGNADA, PC.TIPOCAMBIO,PD.CANTIDAD AS CANTIDAD_PIC, PD.IMPORTEPESOS AS IMPORTE_PIC " & _
			"	FROM TBLREMPIC RP " & _
			"	INNER JOIN TBLCTZCABECERA PC ON RP.IDPIC=PC.IDCOTIZACION " & _
			"	INNER JOIN TBLCTZDETALLE PD ON PC.IDCOTIZACION=PD.IDCOTIZACION AND PD.IDARTICULO=RP.IDARTICULO " & _
			"	WHERE RP.IDREMITO=" & pIdRemito & " AND RP.IDARTICULO=" & pIdArticulo
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		myPrecioUnitario = CLng(rs("IMPORTE_PIC")) / CDbl(rs("CANTIDAD_PIC"))
		myCantidadAsignada = CDbl(rs("CANTIDAD_ASIGNADA"))
		myAcumuladoImportes = myAcumuladoImportes + (myCantidadAsignada * myPrecioUnitario)
		myAcumuladoCantidad = myAcumuladoCantidad + myCantidadAsignada
	rs.movenext
wend
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
if myAcumuladoCantidad <> 0 then rtrn = clng(round(clng(myAcumuladoImportes) / cdbl(myAcumuladoCantidad),0))
getPrecioPicsAsociados = rtrn
end function
'-----------------------------------------------------------------------------------------
function getUltimoPrecio(pDivision, pIdArticulo, pMoneda, pFecha)
dim myDic, rtrn
rtrn = 0
set myDic = obtenerPreciosArticulos(pIdArticulo, pDivision, pMoneda, pFecha)
if myDic.Exists(cstr(pIdArticulo)) then 
	'Response.Write "ENTRO(" & myDic(cstr(pIdArticulo))  & ")"
	rtrn = myDic(cstr(pIdArticulo))
end if	
getUltimoPrecio = rtrn
end function
'-----------------------------------------------------------------------------------------
sub grabarPrecioNuevo(pDivision, pIdArticulo , pPrecioNuevoPesos, pPrecioNuevoDolares, pTipoCambio, pMmtoCompra, pVluPesosCompra, pVluDolaresCompra, pIDPic)
	dim rs, conn, strSQL
	strSQL =" INSERT INTO TBLARTICULOSPRECIOS VALUES(" & session("MmtoSistema")  & "," & pDivision & "," & pIdArticulo & "," & pPrecioNuevoPesos & "," & pPrecioNuevoDolares & "," & pTipoCambio & ", " & pMmtoCompra & ", " & pVluPesosCompra & ", " & pVluDolaresCompra & ", " & pIDPic & ")"
	'Response.Write "<hr>Nuevo Precio" & strSQL
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------
sub grabarPreciosVigentesPorArticulo(pIdVale)
dim rs, conn, strSQL
dim myCdVale, myMmtoVale
'Esta funcion le asigna a cada vale el precio vigente para cada articulo. No hace calculo de precios!
'CASOS
'-->VMS: Precio de lista
'-->VMP: Sin Precio (0)
'-->AJP: Sin Precio (0)
'-->AJS: Precio de lista
'-->AJU: Precio de lista
'-->VMT: Precio de lista
'-->VMR: Toma el precio del VMT Original
'-->VME: Sin Precio (0)
'-->VMD: Sin Precio (0)
'-->AJT: Toma el precio del VMT Original
'-->XMS: Toma el precio del VMS Original
'-->XMT: Toma el precio del VMT Original
'-->XMR: Toma el precio del VMR Original
'-->XJU: Toma el precio del AJU Original
'-->XJS: Toma el precio del AJS Original
'-->XJT: Toma el precio del AJT Original

myCdVale = ""
'Obtener codigo de vale y momento de grabacion
strSQL ="SELECT CDVALE, MOMENTO FROM TBLVALESCABECERA WHERE IDVALE= " & pIdVale 
Call GF_BD_COMPRAS(rs,conn,"OPEN",strSQL)
	if not rs.eof then 
		myCdVale = rs("CDVALE")
		myMmtoVale = rs("MOMENTO")
	end if	
Call GF_BD_COMPRAS(rs,conn,"CLOSE",strSQL)
'Buscar para cada articulo de vale el precio correspondiente de acuerdo a la fecha.
select case ucase(myCdVale)
	case CODIGO_VS_SALIDA, CODIGO_VS_AJUSTE_STOCK, CODIGO_VS_AJUSTE_VALE, CODIGO_VS_TRANSFERENCIA
		strSQL=	"SELECT ART.IDARTICULO, ART.VLUPESOS, ART.VLUDOLARES FROM TBLARTICULOSPRECIOS ART INNER JOIN " & _
				"    ( " & _
				"    SELECT VD.IDARTICULO, AL.IDDIVISION, MAX(MMTOPRECIO) AS MMTOPRECIO FROM TBLVALESCABECERA VC " & _ 
				"            INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _ 
				"            INNER JOIN TBLALMACENES AL ON VC.IDALMACEN=AL.IDALMACEN " & _
				"            INNER JOIN TBLARTICULOSPRECIOS PRE ON VD.IDARTICULO=PRE.IDARTICULO AND AL.IDDIVISION=PRE.IDDIVISION AND PRE.MMTOPRECIO<= " & myMmtoVale & _
				"        WHERE VC.IDVALE= " & pIdVale & _
				"    GROUP BY VD.IDARTICULO, AL.IDDIVISION " & _
				"    ) TF " & _
				"    ON ART.IDARTICULO=TF.IDARTICULO AND TF.MMTOPRECIO=ART.MMTOPRECIO AND ART.IDDIVISION=TF.IDDIVISION"
	case CODIGO_VS_RECLASIFICACION_STOCK
			strSQL=	"SELECT DET.IDARTICULO, ARTD.VLUPESOS, ARTD.VLUDOLARES " & _
					"	FROM " & _
					"		(SELECT * FROM TBLVALESDETALLE WHERE IDVALE= " & pIdVale & ") DET " & _
					"	INNER JOIN TBLARTICULOSPRECIOS ARTD " & _
					"		ON ARTD.IDARTICULO = (SELECT IDARTICULO FROM TBLVALESDETALLE WHERE IDVALE = " & pIdVale & " AND CANTIDAD < 0) " & _
					"			AND MMTOPRECIO=" & _
					"				( " & _
					"				SELECT MAX(MMTOPRECIO) " & _
					"					FROM TBLARTICULOSPRECIOS " & _
					"						WHERE IDARTICULO = ARTD.IDARTICULO " & _
					")"
	case CODIGO_VS_RECEPCION, CODIGO_VS_AJUSTE_TRANSFERENCIA
		strSQL=	"SELECT VD.IDARTICULO, VD1.VLUPESOS, VD1.VLUDOLARES FROM TBLVALESCABECERA VC " & _
				"    INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE=VD.IDVALE " & _
				"    INNER JOIN TBLVALESCABECERA VC1 ON VC.PARTIDAPENDIENTE=VC1.PARTIDAPENDIENTE " & _
				"    INNER JOIN TBLVALESDETALLE VD1 ON VC1.IDVALE=VD1.IDVALE AND VD.IDARTICULO=VD1.IDARTICULO " & _
				"    WHERE VC.IDVALE=" & pIdVale & " AND VC1.CDVALE IN ('" & CODIGO_VS_TRANSFERENCIA & "')	"
	case CODIGO_VS_AJUSTE_TRANSFERENCIA_X, CODIGO_VS_AJUSTE_STOCK_X, CODIGO_VS_AJUSTE_VALE_X, CODIGO_VS_RECEPCION_X, CODIGO_VS_TRANSFERENCIA_X, CODIGO_VS_SALIDA_X,CODIGO_VS_RECLASIFICACION_STOCK_X
		strSQL=	"SELECT VD.IDARTICULO, VD1.VLUPESOS, VD1.VLUDOLARES FROM TBLVALESRELACIONES VR " & _
				"    INNER JOIN TBLVALESDETALLE VD ON VR.IDVALE_2=VD.IDVALE " & _				
				"    INNER JOIN TBLVALESDETALLE VD1 ON VR.IDVALE_1=VD1.IDVALE AND VD1.IDARTICULO=VD.IDARTICULO" & _
				"    WHERE VD.IDVALE=" & pIdVale
	case else
		exit sub
end select
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof 
		if not isNull(rs("VLUPESOS")) then
			call setPreciosVigentesPorArticulo(pIdVale, rs("IDARTICULO"), rs("VLUPESOS"), rs("VLUDOLARES"))
		end if	
	rs.movenext
wend	
end sub
'-----------------------------------------------------------------------------------------
sub setPreciosVigentesPorArticulo(pIdVale, pIdArticulo, pVluPesos, pVluDolares)
dim rs, conn, strSQL
strSQL = "UPDATE TBLVALESDETALLE SET VLUPESOS=" & pVluPesos & ", VLUDOLARES=" & pVluDolares & " WHERE IDVALE=" & pIdVale & " AND IDARTICULO=" & pIdArticulo
call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end sub
'---------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	01/12/2010
' Objetivo:	
'			Inicializa los permisos del usuario sobre el recurso indicado.
' Parametros:
'			pResource [int] 		Recurso sobre el que se va a operar (ver constantes RES_XXX).
' Devuelve:
'			-
Function initAccessInfo(pResource)
	Call initSystemAccessInfo(SEC_SYS_UNASIGNED, pResource)
End Function
'---------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	23/01/2015
' Objetivo:	
'			Inicializa los permisos del usuario sobre el recurso indicado del sistema indicado.
' Parametros:
'			pSystem      [int] 		ID Del sistema al cual pertenece el recurso (ver constantes SEC_SYS_XXX).
'			pResource    [int] 		Recurso sobre el que se va a operar (ver constantes RES_XXX).
' Devuelve:
'			-
Function initSystemAccessInfo(pSystem, pResource)
	Dim Sql, rs, cn

	'Obtengo la lista de divisiones sobre l<b style="color: rgb(0, 0, 0); font-family: &quot;Times New Roman&quot;; font-size: medium; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; text-decoration-style: initial; text-decoration-color: initial;"><b>MACEN from TBLAL</b></b>os que tiene permisos.
	sql = "SELECT IDDIVISION, PERMISO FROM TBLUSUARIOPERMISOS "
	sql = sql & " WHERE IDSISTEMA=" & pSystem & " and IDRECURSO = " & pResource & " AND CDUSUARIO = '" & session("usuario") & "' "
	sql = sql & " AND PERMISO NOT IN (" & SEC_X & ") "
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", sql)
	if (not rs.eof) then
		While (not rs.eof)
			ccValorListaDivision(cInt(rs("IDDIVISION"))) = cInt(rs("PERMISO"))
			rs.movenext
		Wend
		ccListaDivisionAdmin = getListaCargosAdmin()
	else
		response.redirect SITE_ROOT & "comprasAccesoDenegado.asp"
	end if
End Function
'---------------------------------------------------------------------------------------------
function isInList(pIdDivision, pTipo)
	dim rtrn, i, aux
	rtrn = false
	if (pIdDivision <> SIN_DIVISION) then
		rtrn = (ccValorListaDivision(pIdDivision) = pTipo)
	else
		for i = 1 to 4
			aux = (ccValorListaDivision(i) = pTipo)
			if (aux) then rtrn = aux
		next
	end if
	isInList = rtrn
end function
'--------------------------------------------------------------------------------------
Function isAdmin(pIdDivision)
	'Si es administrador tiene acceso sobre todo un Centro Costos por lo menos
	isAdmin = isInList(pIdDivision, SEC_A)
End Function
'--------------------------------------------------------------------------------------
Function isAuditor(pIdDivision)
	'Si es auditor solo tiene permisos de lectura.	
	isAuditor = isInList(pIdDivision, SEC_Y)
End Function
'--------------------------------------------------------------------------------------
Function isUser(pIdDivision)
	isUser = isInList(pIdDivision, SEC_U)	
End Function
'--------------------------------------------------------------------------------------
' Autor: 	Guido Fonticelli - GFG
' Fecha: 	29/12/2011
' Objetivo:	
'			Reemplazar los caracteres que generan errores en el nombre del archivo al descargarlos
' Parametros:
'			[str]	pName
' Devuelve:
'			por referencia, el nombre formateado
'--------------------------------------------------------------------------------------------------
Function FileName2DbName(byref pName)
	Dim invalidChars(1,1),j
	'reemplazo caracteres invalidos que no permiten bajar bien el archivo
	if (len(pName)>0) then
		'posicion 0 -> caracter a reemplazar
		'posicion 1 -> por cual se reemplazara
		
		invalidChars(0,0) = " " ' Reemplaza los espacios por un "_"
		invalidChars(0,1) = "_" '
		
		invalidChars(1,0) = "." ' Reemplaza los puntos por un "-"
		invalidChars(1,1) = "-" '
				
		for j = 0 to ubound(invalidChars)
			pName = replace(pName,invalidChars(j,0),invalidChars(j,1))
		next
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function getDSDocumentoFirmar(tipo)
	Dim ret
	ret = "PROCEDIMIENTO DESCONOCIDO"
	Select case tipo
		case AUTH_TYPE_PCP:
			ret = "Analisis Comparativo (PCP)"
		case AUTH_TYPE_PIC:
			ret = "Pedido Interno de Compras (PIC)"
		case AUTH_TYPE_CEC:
			ret = "Cbte. Electronico de Cumplimiento (CEC)"
		case AUTH_TYPE_AIC:
			ret = "Ajuste de Pedido Interno de Compras (AIC)"
	    case AUTH_TYPE_AEC:
			ret = "Ajuste de Conf. Electr. de Cumplimiento (AEC)"
		case AUTH_TYPE_AFE:
			ret = "Autorizaciones para Gastos (AFE)"
		case AUTH_TYPE_AJS:
			ret = "Ajuste de Stock (AJS)"
		case AUTH_TYPE_XJS:
			ret = "Anulacion Ajuste de Stock (XJS)"
		case AUTH_TYPE_VRS:
			ret = "Reclasificacion de Stock (VRS)"
		case AUTH_TYPE_XRS:
			ret = "Anulacion Reclasificacion de Stock (XRS)"
		case AUTH_TYPE_CTC:
			ret = "Ajuste a Contrato/Servicio (CTC)"
		case AUTH_TYPE_AJD:
			ret = "Ajustes de Draft Survey(AJD)"	
		case AUTH_TYPE_AJC:
			ret = "Ajustes de Calidad(AJD)"	
		case AUTH_TYPE_AJM:
			ret = "Ajustes de Manipuleo(AJM)"		
		case AUTH_TYPE_CCN:
			ret = "Cierre de Almacen(CCN)"
        case AUTH_TYPE_AJV
            ret = "Ajuste Merma volatil (AJV)"
        case AUTH_TYPE_APP
            ret = "Ajuste de Partida Presupuestaria(APP)"        
	End Select
	getDSDocumentoFirmar = GF_TRADUCIR(ret)
End Function
'----------------------------------------------------------------------------------------------------------------------
Function getListMail(pIdDivision, pCdLista)
	Dim strSQL, rs	
	strSQL = "				SELECT EMAIL "
	strSQL = strSQL & "		FROM TBLMAILLSTCABECERA A "
	strSQL = strSQL & "			INNER JOIN TBLMAILLSTSDETALLE B "
	strSQL = strSQL & "				ON  A.IDLISTA = B.IDLISTA "
	strSQL = strSQL & "		WHERE A.CDLISTA = '" & Trim(UCase(pCdLista)) & "' AND A.IDDIVISION = " & pIdDivision	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getListMail = rs
End Function	
'----------------------------------------------------------------------------------------------------------------------	
' Función:	  
'				getMailCoordinadorPto
' Autor: 	
'				CNA - Ajaya Nahuel
' Fecha: 		
'				05/03/2013
' Objetivo:		
'				Obtiene el mail de todos los Coordinadores de los Puertos
' Parametros:	-
' Devuelve:		RecordSet
'----------------------------------------------------------------------------------------------
Function getMailCoordinadorPto()
	Dim strSQL, strCoord
	strSQL = "SELECT DISTINCT CDUSUARIO FROM TBLROLESUSUARIOS WHERE IDROL = " & FIRMA_ROL_SUP_PUERTO	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while(not rs.EoF)
		strCoord = strCoord & getUserMail(rs("CDUSUARIO")) & ";"		
		rs.MoveNext
	wend	
	getMailCoordinadorPto = strCoord
End Function	
'---------------------------------------------------------------------------------------------
'Función:	
'				getDocumentoFirmar	
' Autor: 	
'				CNA - Ajaya Nahuel
' Fecha: 	
'				30/10/2013
' Objetivo:
'				Devuelve todos los tipos de documentos a firmar
' Parametros:
'				-
' Devuelve:
'				Array con los Tipos de Documentos 
'--------------------------------------------------------------------------------------------------
Function getDocumentoFirmar()
	Dim vTipoDocumento()
	Redim vTipoDocumento(13)
	vTipoDocumento(0) = AUTH_TYPE_PCP
	vTipoDocumento(1) = AUTH_TYPE_PIC
	vTipoDocumento(2) = AUTH_TYPE_CEC
	vTipoDocumento(3) = AUTH_TYPE_AIC
	vTipoDocumento(4) = AUTH_TYPE_AEC
	vTipoDocumento(5) = AUTH_TYPE_AFE	
	vTipoDocumento(6) = AUTH_TYPE_CTC	
	vTipoDocumento(7)= AUTH_TYPE_CCN	
    vTipoDocumento(8) = AUTH_TYPE_APP		
	vTipoDocumento(9) = AUTH_TYPE_VRS	
	vTipoDocumento(10) = AUTH_TYPE_AJD	
	vTipoDocumento(11) = AUTH_TYPE_AJV
	vTipoDocumento(12) = AUTH_TYPE_AJC
	vTipoDocumento(13) = AUTH_TYPE_AJM	
	getDocumentoFirmar = vTipoDocumento
End Function
'---------------------------------------------------------------------------------------------
function getValorNorma(cdNorma)
	dim strSQL, rs, conn, rtrn
	rtrn = ""
	strSQL="Select * from TBLNORMASAUDITORIA where CDNORMA='" & cdNorma & "'"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = rs("VALOR")
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getValorNorma = rtrn
end function
'---------------------------------------------------------------------------------------------
function getUnidadNorma(cdNorma)
	dim strSQL, rs, conn, rtrn
	rtrn = ""
	strSQL="Select * from TBLNORMASAUDITORIA where CDNORMA='" & cdNorma & "'"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = rs("UNIDAD")
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getUnidadNorma = rtrn
end function
'-------------------------------------------------------------------------------------------------
function IsToepfer(p_krempresa)
	if p_krempresa = CD_TOEPFER  then
		isToepfer = true
	else
		isToepfer = false
	end if
end function
'------------------------------------------------------------------------------------------------------------------
Sub MyDelay(NumberOfSeconds)
    Dim DateTimeResume
      DateTimeResume= DateAdd("s", NumberOfSeconds, Now())
      Do Until (Now() > DateTimeResume)
      Loop
End Sub
'---------------------------------------------------------------------------------------------
'Devuelve la descripcion del estado de la provision
Function getEstadoProvisionesCancelacion(p_CdEstado)
    Dim rtrn
    select case Cstr(p_CdEstado)
        case PROVISCIONES_ESTADO_GENERADO
            rtrn = "GENERADO" 
        case PROVISCIONES_ESTADO_PENDIENTE
            rtrn = "PENDIENTE"
        case PROVISCIONES_ESTADO_AUTORIZADO
            rtrn = "AUTORIZADO" 
        case PROVISCIONES_ESTADO_APLICADO
            rtrn = "APLICADO" 
        case PROVISCIONES_ESTADO_ERROR
            rtrn = "ERROR" 
    end select
    getEstadoProvisionesCancelacion = rtrn
End Function
%>