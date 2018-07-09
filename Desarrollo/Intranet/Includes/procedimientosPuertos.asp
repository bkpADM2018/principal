
<%
Dim g_cdPesoHect, g_cdProteina, g_cdTemperatura, g_cdHumedad
Dim gDicConv

Const cnstCdCliente = 44 'MateriaPampa
Const cnstCdProducto = 15 'trigo pan

'Constantes de Puertos
Const TERMINAL_ARROYO		= "ARROYO"
Const TERMINAL_TRANSITO		= "TRANSITO"
Const TERMINAL_PIEDRABUENA	= "PIEDRABUENA"

'Constantes para la aceptacion en calada.
Const ACEPTACION_SIN_AUTORIZACION = 10  'Codigo aceptacion - SIN AUTORIZACION
Const ACEPTACION_CONFORME = 1           'Es conforme
Const ACEPTACION_COND_CAMARA = 2        'Condicion camara
Const ACEPTACION_REBAJA_CONVENIDA = 3   'Rebaja convenida
Const ACEPTACION_ANALISIS = 4           'Analisis
Const ACEPTACION_RECHAZO = 9            'Rechazar - pasa a ser Estado RECHAZADO
Const ACEPTACION_AUT_ENTREGADOR = 15    'Autoriza Entregador

Const OPERATIVOS_ESTADO_CARGADO   = 1
Const OPERATIVOS_ESTADO_ARRIBADO  = 2
Const OPERATIVOS_ESTADO_INICIADO  = 3
Const OPERATIVOS_ESTADO_TERMINADO = 4
Const OPERATIVOS_ESTADO_BAJA      = 5

'------------ PARAMETROS ------------'
const ACCION_BUSCAR_PARAMETRO   		= 1
const ACCION_MODIFICAR_PARAMETRO 		= 2
const ACCION_AGREGAR_PARAMETRO   		= 3
const ACCION_ELIMINAR_PARAMETRO  		= 4
const ACCION_MODIFICAR_PARAMETRO_EXTRA 	= 5
const ACCION_COMPROBAR_PARAMETRO   		= 6
const ACCION_AGREGAR_LOG  		 		= 8
const ACCION_MODIFICAR_LOG   			= 9

const PARAMETRO_NO_EXISTENTE 	  		= 7

'----------PERMISOS PARAMETROS-------'
'-- Aplicativo de Parametros
const TASK_PARAM_ADMIN	= 702 
const TASK_PARAM_USER	= 701
const TASK_PARAM_AUDIT	= 700
'-- ABM de Productos --
const TASK_PRODUCT_USER	=  94
'-- Modificacion Historica
const TASK_MH_CTA_PTE =  707
const TASK_MH_CALIDAD =  708

const NO_TIENE_PERMISO  = 0

'----------------------------
const PARAMETRO_EDITABLE   	= "S"
const PARAMETRO_NO_EDITABLE = "N"
'----------------------------
const RES_PUERTO_PARAMETROS = 1
'-----------------------------

Const TIPO_HISTORIAL = "H"
Const TIPO_DIARIO = "D"

'----------BALANZA DE CAMIONES ------------'
const TASK_BZA_CAM_STK_USER		= 703
const TASK_BZA_CAM_STK_USRPRO 	= 704
const TASK_BZA_CAM_STK_AUDIT	= 705
const TASK_BZA_CAM_STK_ADMIN	= 706

Const BZA_CAM_ESTADO_TODOS		= -1
Const BZA_CAM_ESTADO_EN_CURSO   = 5
Const BZA_CAM_ESTADO_FINALIZADO = 0
Const BZA_CAM_ESTADO_CANCELADO  = 99

Const BZA_CAM_TIPO_CTRL_TODOS	= "T"
Const BZA_CAM_TIPO_CTRL_MANUAL  = "M"
Const BZA_CAM_TIPO_CTRL_AUTOM   = "A"

Const PARAM_CANT_BZA_CAMIONES   = "CANTBZACAMIONES" 
Const PARAM_CANT_BZA_CONTROLES  = "CANTBZACONTROLES"
Const PARAM_TIPO_CTRL_BZA_CAMIONES = "TIPOCTRLBZACONTROLES"

Const BZA_CAM_MAX   = 5 'Maximo nro de balanzas disponibles para activar

'----------ARCHIVOS DE LOGS-------'
Const NOMBRE_RUTA_PARAMETRO_TRA = "TRA-PARAM-"
Const NOMBRE_RUTA_PARAMETRO_ARR = "ARR-PARAM-"
Const NOMBRE_RUTA_PARAMETRO_LPB = "LPB-PARAM-"

'----------CODIGO DE AJUSTES -------'
Const AJUSTE_DRAFT_SURVEY = "AJD"
Const AJUSTE_CALIDAD	  = "AJC"
Const AJUSTE_MANIPULEO    = "AJM"
Const AJUSTE_MERMA_VOLATIL = "AJV"
'----------ESTADOS DE AJUSTES -------'
Const AJUSTE_ESTADO_NOAUTORIZADO = 1 ' Sin ninguna autorizacion (pendiente)
Const AJUSTE_ESTADO_AUTORIZADO   = 5 ' Con todas las autorizaciones
Const AJUSTE_ESTADO_CANCELADO    = 6 ' Cancelado
'----------SECUENCIA DE LAS FIRMAS DE AJUSTES -------'
Const AJS_FIRMA_GERENTE_PUERTOS = 0
Const AJS_FIRMA_CONTROLLER	= 1
Const AJS_FIRMA_DIRECTOR = 2
'----------GRADO DE LA CAMARA -------'
CONST GRADO_CAMARA_1 = 1
CONST GRADO_CAMARA_2 = 2
CONST GRADO_CAMARA_3 = 3
CONST GRADO_CAMARA_FE = 4
'------------------------------------'
Const LISTA_DRAFT_AUTORIZADO		 = "LISTA-AJD"
'------------------------------------'
Const TIPO_TRANSPORTE_CAMION = 1 'Camiones
Const TIPO_TRANSPORTE_VAGON  = 2 'Vagones
Const TIPO_TRANSPORTE_CAMVAG = 3 'Camiones y/o vagones
'------------------------------------'
Const PESADA_BRUTO = 1
Const PESADA_TARA = 2
'------------ UNIDADES DE PESO ------------------------'
Const TIPO_PESO_KILO = 1
Const TIPO_PESO_TONELADA = 2
Const TIPO_PESO_BUSHEL = 3
'------------ CLAVES DE CONVERSION DE CODIGOS A TERCEROS ------------------'
Const CONV_KEY_PUERTO = "PUERTO"
Const CONV_KEY_PLANTA = "PLANTA"
Const CONV_KEY_PRODUCTO = "PRODUCTO"
Const CONV_KEY_PROVINCIA = "PROVINCIA"
Const CONV_KEY_SERVICIO = "SERVICIO"
Const CONV_KEY_CATEG_IVA = "CATEG.IVA"
Const CONV_KEY_AFIP = "AFIP"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsEstado(p_cdEstado)
	dim strSql, l_rsEstado	
	
	strSql = ""
	strSql = strSql & "SELECT dsEstado "
	strSql = strSql & "FROM   Estados "
	strSql = strSql & "WHERE  cdEstado = " & p_cdEstado
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsEstado, "OPEN",strSql 
	if not l_rsEstado.eof then
		getDsEstado = l_rsEstado("dsEstado")
	else
		getDsEstado = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function isEstadoTerminal(p_circuito, p_cdEstado)
	dim strSql, l_rsEstado, ret
	
	if (session("isEstadoTerminal_" & p_circuito & "_" & p_cdEstado) = "") then	
	    strSql = "SELECT cdEstado "
	    strSql = strSql & " FROM   ESTADOSTERMINALES "
	    strSql = strSql & " WHERE  cdEstado = " & p_cdEstado
	    strSql = strSql & " AND  cdTipoCAmion = " & p_circuito
    	
	    GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsEstado, "OPEN",strSql 
	    ret = false	   
	    if (not l_rsEstado.eof) then ret = true
	    session("isEstadoTerminal_" & p_circuito & "_" & p_cdEstado) = ret
    else
        ret = session("isEstadoTerminal_" & p_circuito & "_" & p_cdEstado)
    end if    
    isEstadoTerminal = ret
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsCliente(p_cdCliente)
		
	dim strSql, l_rsCliente	
	
	getDsCliente = " "
	
	if (p_cdCliente <> "") then
	    strSql = ""
	    strSql = strSql & "SELECT dsCliente "
	    strSql = strSql & "FROM   Clientes "
	    strSql = strSql & "WHERE  cdCliente = " & p_cdCliente	
	    GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsCliente, "OPEN",strSql 
	    if not l_rsCliente.eof then getDsCliente = l_rsCliente("dsCliente")
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsClienteByCUIT(p_cuitCliente)
		
	dim strSql, l_rsCliente	
	
	getDsClienteByCUIT = " "
	
	if (p_cuitCliente <> "") then	    
	    if (session("getDsClienteByCUIT_" & p_cuitCliente) = "") then
	        strSql = ""
	        strSql = strSql & "SELECT dsCliente "
	        strSql = strSql & "FROM   Clientes "
	        strSql = strSql & "WHERE  nucuit = '" & Trim(p_cuitCliente) & "' "
            strSql = strSql & "order by cdcliente"                    
	        Call executeQueryDb(session("TERMINAL_ACTUAL"), l_rsCliente, "OPEN", strSql) 
	        if not l_rsCliente.eof then 
	            getDsClienteByCUIT = Trim(l_rsCliente("dsCliente"))
	            session("getDsClienteByCUIT_" & p_cuitCliente) = Trim(l_rsCliente("dsCliente"))
            end if
        else
            getDsClienteByCUIT = session("getDsClienteByCUIT_" & p_cuitCliente)
        end if            
	end if	
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getCUITCliente(p_cdCliente)
		
	dim strSql, l_rsCliente	
	
	getCUITCliente = 0
	
	if (p_cdCliente <> "") then
	    strSql = ""
	    strSql = strSql & "SELECT NUCUIT "
	    strSql = strSql & "FROM   Clientes "
	    strSql = strSql & "WHERE  cdCliente = " & p_cdCliente	
	    GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsCliente, "OPEN",strSql 
	    if not l_rsCliente.eof then getCUITCliente = Trim(l_rsCliente("NUCUIT"))
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Devuelve el codigo de corredor del puerto por su Cuit
Function getCdCorredorByCuit(p_Pto,p_cuitCorredor)
    Dim strSQL,rtrn
    rtrn = 0
    strSQL = "SELECT CDCORREDOR FROM CORREDORES WHERE RTRIM(NUCUIT) = '"& Trim(p_cuitCorredor) &"'"
    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then rtrn = rs("CDCORREDOR")
    getCdCorredorByCuit = rtrn
End function
'-------------------------------------------------------------------------------------------------------------------------
'Devuelve el CUIT de corredor del puerto por su Codigo
Function getCuitCorredorByCd(p_Pto,p_cdCorredor)
    Dim strSQL,rtrn
    rtrn = 0
    strSQL = "SELECT NUCUIT FROM CORREDORES WHERE CDCORREDOR = "& p_cdCorredor
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then rtrn = Trim(rs("NUCUIT"))
    getCuitCorredorByCd = rtrn
End function
'-------------------------------------------------------------------------------------------------------------------------
'Devuelve el codigo de Vendedor del puerto por su Cuit
Function getCuitVendedorByCd(p_Pto,p_cuitVendedor)
    Dim strSQL,rtrn
    rtrn = 0
    strSQL = "SELECT NUDOCUMENTO FROM VENDEDORES WHERE CDVENDEDOR = " & p_cuitVendedor
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then rtrn = Trim(rs("NUDOCUMENTO"))
    getCuitVendedorByCd = rtrn
End function
'-------------------------------------------------------------------------------------------------------------------------
'Devuelve el codigo de Vendedor del puerto por su Cuit
Function getCdVendedorByCuit(p_Pto,p_cuitVendedor)
    Dim strSQL,rtrn
    rtrn = 0
    strSQL = "SELECT CDVENDEDOR FROM VENDEDORES WHERE RTRIM(NUDOCUMENTO) = '"& Trim(p_cuitVendedor) &"' order by CDVENDEDOR"
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then rtrn = CLng(rs("CDVENDEDOR"))
    getCdVendedorByCuit = rtrn
End function
'-------------------------------------------------------------------------------------------------------------------------
function getDsProducto(p_cdProducto)
	dim strSql, l_rsProducto	
	strSql = ""
	strSql = strSql & "SELECT dsProducto "
	strSql = strSql & "FROM   Productos "
	strSql = strSql & "WHERE  cdProducto = " & p_cdProducto
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsProducto, "OPEN",strSql 
	if not l_rsProducto.eof then
		getDsProducto = Trim(l_rsProducto("dsProducto"))
	else
		getDsProducto = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsAceptacion(p_cdAceptacion)
	dim strSql, l_rsAceptacion
	strSql = ""
	strSql = strSql & "SELECT dsAceptacion "
	strSql = strSql & "FROM   ACEPTACIONCALIDAD "
	strSql = strSql & "WHERE  cdAceptacion = " & p_cdAceptacion
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsAceptacion, "OPEN",strSql 
	if not l_rsAceptacion.eof then
		getDsAceptacion = Trim(l_rsAceptacion("dsAceptacion"))
	else
		getDsAceptacion = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsTransportista(p_cdTransportista)
	dim strSql, l_rsTransportista	

	strSql = ""
	strSql = strSql & "SELECT dsTransportista "
	strSql = strSql & "FROM   Transportistas "
	strSql = strSql & "WHERE  cdTransportista = " & p_cdTransportista
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsTransportista, "OPEN",strSql 
	if not l_rsTransportista.eof then
		getDsTransportista = l_rsTransportista("dsTransportista")
	else
		getDsTransportista = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsCorredor(p_cdCorredor)
	dim strSql, rsCorredor	

    getDsCorredor = " "

    if (p_cdCorredor <> "") then
	    strSql = strSql & "SELECT dsCorredor "
	    strSql = strSql & "FROM   CORREDORES "
	    strSql = strSql & "WHERE  CDCORREDOR = " & p_cdCorredor
    	
	    GF_BD_Puertos session("TERMINAL_ACTUAL"), rsCorredor, "OPEN",strSql 
	    if not rsCorredor.eof then getDsCorredor = rsCorredor("dsCorredor")
    end if
    
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsEntregador(p_cdEntregador)
	dim strSql, rsEntregador

	strSql = strSql & "SELECT DSENTREGADOR "
	strSql = strSql & "FROM   ENTREGADORES "
	strSql = strSql & "WHERE  CDENTREGADOR = " & p_cdEntregador
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), rsEntregador, "OPEN",strSql 
	if not rsEntregador.eof then
		getDsEntregador = rsEntregador("DSENTREGADOR")
	else
		getDsEntregador = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsVendedor(p_cdVendedor)
	dim strSql, rsVendedor
	strSql = strSql & "SELECT DSVENDEDOR "
	strSql = strSql & "FROM   VENDEDORES "
	strSql = strSql & "WHERE  CDVENDEDOR = " & p_cdVendedor
	GF_BD_Puertos session("TERMINAL_ACTUAL"), rsVendedor, "OPEN",strSql 
	if not rsVendedor.eof then
		getDsVendedor = rsVendedor("DSVENDEDOR")
	else
		getDsVendedor = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsComprador(p_cdComprador)
	dim strSql, rsComprador
	strSql = strSql & "SELECT DSCOMPRADOR "
	strSql = strSql & "FROM   COMPRADORES "
	strSql = strSql & "WHERE  CDCOMPRADOR = " & p_cdComprador
	GF_BD_Puertos session("TERMINAL_ACTUAL"), rsComprador, "OPEN",strSql 
	if not rsComprador.eof then
		getDsComprador = rsComprador("DSCOMPRADOR")
	else
		getDsComprador = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getDsProcedencia(p_cdProcedencia)
	dim strSql, l_rsProcedencia	
	
	strSql = ""
	strSql = strSql & "SELECT dsProcedencia "
	strSql = strSql & "FROM   Procedencias "
	strSql = strSql & "WHERE  cdProcedencia = " & p_cdProcedencia
	
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsProcedencia, "OPEN",strSql 
	if not l_rsProcedencia.eof then
		getDsProcedencia = l_rsProcedencia("dsProcedencia")
	else
		getDsProcedencia = " "
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function editKilos (p_valor)
	editKilos = Clng(p_valor)/1000
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function hayBuque ()
	'verifica, si hay buque en el puerto - estados embarque
	' 5 - ammarado, 10-iniciada carga, 11 - finalizada carga
	dim strSql, l_rsEmbarques
	'EAB VER
	'strSql = "select * from embarques where icestadoliq in (5,10,11)"
	strSql = "select * from embarques"
	GF_BD_Puertos session("TERMINAL_ACTUAL"), l_rsEmbarques, "OPEN",strSql 
	if l_rsEmbarques.eof then
		hayBuque = false
	else
		hayBuque = true
	end if
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function getCantidadCamionesAutorizar()
'Toma la cantidad de camiones que deberian estar actualizados.
dim strSql, p_rsCamionesAutorizar	
	strSql = "select count(*) as CantCamionesAutorizar " 
	strSql = strSql & " from camionesdescarga, camiones, caladadecamiones "
	strSql = strSql & " where camiones.idcamion = camionesdescarga.idcamion "
	strSql = strSql & " and camiones.idcamion = caladadecamiones.idcamion "
	strSql = strSql & " and camiones.cdproducto = " & cnstCdProducto & " and camionesdescarga.cdcliente <> " & cnstCdCliente
	strSql = strSql & " and caladadecamiones.cdaceptacion = " & ACEPTACION_SIN_AUTORIZACION
	
	'response.write strSql
	'response.end
	GF_BD_Puertos g_strPuerto, p_rsCamionesAutorizar, "OPEN",strSql 
	getCantidadCamionesAutorizar = CInt(p_rsCamionesAutorizar("CantCamionesAutorizar"))
end function
'--------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'Objetivo: Leer y tomar aquellos registros de un determinado parametro que puede ser 
'		   unicamente modificados por el administrador(Editable y Puesto).
Function traerParametrosEditables(pCdParam, ppto)
	dim strSQL,rs, rtrn
	strSQL = "SELECT * FROM tblparametrosextra where CDPARAMETRO = '"& pCdParam &"'"
	GF_BD_Puertos ppto, rs, "OPEN",strSQL 	
	if not rs.EoF then 
		rtrn = rs		
	end if
	set traerParametrosEditables = rtrn
End Function

'-----------------------------------------------------------------------------------
'Lee los parametros de la cabecera, y en caso de que tenga detalle tambien los trae. A esto se
'le aplica un filtro segun la busqueda que realiza
Function leerParametros(ppto,pcdParam,pnomParam,pidPuesto,setOrder,pTipoFiltro)
	dim strSQL, rs, myWhere, myInnerJoin
	call buscarFiltrosParametros(myWhere,pcdParam,pnomParam,pidPuesto,pTipoFiltro)
	if(pidPuesto <> 0)then myInnerJoin = " INNER JOIN TBLPARAMETROSEXTRA B ON A.CDPARAMETRO = B.CDPARAMETRO "
	strSQL = "SELECT * FROM PARAMETROS A" & myInnerJoin & myWhere &setOrder	
	GF_BD_Puertos ppto, rs, "OPEN",strSql 
	Set leerParametros = rs
End function
'-----------------------------------------------------------------------------------
'Permite ver si un parametro dado es editable y si tiene asignado un puesto 
Function tieneParametrosExtra(pCdParam, ppto)
	dim strSQL,rs, rtrn	
	strSQL = "SELECT * FROM TBLPARAMETROSEXTRA where CDPARAMETRO = '"& pCdParam &"'"	
	GF_BD_Puertos ppto, rs, "OPEN",strSQL 	
	rtrn = false
	if not rs.EoF then 
		rtrn = true 
	end if
	tieneParametrosExtra = rtrn
End Function
'-----------------------------------------------------------------------------------
Function leerPermisos(ppto, taskList)
	dim strSQL, rtrn, rs2
	'Averiguo el cargo del usuario	
	
	strSQL = "SELECT P.cdTASK FROM groupUser GU inner join PERMISSIONS P on GU.CDGROUP=P.CDGROUP "
	strSQL = strSQL & "	WHERE GU.cdusername = '"&session("Usuario")&"' and P.CDTASK in (" & taskList & ")"
	strSQL = strSQL & " order by P.CDTASK DESC"	
	Call GF_BD_Puertos(ppto, rs, "OPEN",strSQL)	
	rtrn = NO_TIENE_PERMISO	
	if (not rs.EOF) then rtrn = Cint(rs("CDTASK"))
	leerPermisos = rtrn
End function
'-----------------------------------------------------------------------------------
'obtiene el nombre de un puesto 
Function obtenerNombrePuesto(pIdPuesto,ppto)
	dim strSQL, rs,rtrn
	strSQL = "SELECT dspuesto FROM tblpuestosplanta WHERE idpuesto =" & pIdPuesto
	GF_BD_Puertos ppto, rs,"OPEN",strSQL 	
	if not rs.EoF then 		
		rtrn = rs("dspuesto")		
	end if	
	obtenerNombrePuesto = rtrn
End function
'-----------------------------------------------------------------------------------
Function buscarFiltrosParametros(ByRef myWhere,pcdParam,pnomParam,pidPuesto,pTipoFiltro)
	if(pTipoFiltro)then
		if (pcdParam <> "") then Call mkWhere(myWhere, "A.CDPARAMETRO", pcdParam, "LIKE", 3)
		if (pnomParam <> "") then Call mkWhere(myWhere, "A.DSPARAMETRO", pnomParam, "LIKE", 3)
	else 
		if (pcdParam <> "") then myWhere = " Where A.CDPARAMETRO LIKE '%"&pcdParam&"%'"
		if (pnomParam <> "") then 
			if(len(myWhere) > 0)then 
				myWhere = myWhere &" and A.DSPARAMETRO LIKE '%"&pnomParam&"%'"
			else
				myWhere = " Where A.DSPARAMETRO LIKE '%"&pnomParam&"%'"
			end if
		end if
	end if
	if (pidPuesto <> 0) then Call mkWhere(myWhere, "B.PUESTO", pidPuesto, "=", 1)
	buscarFiltrosParametros = myWhere
End function
'---------------------------------------------------------------------------------
Function leerPuestos(pto)
	dim strSQL, rs
	strSQL = "SELECT * FROM tblpuestosplanta"
	GF_BD_Puertos pto, rs, "OPEN",strSQL
	Set leerPuestos = rs
End Function						
'---------------------------------------------------------------------------------
Function leerControlBalanza(pPto,pFechaDesde,pFechaHasta,pPatente,pAcoplado,pEstado, pTControl)
	Dim rs, strSQL, myWhere
	Call buscarFiltrosControlBalanza(myWhere,pFechaDesde,pFechaHasta,pPatente,pAcoplado,pEstado, pTControl)
	strSQL = " SELECT * FROM CTRLBZACAMIONES " & myWhere & " ORDER BY IDCONTROL DESC"	
	Call GF_BD_Puertos(pPto, rs, "OPEN",strSQL)
	Set leerControlBalanza = rs
End Function
'-----------------------------------------------------------------------------------
Function buscarFiltrosControlBalanza(ByRef myWhere,myFechaDesde,myFechaHasta,patente,acoplado,estado, tControl)
	if (myFechaDesde <> "") then Call mkWhere(myWhere, "FECHA", myFechaDesde, ">=", 1)
	if (myFechaHasta <> "") then Call mkWhere(myWhere, "FECHA", myFechaHasta, "<=", 1)	
	if (patente <> "") then Call mkWhere(myWhere, "CDCHAPACAMION", patente, "=", 3)
	if (acoplado <> "") then Call mkWhere(myWhere, "CDCHAPAACOPLADO", acoplado, "=", 3)
	if (estado <> BZA_CAM_ESTADO_TODOS)then
		if (estado = BZA_CAM_ESTADO_EN_CURSO)then		
			Call mkWhere(myWhere, "ESTADO", BZA_CAM_ESTADO_FINALIZADO, ">", 1)
		else	
			Call mkWhere(myWhere, "ESTADO", estado , "=", 1)
		end if
	end if		
	if (tControl <> BZA_CAM_TIPO_CTRL_TODOS)then Call mkWhere(myWhere, "TIPOCONTROL", tControl , "=", 3)
		
	buscarFiltrosControlBalanza = myWhere
End function
'-----------------------------------------------------------------------------------------------------------------
Function getLetraPuerto(pName)
    dim rtrn
    rtrn = ""
    select case ucase(pName) 
	    case TERMINAL_TRANSITO
		    rtrn = "T"
	    case TERMINAL_PIEDRABUENA
		    rtrn = "P"
	    case TERMINAL_ARROYO
		    rtrn = "N"
    end select 	
    getLetraPuerto = rtrn
End Function
'-----------------------------------------------------------------------------------------------------------------
Function getIdDivision(pPto)
    Dim ret, strSQL, rs
    
    ret = 0
    strSQL="SELECT IDDIVISION FROM TBLDIVISIONES WHERE CDDIVISIONABR = '"& getLetraPuerto(pPto) & "'"    
    Call GF_BD_Puertos(session("TERMINAL_ACTUAL"), rs, "OPEN",strSQL)
    if (not rs.eof) then ret = CInt(rs("IDDIVISION"))
    getIdDivision = ret
    
End Function
'-----------------------------------------------------------------------------------------------------------------
Function getNumeroPuerto(pName)
    dim rtrn
    rtrn = -1
    select case ucase(pName) 
	    case TERMINAL_TRANSITO
		    rtrn = 10
	    case TERMINAL_PIEDRABUENA
		    rtrn = 91
	    case TERMINAL_ARROYO
		    rtrn = 36
    end select 	
    getNumeroPuerto = rtrn
End function
'-----------------------------------------------------------------------------------------------------------------
Function getNombrePuerto(pName)
    dim rtrn
    rtrn = ""
    select case ucase(pName) 
	    case TERMINAL_TRANSITO
		    rtrn = "Terminal Pto. San Martin/San Lorenzo"
	    case TERMINAL_PIEDRABUENA
		    rtrn = "Terminal Piedrabuena"
	    case TERMINAL_ARROYO
		    rtrn = "Terminal Arroyo"
    end select 	
    getNombrePuerto = rtrn
End function
'-------------------------------------------------------------------------------------------------
Function getDsPuertoByNro(pNro)
    dim rtrn
    rtrn = ""
    Select Case UCase(pNro)
		Case 10
			rtrn = TERMINAL_TRANSITO
		Case 91
			rtrn = TERMINAL_PIEDRABUENA
		Case 36
			rtrn = TERMINAL_ARROYO
	End Select	
    getDsPuertoByNro = rtrn
End Function

'----------------------------------------------------------------------------------------------
' Funci�n:	  getDsCodigoAjustePuerto
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  26/08/2013
' Objetivo:	  Se obtiene la descripcion de un Codigo de Ajuste de Puerto.
' Parametros: 
'			  pCdAjuste	  [string]
' Devuelve:
'			  dsAjuste    [string]
'----------------------------------------------------------------------------------------------
Function getDsCodigoAjustePuerto(pCdAjuste)
	dim rtrn
    rtrn = ""
    select case ucase(pCdAjuste) 
	    case AJUSTE_DRAFT_SURVEY
		    rtrn = "Diferencia Draft/Balanza"
	    case AJUSTE_CALIDAD
		    rtrn = "Merma por Calidad"
	    case AJUSTE_MERMA_VOLATIL
		    rtrn = "Merma volatil"
        case AJUSTE_MANIPULEO
		    rtrn = "Other Blendings"        
    end select 	
    getDsCodigoAjustePuerto = rtrn	
End Function 
'----------------------------------------------------------------------------------------------
' Funci�n:	  updateAjustePuerto
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  27/08/2013
' Objetivo:	  Actualiza un Ajuste de Pto, de un determinado Codigo Ajuste y Origen
' Parametros: 
'			  pCdAjuste   [string]
'			  pIdDraft    [int]
'			  pFechaDS    [string]
'			  pDifKilos   [decimal]
' Devuelve:	  -
'----------------------------------------------------------------------------------------------
Function updateAjustePuerto(pCdAjuste, pIdDraft, pFechaDS, pDifKilos, pEstado)
	Dim strSQL
	strSQL = "UPDATE TBLAJUSTES SET KILOSAJUSTE="&pDifKilos&",ESTADO="&pEstado&",FECHADESDE="&pFechaDS&",FECHAHASTA="&pFechaDS&",CDUSER='"&session("usuario")&"',MMTO="&session("MmtoSistema")&" WHERE CDAJUSTE='"&pCdAjuste&"' AND IDORIGEN="&pIdDraft	
	Call GF_BD_Puertos (g_strPuerto, rs, "EXEC",strSQL)
End Function
'----------------------------------------------------------------------------------------------
' Funci�n:	  grabarDraftSurvey
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  27/08/2013
' Objetivo:	  Graba un nuevo Draft Survey, y devuelve el idDraft creado
' Parametros: 
'			  pCdAviso   [string]
'			  pCdProducto    [int]
'			  pKilosDS    [decimal]
'			  pFechaDS   [string]
'			  pFechaBza   [string]
'			  pKilosBza   [decimal]
' Devuelve:	  idDraft   [int]
'----------------------------------------------------------------------------------------------
Function grabarDraftSurvey(pCdAviso, pCdProducto, pKilosDS, pFechaDS, pFechaBza, pKilosBza, pKilosBzaToepfer)
	Dim strSQL, idDraft 			
	strSQL = "INSERT INTO TBLEMBARQUESDRAFTSURVEY(CDAVISO,CDPRODUCTO,FECHABALANZA,TOTALBALANZA,FECHADRAFT,TOTALDRAFT,CDESTADO,CDUSUARIO,MMTO, KGBZATOEPFER)"&_
			 " VALUES ("& pCdAviso &","& pCdProducto &","& pFechaBza &","& pKilosBza &","& pFechaDS &","& pKilosDS &","& ESTADO_ACTIVO &",'"& UCase(session("Usuario")) &"',"& session("MmtoSistema") &", " & pKilosBzaToepfer & ")"			
	Call GF_BD_Puertos (g_strPuerto, rs, "EXEC",strSQL)	
	strSQL = "SELECT MAX(IDDRAFT) AS IDDRADT FROM TBLEMBARQUESDRAFTSURVEY " 			 
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)	
	if (not rs.Eof) Then idDraft = rs("IDDRADT")
	grabarDraftSurvey = idDraft
End Function
'----------------------------------------------------------------------------------------------
' Funci�n:	  saveDraftAttach
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  --/--/----
' Objetivo:	  Guarda el archivo adjunto de un Draft Survey 
' Parametros: 
'			  pFilePath   [string] (nombre del archivo y extension)
'			  pIdDraft    [int]
' Devuelve:	  -
'----------------------------------------------------------------------------------------------
Function saveDraftAttach(pFilePath, pIdDraft)
	Dim strSQL,rs,conn,extension,fileName, pathOrigen, pathDestino
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	extension = fso.GetExtensionName(pFilePath)
    fileName = fso.getfilename(pFilePath)
	fileName = left(pFilePath,InStrRev(pFilePath,".")-1) 'le quito la extension	
	Call FileName2DbName(filename)
	pathOrigen = server.MapPath(".") & "\Temp\" & pFilePath
	pathDestino = server.MapPath("..") & "\Documentos\Draft Survey\" & g_strPuerto & "\" & filename & "." & extension
	fso.MoveFile pathOrigen, pathDestino 
	
	strSQL = "Update TBLEMBARQUESDRAFTSURVEY set NAMEFILE='" & filename & "', EXTFILE='" & extension & "' where iddraft="&pIdDraft
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)		
End Function
'----------------------------------------------------------------------------------------------
' Funci�n:	  updateDraftSurvey
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  --/--/----
' Objetivo:	  Actualiza los datos de un Draft Survey
' Parametros: 
'			  pIdDraft   [int]
'			  pFechaDS   [string] 
'			  pKilosDs   [decimal] 
'			  pFechaBza  [string]  
'			  pKilosBza  [decimal]
' Devuelve:	  -
'----------------------------------------------------------------------------------------------
Function updateDraftSurvey(pIdDraft,pFechaDS,pKilosDs,pFechaBza,pKilosBza, pKilosBzaToepfer)		
	Dim strSQL	
	strSQL = "UPDATE TBLEMBARQUESDRAFTSURVEY SET FECHADRAFT="&pFechaDS&",FECHABALANZA="&pFechaBza&",TOTALBALANZA="&pKilosBza&",TOTALDRAFT="&pKilosDs&",KGBZATOEPFER=" & pKilosBzaToepfer & ",CDUSUARIO='"&session("usuario")&"',MMTO="&session("MmtoSistema")&" WHERE IDDRAFT="&pIdDraft		
	Call GF_BD_Puertos (g_strPuerto, rs, "EXEC",strSQL)
End Function 
'----------------------------------------------------------------------------------------------
' Funci�n:	  deleteDraftSurvey
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  27/08/2013
' Objetivo:	  Se asigna el estado BAJA a un Draft Survey
' Parametros: 
'			  pIdDraft   [int]
' Devuelve:	  -
'----------------------------------------------------------------------------------------------
Function deleteDraftSurvey(pIdDraft)
	Dim strSQL
	strSQL = "UPDATE TBLEMBARQUESDRAFTSURVEY SET CDESTADO = " & ESTADO_BAJA & " WHERE IDDRAFT = " & pIdDraft
	Call GF_BD_Puertos (g_strPuerto, rs, "EXEC",strSQL)
End Function				
'----------------------------------------------------------------------------------------------
function getDsEmpresa(p_cdEmpresa)	
	dim strSQL		
	strSQL = " SELECT DSEMPRESA " &_
			 " FROM   EMPRESAS " &_
			 " WHERE  CDEMPRESA = " & p_cdEmpresa
	GF_BD_Puertos g_strPuerto, rs, "OPEN",strSQL
	if not rs.eof then
		getDsEmpresa = rs("DSEMPRESA")
	else
		getDsEmpresa = ""
	end if
end function
'------------------------------------------------------------------------------------------------
function getValueParametro(pCdParametro,pPto)
	dim strSQL, rtrn	
	rtrn = ""
	if (session("VAL_PTO" & pPto & "_PRM" & pCdParametro) <> "") then
		rtrn = session("VAL_PTO" & pPto & "_PRM" & pCdParametro)
	else
		strSQL = " SELECT VLPARAMETRO FROM PARAMETROS WHERE CDPARAMETRO = '" & pCdParametro &"'"
		Call executeQueryDb(pPto, rs, "OPEN", strSQL)	
		if not rs.eof then rtrn = rs("VLPARAMETRO")
	end if
	session("VAL_PTO" & pPto & "_PRM" & pCdParametro) = rtrn
	getValueParametro = rtrn
end function
'------------------------------------------------------------------------------------------------
Function getCodigoBolsa(pName)
    dim rtrn
    rtrn = -1
    select case ucase(pName) 
	    case TERMINAL_TRANSITO
		    rtrn = 91
	    case TERMINAL_PIEDRABUENA
		    rtrn = 92
	    case TERMINAL_ARROYO
		    rtrn = 90
    end select 	
    getCodigoBolsa = rtrn
End function
'------------------------------------------------------------------------------------------------
Function getDsRubro(pCdRubro)
	Dim strSQL
	strSQL = "SELECT DSRUBRO FROM RUBROS WHERE CDRUBRO ="& pCdRubro
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)	
	if not rs.eof then rtrn = rs("DSRUBRO")
	getDsRubro = rtrn
End Function
'------------------------------------------------------------------------------------------------
Function updateValueParametro(pCdParametro,pVlParametro,pPto)
	Dim strSQL	
	strSQL = "UPDATE PARAMETROS SET VLPARAMETRO = '"&pVlParametro&"' WHERE CDPARAMETRO = '"&pCdParametro&"'"	
	Call executeQueryDb(pPto, rsPto, "EXEC", strSQL)
	session("VAL_PTO" & pPto & "_PRM" & pCdParametro) = pVlParametro
End Function
'------------------------------------------------------------------------------------------------
function getDsTransporte(p_cdTransporte)
	dim strSql, l_rsTransporte

	strSql = ""
	strSql = strSql & "SELECT dsTipoTransporte "
	strSql = strSql & "FROM   TiposTransportes "
	strSql = strSql & "WHERE  cdTipoTransporte = " & p_cdTransporte	
	Call GF_BD_Puertos(g_strPuerto, l_rsTransporte, "OPEN",strSql) 
	
	if not l_rsTransporte.eof then
		getDsTransporte = l_rsTransporte("dsTipoTransporte")
	else
		getDsTransporte = " "
	end if
	
end function
'-------------------------------------------------------------------------------------------------
Function getDsPuertoByLetra(pLetra)
    dim rtrn
    rtrn = ""
    select case Trim(ucase(pLetra))
	    case "T"
		    rtrn = TERMINAL_TRANSITO
	    case "P"
		    rtrn = TERMINAL_PIEDRABUENA
	    case "N"
		    rtrn = TERMINAL_ARROYO
    end select
    getDsPuertoByLetra = rtrn
End Function
'-------------------------------------------------------------------------------------------------
Function getPuertoByCodigoBolsa(pCdBolsa)
    dim rtrn
    rtrn = ""
    select case Cdbl(pCdBolsa)
	    case 91
		    rtrn = TERMINAL_TRANSITO
	    case 92
		    rtrn = TERMINAL_PIEDRABUENA
	    case 90
		    rtrn = TERMINAL_ARROYO
    end select 	
    getPuertoByCodigoBolsa = rtrn
End function
'---------------------------------------------------------------------------------------------------------
Function cargarTablaConversion(pCuitCliente, pPto)
    
    Dim strSQL, rs, ret, auxkey, auxval  
    
    Set gDicConv = createObject("Scripting.Dictionary")    
    ret = false    
    strSQL="Select * from TBLCONVERSIONES where NUCUITCLIENTE='" & pCuitCliente& "'"
    Call executeQueryDb(pPto, rs, "OPEN", strSQL)    
    while (not rs.eof)
        auxkey = rs("TIPODATO") & "_" & rs("CDPROPIO")
        auxval = rs("CDTERCERO")
        gDicConv.Add auxkey, auxval
        ret = true
        rs.MoveNext()
    wend    
    cargarTablaConversion = ret
    
End Function
'--------------------------------------------------------------------------------------------------
'Para convertir datos se debe previamente cargar la tabla de conversi�n llamando a la funcion: cargarTablaConversion
Function convertirDatoPuerto(pTipo, pCodigoPropio, ByRef pErrMsg)
    Dim rtrn
    
    rtrn=""
	pErrMsg = ""
    if (gDicConv.Exists(pTipo & "_" & pCodigoPropio)) then        
        rtrn = gDicConv.Item(pTipo & "_" & pCodigoPropio)
	else
		pErrMsg = "Falta traduccion TIPODATO: " &  pTipo & ", CDPROPIO: " & pCodigoPropio & "<br>"
    end if
    convertirDatoPuerto = rtrn    
End Function


%>

						
						
						
						
						
						
						
						
						
						
						
						
						
						