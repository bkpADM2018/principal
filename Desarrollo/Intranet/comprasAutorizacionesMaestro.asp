<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<%
Const COMPRAS_AUTORIZACIONES	 = 2 'PROVIENE DE COMPRAS AUTORIZACIONES
Const COMPRAS_AVISOS_AUTOMATICOS = 1 'PROVIENE DE AVISOS AUTOMATICAS
Const PREFIJO_ANULACION = "X"
'------------------------------------------------------------------------------------------------
Function armarSQLAjustePartidaPresupuestaria()
    Dim rs
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACION_GET_CBTES_A_FIRMAR_ALL_USERS", "")
    'if (not rs.eof) then armarSQLAjustePartidaPresupuestaria = armarLineaTipoDocumento(rs)
	if (not rs.eof) then armarSQLAjustePartidaPresupuestaria = ArmarListaTipoDocumento(rs, AUTH_TYPE_APP)
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjustePuerto(pPto, pTipoAjuste)
    Call executeProcedureDb(pPto, rs, "TBLAJUSTES_GET_CBTES_A_FIRMAR_ALL_USERS", pPto &"||"& pTipoAjuste)
    if (not rs.eof) then armarSQLAjustePuerto = armarLineaTipoDocumento(rs)
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLPlanillas()
	
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPCTCABECERA_GET_CBTES_A_FIRMAR_ALL_USERS", "1||0")
	if not rs.eof then armarSQLPlanillas = ArmarListaTipoDocumento( rs, AUTH_TYPE_PCP)	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuCTC()
	Dim strSQL
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACTCAJUSTES_GET_CBTES_A_FIRMAR_ALL_USERS", "1||0" )	
	if not rs.eof then armarSQLAjuCTC = ArmarListaTipoDocumento( rs, AUTH_TYPE_CTC)	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLPics()
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_CBTES_A_FIRMAR_ALL_USERS", "1||1||0" )
	if not rs.eof then armarSQLPics = ArmarListaTipoDocumento( rs, AUTH_TYPE_PIC)
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLCEC()
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_CBTES_A_FIRMAR_ALL_USERS", "0||1||0" )
	if not rs.eof then armarSQLCEC = ArmarListaTipoDocumento( rs, AUTH_TYPE_CEC)
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuPics()
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZAJUSTES_GET_CBTES_A_FIRMAR_ALL_USERS", "1||1||0" )
	if not rs.eof then armarSQLAjuPics = ArmarListaTipoDocumento( rs, AUTH_TYPE_AIC)	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuCEC()
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZAJUSTES_GET_CBTES_A_FIRMAR_ALL_USERS", "0||1||0" )
	if not rs.eof then armarSQLAjuCEC = ArmarListaTipoDocumento( rs, AUTH_TYPE_AEC)	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLVales(pTipo)
	Dim strSqlBase, tipoVale, estado
	'Puede firmar ajustes y sus anulaciones
	'Determino el estado del comprobante
	tipoVale = "'"& pTipo &"'"
	if (Left(pTipo,1) <> PREFIJO_ANULACION) then
		estado = ESTADO_ACTIVO		
	else
		estado = ESTADO_ANULACION		
	end if
	
	strSqlBase = 		 	  " select vf.idVALE iddocumento, vc.CDVALE tipo, 0 idpedido, vc.IDOBRA IDOBRA, vc.NRVALE CDDOCUMENTO, C.cdobra CDOBRA, 0 AS CDPEDIDO, 0 AS IDPROVEEDOR, '' DSPROVEEDOR, vf.CDUSUARIO CDUSUARIO, vc.IDALMACEN  IDALMACEN,'' PTO "
	strSqlBase = strSqlBase & " from (Select IDVALE, CDUSUARIO, Min(SECUENCIA) as secuencia from TBLVALESFIRMAS where HKEY is null OR HKEY = '' group by IDVALE, CDUSUARIO) vf "
	strSqlBase = strSqlBase & " inner join (Select * from TBLVALESCABECERA where ESTADO = " & estado & " and CDVALE in (" & tipoVale & ")) vc on vc.IDVALE = vf.IDVALE "
	strSqlBase = strSqlBase & "	LEFT JOIN TBLDATOSOBRAS C on C.IDOBRA = vc.IDOBRA " 	
	rtrn = strSqlBase			

	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", rtrn)
	if not rs.eof then armarSQLVales = ArmarListaTipoDocumento(rs, pTipo)		
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAfes()
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLDATOSAFE_GET_CBTES_A_FIRMAR_ALL_USERS", "1||0")

	if not rs.eof then armarSQLAfes = ArmarListaTipoDocumento( rs, AUTH_TYPE_AFE)	
End Function
'------------------------------------------------------------------------------------------------
'Funcion: getDsUsuarioTipoDocumento(cdUsuario, DsUsuario, IdDocumento)
'		  Devuelve la descrpcion del Usuario
Function getDsUsuarioTipoDocumento(pCdUsuario, pIdDocumento, pTipo)
Dim DsUsuario
if (isNull(pCdUsuario) or (pCdUsuario = "") or (isNumeric(pCdUsuario)) or (pCdUsuario = VS_AUDIT_USER) or (pCdUsuario = VS_PORT_SUPERVISOR_USER) or (pCdUsuario = VS_NO_USER)or (pCdUsuario = DIRECTOR_USER)or (pCdUsuario = CONTROLLER_USER) or (pCdUsuario = VS_PORT_GERENTE_USER) or (pCdUsuario = LEGALES_USER)) then
	select case pTipo
		case AUTH_TYPE_AJS, AUTH_TYPE_XJS
			Select case pCdUsuario
				case VS_AUDIT_USER:
					DsUsuario = "AUDITORIA"
				case VS_PORT_SUPERVISOR_USER:
					DsUsuario = "COORDINADOR DE PUERTOS"
				case VS_NO_USER
					DsUsuario = "GERENTE PUERTO"
				case DIRECTOR_USER:
					DsUsuario = "DIRECTOR"	
			End Select
		case AUTH_TYPE_VRS, AUTH_TYPE_XRS:
			DsUsuario = "COORDINADOR DE PUERTOS"
		case AUTH_TYPE_PIC, AUTH_TYPE_AIC, AUTH_TYPE_PCP:								
			Select case pCdUsuario
				case DIRECTOR_USER:
					DsUsuario = "DIRECTOR"
				case else	
					DsUsuario = "COORDINADOR DE PUERTOS"
			End Select		
		case AUTH_TYPE_AJD, AUTH_TYPE_AJC, AUTH_TYPE_AJM:
			Select case pCdUsuario
				case DIRECTOR_USER:
					DsUsuario = "DIRECTOR"
				case CONTROLLER_USER:
					DsUsuario = "CONTROLLER"
				case VS_PORT_GERENTE_USER
					DsUsuario = "GERENTE PUERTO"
			End Select
		case AUTH_TYPE_CCN:			
			Select case Cdbl(pCdUsuario)
				case FIRMA_ROL_RESP_CONTADURIA:
					DsUsuario = "RESPONSABLE DE CONTADURÍA"
				case FIRMA_ROL_RESP_PUERTO:
					DsUsuario = "GERENTE PUERTO"
			End select	
		case else
			DsUsuario = "ERROR: usr=" & pCdUsuario & "- doc=" & pTipo
	end select
else								
	DsUsuario = getUserDescription(pCdUsuario)
end if	
getDsUsuarioTipoDocumento = DsUsuario
End function
'------------------------------------------------------------------------------------------------
'Funcion: ArmarListaTipoDocumento(pRs, pTipo) : Se encarga de ir armando una lista de los registro del recordset, cada campo esta separado
'									   por "|" y cada registro es separado por ";"
'Parametros: 
'		    - pRs:   es el recordset de la SQL
'			- pTipo: es el Tipo de Documento 
'Fecha: 12/09/2012   -  CNA
Function ArmarListaTipoDocumento(pRs, pTipo)
	Dim	listOfDocumento, cdPedido, cdObra, idProveedor, dsProveedor, idPedido, idObra, myUsuario, idAlmacen
	While (not pRs.eof) 
		cdPedido = getRSValue(pRs, "CDPEDIDO", "Sin Pedido")		
		cdObra = getRSValue(pRs, "CDOBRA", "Sin Obra")		
		idProveedor = getRSValue(pRs, "IDPROVEEDOR", 0)		
		idPedido = getRSValue(pRs, "IDPEDIDO", 0)		
		idObra = getRSValue(pRs, "IDOBRA", 0)		
		dsProveedor = getDescripcionProveedor(idProveedor)
		idAlmacen = getRSValue(pRs, "IDALMACEN", 0)		
		'CON EL PARAMETRO 'ORIGEN' VERIFICO DE DONDE ES LLAMADA LA PAGINA
		if origen = COMPRAS_AVISOS_AUTOMATICOS then 
			'SI ES DE COMPRAS AUTORIZACIONES SE LE PASA LOS PARAMETROS NECESARIOS
			listOfDocumento = listOfDocumento & Trim(pRs("CDUSUARIO")) & "|" & idAlmacen &"|"& Trim(pRs("TIPO")) &";"
		else if (origen = COMPRAS_AUTORIZACIONES) then				
				if (checkUserDireccion(Trim(pRs("CDUSUARIO")))) then
					myUsuario = Trim(pRs("CDUSUARIO"))
					listOfDocumento = listOfDocumento & Trim(pRs("TIPO")) & "|" & Trim(cdObra) & "|" & idObra & "|" & Trim(cdPedido) &"|" & idPedido &"|" & idProveedor  &"|" & Trim(dsProveedor) &"|" & pRs("IDDOCUMENTO") &"|" & Trim(pRs("CDDOCUMENTO")) &"|" & getDSDocumentoFirmar(pRs("TIPO")) &"|" & myUsuario &"|"& pRs("PTO") &";"
				end if
			else
				'SI ES OTRA PAGINA SE LE PASA LA DESCRIPCION DEL USUARIO				
				myUsuario = getDsUsuarioTipoDocumento(Trim(pRs("CDUSUARIO")),pRs("IDDOCUMENTO"),pRs("TIPO"))
				listOfDocumento = listOfDocumento & Trim(pRs("TIPO")) & "|" & Trim(cdObra) & "|" & idObra & "|" & Trim(cdPedido) &"|" & idPedido &"|" & idProveedor  &"|" & Trim(dsProveedor) &"|" & pRs("IDDOCUMENTO") &"|" & Trim(pRs("CDDOCUMENTO")) &"|" & getDSDocumentoFirmar(pRs("TIPO")) &"|" & Trim(myUsuario)  &"|"&  idAlmacen &"|"& pRs("PTO") &"|0;"				
			end if
		end if
		pRs.MoveNext()	
	wend
	ArmarListaTipoDocumento = listOfDocumento
End function
'------------------------------------------------------------------------------------------------
Function getRSValue(rs, fieldName, defValue)
    
   Dim fld
   
   fieldName = UCase(fieldName)
    
   getRSValue = defValue
   For Each fld In rs.Fields
      If UCase(fld.Name) = fieldName Then
         if((not isNull(rs(fieldName))) and (rs(fieldName) <> "")) then getRSValue = rs(fieldName)
         Exit Function
      End If
   Next
End Function
'------------------------------------------------------------------------------------------------
'armarSQLCierreAlmacenes: Arma el sql de las firmas pendientes de los cierres contables.
'						  Este tipo de documento tiene una página donde muestra todos los cierres juntos(por division)
'						  que se tiene para firmar, es decir que para identificar a un cierre necesito solo la division,
'						  por este motivo en la SQL se asigna al IdDocumento el valor de la division
Function armarSQLCierreAlmacenes()
	Dim strSQL	
	strSQL = "SELECT DISTINCT cab.IDDIVISION iddocumento, " &_
			 "		 '"& AUTH_TYPE_CCN &"' AS tipo," &_
			 "		 0 as idpedido, " &_
			 "		 0 as IDOBRA, " &_
			 "       Cast(DIV.DSDIVISION as varchar(30)) + '-' + Cast(cab.anio as varchar(10)) + '/' + Cast(cab.mes as varchar(10)) CDDOCUMENTO, "&_ 				 
			 "		 '' as CDOBRA, " &_
			 "		 '' as CDPEDIDO, " &_
			 "		 0 as IDPROVEEDOR, " &_
			 "		 '' as DSPROVEEDOR, " &_
			 "       fir.secuencia cdusuario, " &_
			 "		 0 idalmacen,  " &_
			 "		 '' PTO  " &_
			 "FROM   TBLCIERRESCABECERA2 cab " &_
			 "LEFT JOIN TBLCIERRESFIRMAS2 fir " &_
			 "		ON cab.idcierre = fir.idcierre " &_
			 "INNER JOIN (SELECT idcierre,MIN(SECUENCIA) MINSEC  " &_
			 "			  FROM TBLCIERRESFIRMAS2 " &_
			 "			  WHERE    (HKEY = '' or HKEY is Null) " &_
			 "			  GROUP BY idcierre) AS minFirma  " &_
			 "		ON minFirma.idcierre = fir.idcierre AND minFirma.MINSEC = fir.SECUENCIA " &_
			 "INNER JOIN tbldivisiones div on div.iddivision = cab.iddivision " &_
			 "WHERE cab.ESTADO = '"& TIPO_CIERRE_PROVISORIO &"'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then armarSQLCierreAlmacenes = ArmarListaTipoDocumento( rs, AUTH_TYPE_CCN)
End Function
'------------------------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
           if (not isEmpty(p_strValue)) then
              if (isEmpty(p_strParam)) then
                 p_strParam = "?"
              else
                 p_strParam = p_strParam & "&"
              end if
              p_strParam = p_strParam & p_strKey & "=" & p_strValue
           end if
End Function
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'Verifica si el usuario pertenece a la Direccion
'		True  --> Si pertenece
'		False --> No pertenece
Function checkUserDireccion(pUsuario)
	Dim rs, ret
	
	ret = false
	if (pUsuario = DIRECTOR_USER) then
		ret = true
	else	    
	    'Puede que haya venido la descrición de un rol, buco la descripcino del rol director para comparar.
	    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_DIRECTOR)
	    if (not rs.eof) then
	        if (UCase(pUsuario) = UCase(rs("DSROL"))) then ret = true
	    end if
	    if (not ret) then	    	    
	        idRol = getRolFirma(left(pUsuario, 3), SEC_SYS_COMPRAS)	        
	        if (idRol = FIRMA_ROL_DIRECTOR) then ret = true	    
	    end if
	end if	
	
	checkUserDireccion = ret
End Function
'------------------------------------------------------------------------------------------------
'Esta funcion arma el string con los campos necesarios para mostrar en pantalla (por ajax), ademas se carracteriza por 
' trabajar con el nuevo formato de ROLES, es decir que integra ya el dato del rol traido de la base de datos pero
' ojo que en algunos sistemas tambien se sigue usando el usuario que tiene que firmar
'IMPORTANTE|Para el correcto funcionamiento de esta funcion el recordset recibido por parametro debera tener los siguientes campos:
'       CANTIDAD --> cantidad de documentos que tiene para firmar 
'       IDROL --> identificador del rol a firmar
'       DSROL --> descripcion del rol a firmar
'       CDUSUARIO --> usuario que debera firmar
'       TIPO --> tipo de documento a firmar
'       PTO --> el puerto del documento
Function armarLineaTipoDocumento(pRs)
	Dim listOfDocumento, auxFirmante
	'Solo arma la linea si es llamada de la pagina comprasAutorizacionesMaestro.asp
	if ((origen <> COMPRAS_AVISOS_AUTOMATICOS)and(origen <> COMPRAS_AUTORIZACIONES)) then
		while not pRs.Eof		
            auxFirmante = Ucase(Trim(pRs("DSROL")))
            'Por defecto tomamos el rol a firmar, si tiene un usuario asignado obtnemos y mostramos la descricion del usuario
            if (pRs("CDUSUARIO") <> "") then auxFirmante = getUserDescription(pRs("CDUSUARIO"))
			listOfDocumento = listOfDocumento & Trim(pRs("TIPO")) & "||0||0|0||0||" & getDSDocumentoFirmar(pRs("TIPO")) &"|" & auxFirmante &"|0|"& pRs("PTO") &"|"& pRs("CANTIDAD") &";"
			pRs.MoveNext()
		wend	
	end if	
	armarLineaTipoDocumento = listOfDocumento
End Function
'------------------------------------------------------------------------------------------------
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim  rs, conn, strSQL, cdObra, paginaActual, mostrar, idProv, dsProv, dsUsuario, cdusuario,pTipo,vTipoDocumento,params,origen

'Call comprasControlAccesoCM(RES_ADM)

origen = GF_PARAMETROS7("origen",0,6)
call addParam("origen", origen, params)
pTipo = GF_PARAMETROS7("Tipo","",6)
call addParam("Tipo", pTipo, params)
accion = GF_PARAMETROS7("accion","",6)
call addParam("accion", accion, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual = 1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10


GP_ConfigurarMomentos


If(accion = ACCION_PROCESAR)then
	Select case (pTipo)
		case AUTH_TYPE_PCP: ' PCP '
			listOfDocumento = armarSQLPlanillas()
		case AUTH_TYPE_CTC: ' CONTRATOS '		
			listOfDocumento = armarSQLAjuCTC()
		case AUTH_TYPE_PIC: ' PIC '
			listOfDocumento = armarSQLPics()
        case AUTH_TYPE_CEC: ' CEC '
			listOfDocumento = armarSQLCEC()			
		case AUTH_TYPE_AIC: '  AJS PIC	'	
			listOfDocumento = armarSQLAjuPics()
        case AUTH_TYPE_AEC: '  AJS CEC	'	
			listOfDocumento = armarSQLAjuCEC()			
		case AUTH_TYPE_AFE: ' AFE	'
			listOfDocumento = armarSQLAfes()
		case AUTH_TYPE_VRS: ' RECASIFICACION STOCK , RECASIFICACION STOCK ANULACION,AJUSTE STOCK , AJUSTE STOCK ANULACION '
			listOfDocumento = armarSQLVales(CODIGO_VS_RECLASIFICACION_STOCK)
			'listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_RECLASIFICACION_STOCK_X)
			listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_AJUSTE_STOCK)
			'listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_AJUSTE_STOCK_X)
		case AUTH_TYPE_AJD: ' AJS DRAFT SURVEY (ARROYO, TRANSITO, PIEDRABUENA)'
			listOfDocumento = armarSQLAjustePuerto(DBSITE_ARROYO, AUTH_TYPE_AJD)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_TRANSITO, AUTH_TYPE_AJD)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_BAHIA, AUTH_TYPE_AJD)	
		case AUTH_TYPE_AJC: ' AJS CALIDAD (ARROYO, TRANSITO, PIEDRABUENA)'
			listOfDocumento = armarSQLAjustePuerto(DBSITE_ARROYO, AUTH_TYPE_AJC)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_TRANSITO, AUTH_TYPE_AJC)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_BAHIA, AUTH_TYPE_AJC)
		case AUTH_TYPE_AJM: ' AJS MANIPULEO (ARROYO, TRANSITO, PIEDRABUENA)'
			listOfDocumento = armarSQLAjustePuerto(DBSITE_ARROYO, AUTH_TYPE_AJM)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_TRANSITO, AUTH_TYPE_AJM)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(DBSITE_BAHIA, AUTH_TYPE_AJM)
		case AUTH_TYPE_CCN: 
			listOfDocumento = armarSQLCierreAlmacenes()
        case AUTH_TYPE_AJV: ' AJS MERMA VOLATIL (ARROYO, TRANSITO, PIEDRABUENA)'
			listOfDocumento = armarSQLAjustePuerto(TERMINAL_ARROYO, AUTH_TYPE_AJV)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_TRANSITO, AUTH_TYPE_AJV)
			listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_PIEDRABUENA, AUTH_TYPE_AJV)	
        case AUTH_TYPE_APP: 'Ajuste Partida Presupuestaria
            listOfDocumento = armarSQLAjustePartidaPresupuestaria()
	end select
	if (Len(listOfDocumento) > 0) then listOfDocumento = left(listOfDocumento,Len(listOfDocumento)-1)
	Response.Write listOfDocumento
	response.end
else
	vTipoDocumento = getDocumentoFirmar()
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}

.divAutorizaciones
{
    TEXT-ALIGN: center;
    BACKGROUND-COLOR: #4682B4;
    COLOR: #ffffff;
    FONT-WEIGHT: bold;
	border-bottom-right-radius : 5px;
	border-bottom-left-radius  : 5px;
	border-top-right-radius    : 5px;
	border-top-left-radius     : 5px;	
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/framework.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var ch = new channel();	
	var imgTipo;
	var onclickTipo;
	var titleTipo;	
	var registrosAcumulados = 0;	
	var arrTipoDomuento = new Array();			
	var vRegs;
	var vTipoDocumento = new Array();
	
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function irHome() {
		location.href = "comprasIndex.asp";
	}		
	function irDirecta(){
		location.href = "comprasAdministrarCotizaciones.asp";
	}	
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}	
	function refresh() {
		location.href = "comprasAutorizacionesMaestro.asp"
	}
	function irIndividual() {
		location.href = "comprasAutorizaciones.asp"
	}
	function abrirVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_blank", "location=no,menubar=no,statusbar=no",false);
	}
	
	function abrirAfe(id) {
		window.open("comprasAFEPrint.asp?idafe=" + id, "_blank", "location=no,menubar=no,statusbar=no",false);
	}
	function abrirCotizacion(id) {
		window.open("comprasPICPrint.asp?idCotizacionElegida=" + id, "_blank", "location=no,menubar=no,statusbar=no",false);
	}
	function abrirPedido(id) {
		window.open("comprasFichaPedidoCotizacion.asp?idPedido="+id+"&tab=1", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
	}
	
	function abrirObra(id) {
		window.open("comprasTableroObra.asp?idObra=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);		
	}
	
	function abrirPlanilla(id) {
		window.open("comprasComparativoDeOfertasPrint.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
	}
	
    function bodyOnLoad() {	
		vRegs = new Paginacion("paginacion");
		var tb = new Toolbar('toolbar', 5, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");
		tb.addButtonREFRESH("Recargar", "refresh()");
		tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");
		tb.addButton("direct_Purchase-16x16.png", "Compra Directa", "irDirecta()");
		tb.addButton("see_all-16x16.png", "Mis Autorizaciones", "irIndividual()");
		tb.draw();
		//Cargo el vector con los tipos de documento.
		<% for i=0 to UBound(vTipoDocumento)-1 %>
			vTipoDocumento.push('<% =vTipoDocumento(i) %>');
		<% next %>		
		//Hago el primer llamado para iniciar la cadena de carga
		var tipo = vTipoDocumento.pop();
		document.getElementById("msjCargando").innerHTML= getDescripcionTipo(tipo);
		ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" + tipo + "&accion=<%=ACCION_PROCESAR%>","CallBack_getAutorizaciones()");
		ch.send();		
	}
	
	var hayDatos = false;
	
	function MostrarListaTipoDocumento(isLast) {
		while (arrTipoDomuento.length > 0) {
			var linea = String(arrTipoDomuento.splice(0,1));			
			var vals = linea.split("|");
			hayDatos = true;
			agregarLineaAutorizaciones(vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],vals[6],vals[7],vals[8],vals[9],vals[10],vals[12],vals[13]);									
		}
		document.getElementById("msjCargando").innerHTML="";
	}
	
	
	function CallBack_getAutorizaciones(){
		var isLast = false;
		//Lanzo la busqueda del siguiente tipo de documento		
		if (vTipoDocumento.length > 0) {
			var tipo = vTipoDocumento.pop();
			document.getElementById("msjCargando").innerHTML= getDescripcionTipo(tipo);				
			ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" +  tipo + "&accion=<%=ACCION_PROCESAR%>","CallBack_getAutorizaciones()");
			ch.send();
		} else {
			isLast = true;
			document.getElementById("msjCargando").innerHTML="";
			if (!hayDatos) {
				var tblAutorizaciones = document.getElementById("MyTablaAutorizaciones");
				var trMsj = tblAutorizaciones.insertRow(1);		
				var tdMsj = trMsj.insertCell(0);		
				tdMsj.align = 'center';
				tdMsj.setAttribute("colspan","10");
				tdMsj.className = 'TDNOHAY';
				tdMsj.innerHTML = "No hay informacion disponible en estos momentos";		
				trMsj.appendChild(tdMsj);
				tblAutorizaciones.appendChild(trMsj);
			}			
		} 		
		//Proceso los resultados obtenidos del servidor		
		var rtrn = ch.response();		
		if (rtrn.length > 0) {		
			var arr = rtrn.split(";");
			arrTipoDomuento = arrTipoDomuento.concat(arr);
			MostrarListaTipoDocumento(isLast);
		}
	}	
	
	function getDescripcionTipo(pTipo){
		var dsTipo;
		if(pTipo == "<%=AUTH_TYPE_PCP%>") dsTipo =  "Cargando... PCP";	
		if(pTipo == "<%=AUTH_TYPE_PIC%>") dsTipo =  "Cargando... PIC";
		if(pTipo == "<%=AUTH_TYPE_CEC%>") dsTipo =  "Cargando... CEC";
		if(pTipo == "<%=AUTH_TYPE_AIC%>") dsTipo =  "Cargando... Ajuste de PIC";
		if(pTipo == "<%=AUTH_TYPE_AEC%>") dsTipo =  "Cargando... Ajuste de CEC";
		if(pTipo == "<%=AUTH_TYPE_AFE%>") dsTipo =  "Cargando... AFE";
		if(pTipo == "<%=AUTH_TYPE_CTC%>") dsTipo =  "Cargando... Contratos";	
		if(pTipo == "<%=AUTH_TYPE_VRS%>") dsTipo =  "Cargando... Vales";
		if(pTipo == "<%=AUTH_TYPE_AJD%>") dsTipo =  "Cargando... Ajuste de Draft Survey";
		if(pTipo == "<%=AUTH_TYPE_AJC%>") dsTipo =  "Cargando... Ajuste de Calidad";
		if(pTipo == "<%=AUTH_TYPE_AJM%>") dsTipo =  "Cargando... Ajuste de Manipuleo";
		if(pTipo == "<%=AUTH_TYPE_CCN%>") dsTipo =  "Cargando... Cierres de Almacenes";
		if(pTipo == "<%=AUTH_TYPE_AJV%>") dsTipo =  "Cargando... Merma Volatil";
		if(pTipo == "<%=AUTH_TYPE_APP%>") dsTipo =  "Cargando... Ajuste Partida Presupuestaria";
		return dsTipo;
		
	}
	
	/* funcion:    agregarLineaAutorizaciones()
	 *			   Se encarga de crear las lineas de la tabla con sus respectivos registros 
	 * Parametros: Todos los paramtros pasadosson el resultado de la consulta SQL traida por ajax, cada uno se va a mostrar en la tabla
	 */
	function agregarLineaAutorizaciones(pTipo, pCdObra, pIdObra, pCdPedido ,pIdPedido ,pIdProvEleg ,pDsProvEleg ,pIdDocumento ,pCdDocumento,pDSDocumentoFirmar,pDsUsuario, pPuerto, pCantidadDoc) {
		var tblAutorizaciones = document.getElementById("MyTablaAutorizaciones");		
		/* Se crea la Fila*/
		var tBody = document.createElement('tbody');
		var trTipo  = document.createElement('tr');
		trTipo.className = 'reg_Header_navdos'; 
		setEvent(trTipo, FWRK_EVT_ON_MOUSE_OUT,	"lightOff(this)")
		setEvent(trTipo, FWRK_EVT_ON_MOUSE_OVER,	"lightOn(this)")
		/* * * * * * * * * * * * * * * * * * * * * *Se crea las Columnas* * * * * * * * * * * * * * * * * * * * /		
		/* TIPO */
				var tdTipo  = document.createElement('td');
				tdTipo.align = 'center';		
				tdTipo.innerHTML = pDSDocumentoFirmar;				
				setEvent(tdTipo, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ")");
				trTipo.appendChild(tdTipo);			
		/* OBRA */
				var tdObra  = document.createElement('td');
				tdObra.align = 'center';		
				if(pIdObra > 0) tdObra.innerHTML = pCdObra;
				trTipo.appendChild(tdObra);	
		/* IMAGEN OBRA */
				var tdImgObra  = document.createElement('td');
				tdImgObra.align = 'center';			
				if(pIdObra > 0)
				{			
					var imgObra = document.createElement("img");
					imgObra.src   = "images/compras/OBR-16x16.png";
					imgObra.title = "Ver Partida";
					imgObra.className = "cursorStyle";					
					setEvent(imgObra, FWRK_EVT_ON_CLICK, "abrirObra(" + pIdObra + ")"); 
					tdImgObra.appendChild(imgObra);				
				}		
				trTipo.appendChild(tdImgObra);		
		/* PEDIDO */
				var tdPedido  = document.createElement('td');
				tdPedido.align = 'center';		
				if(pIdPedido > 0) tdPedido.innerHTML = pCdPedido;								
				trTipo.appendChild(tdPedido);			
		/* IMAGEN PEDIDO */
				var tdImgPedido  = document.createElement('td');
				tdImgPedido.align = 'center';
				if(pIdPedido > 0)
				{			
					var imgPedido = document.createElement("img");
					imgPedido.src   = "images/compras/PCT-16x16.png";
					imgPedido.title = "Ver Pedido";
					imgPedido.className = "cursorStyle";					
					setEvent(imgPedido, FWRK_EVT_ON_CLICK, "abrirPedido(" + pIdPedido + ")");
					tdImgPedido.appendChild(imgPedido);			
				}
				trTipo.appendChild(tdImgPedido);
		/* PROVEEDOR */
				var tdProveedor  = document.createElement('td');
				tdProveedor.align = 'left';										
				var strProveedor = "Sin Datos";
				if(pIdProvEleg > 0) strProveedor = pIdProvEleg +" - "+ pDsProvEleg;
				tdProveedor.innerHTML = strProveedor;
				trTipo.appendChild(tdProveedor);	
		/* IMAGEN PROVEEDOR */
				var tdImgProveedor  = document.createElement('td');
				tdImgProveedor.align = 'center';
				if(pIdProvEleg > 0)
				{			
					var imgProveedor = document.createElement("img");
					imgProveedor.src   = "images/compras/PCP-16x16.png";
					imgProveedor.title = "Ver Planilla Comparativa";
					imgProveedor.className = "cursorStyle";					
					setEvent(imgProveedor, FWRK_EVT_ON_CLICK, "abrirPlanilla(" + pIdPedido + ")");
					tdImgProveedor.appendChild(imgProveedor);			
				}
				trTipo.appendChild(tdImgProveedor);
		/* DOCUMENTO */
				var tdDocumento  = document.createElement('td');
				tdDocumento.align = 'right';								
				var strDocumento = pIdDocumento;
				if (pCantidadDoc == 0){
					tdDocumento.innerHTML = pCdDocumento;					
				}
				else{
					tdDocumento.innerHTML = "Pendientes: " + pCantidadDoc;
				}	
				trTipo.appendChild(tdDocumento);	
		/* IMAGEN DOCUMENTO */
				var tdImgDocumento  = document.createElement('td');
				tdImgDocumento.align = 'center';
				getTipoImagen(pTipo,pIdDocumento, pPuerto);			
				if(imgTipo != ""){					
					var imgDocumento = document.createElement("img");
					imgDocumento.src   = imgTipo;
					imgDocumento.title = titleTipo;			
					imgDocumento.className = "cursorStyle";					
					setEvent(imgDocumento, FWRK_EVT_ON_CLICK, onclickTipo);
					tdImgDocumento.appendChild(imgDocumento);
					}
				trTipo.appendChild(tdImgDocumento);
		/* USUARIO */	
				var tdUsuario = document.createElement('td');
				tdUsuario.align = 'center';
				tdUsuario.innerHTML = pDsUsuario;
				trTipo.appendChild(tdUsuario);			
				tBody.appendChild(trTipo);
		        tblAutorizaciones.appendChild(tBody);
	}
	
	/* funcion: getTipoImagen(p_tipo,p_IdDocumento,pPuerto)
	 * 			Se encarga de obtener la imagen, el titulo de la misma y la accion a realizar en caso de elegirla para cada Tipo de Documento
	 *	Parametros:  - p_tipo: el tipos de Documento
	 *			     - p_IdDocumento: el Id de Documento
	 *				 - pPuerto : puerto (en caso de que sea un Ajuste Pto, sino es vacio)
	 */
	function getTipoImagen(p_tipo,p_IdDocumento,p_Pto){		
		switch (p_tipo)
			{
			case "<%=AUTH_TYPE_PIC%>":
				imgTipo     = "images/compras/PIC-16x16.png"			  
				titleTipo   = "Ver Pedido Interno de Compra"
				onclickTipo = "abrirCotizacion(" + p_IdDocumento + ")"							  
			break;
			case "<%=AUTH_TYPE_CEC%>":
				imgTipo     = "images/compras/CEC-16x16.png"			  
				titleTipo   = "Ver Comprobante Electronico de Cumplimiento"
				onclickTipo = "abrirCotizacion(" + p_IdDocumento + ")"							  
			break;
			case "<%=AUTH_TYPE_AFE%>":
				imgTipo = "images/compras/AFE-16x16.png"
				titleTipo   = "Ver AFE"
				onclickTipo = "abrirAfe(" + p_IdDocumento + ")"	
			break;
			case "<%=AUTH_TYPE_AJS%>":
				imgTipo = "images/almacenes/AJS-16x16.png"
				titleTipo   = "Ver Ajuste de Stock"
				onclickTipo = "abrirVale(" + p_IdDocumento + ")"	
			break;				
			case "<%=AUTH_TYPE_XJS%>":
				imgTipo = "images/almacenes/XJS-16x16.png"
				titleTipo   = "Ver Anulación de Ajuste de Stock"
				onclickTipo = "abrirVale(" + p_IdDocumento + ")"	
			break;	
			case "<%=AUTH_TYPE_VRS%>":
				imgTipo = "images/almacenes/VRS-16x16.png"
				titleTipo   = "Ver Reclasificación de Stock"
				onclickTipo = "abrirVale(" + p_IdDocumento + ")"		
			break;	
			case "<%=AUTH_TYPE_XRS%>":
				imgTipo = "images/almacenes/XRS-16x16.png"
				titleTipo   = "Ver Anulación de Reclasificación de Stock"
				onclickTipo = "abrirVale(" + p_IdDocumento + ")"		
			break;	
			case "<%=AUTH_TYPE_CTC%>":
				imgTipo = "images/compras/CTC-16x16.png"
				titleTipo   = "Ver Contrato"
				onclickTipo = "abrirContrato(" + p_IdDocumento + ")"
			break;	
			default:
				imgTipo = "";				
			} 
	}			
	function abrirAjustePic(p_IdDocumento){
		window.open("comprasAjustePICFirmas.asp?idCotizacion=" + p_IdDocumento);
	}
	
	
</script>
</head>
<body onLoad="bodyOnLoad()">		
	<form id="Myform" name="Myform">
		<div id="toolbar"></div>		
		<br>	
		<table id="MyTablaAutorizaciones" name="MyTablaAutorizaciones" align="center" width="100%" class="reg_Header">			
			<tr><td colspan="10"><div id="paginacion"></div></td></tr>						
			<tr><td colspan="10"><div id="msjCargando" align="center" class="divAutorizaciones" ></div></td></tr>
			<tr class="reg_Header_nav">
				<td width="25%" rowspan="2" style="text-align: center"><% =GF_TRADUCIR("Que se Autoriza/ Aprueba") %></td>
				<td width="10%" rowspan="2" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Ptda.Presup.") %></td>
				<td width="15%" rowspan="2" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Pedido") %></td>
				<td width="20%" colspan="2"style="text-align: center"><% =GF_TRADUCIR("Adjudicado A") %></td>
				<td width="10%" rowspan="2" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Documento") %></td>
				<td width="20%" rowspan="2" style="text-align: center"><% =GF_TRADUCIR("Usuario a Firmar") %></td>
			</tr>			
		</table>	
	</form>	
</body>
</html>