<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<%
Const ALMACENES     = "Almacenes"
Const DIV_EXPORTACION = 1
Const DIV_ARROYO      = 2
Const DIV_PIEDRABUENA = 3
Const DIV_TRANSITO    = 4
Const PREFIJO_ANULACION = "X"
'------------------------------------------------------------------------------------------------
Function armarSQLAjustePartidaPresupuestaria()
    Dim rs
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACION_GET_CBTES_A_FIRMAR", Session("Usuario") &"||1||0$$totalRegistros" )
    if (not rs.eof) then armarSQLAjustePartidaPresupuestaria = armarLineaTipoDocumento(sp_ret("totalRegistros"),AUTH_TYPE_APP,"")
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjustePuerto(pPto, pTipoAjuste)	
	Call executeProcedureDb(pPto, rs, "TBLAJUSTES_GET_CBTES_A_FIRMAR", Session("Usuario") &"||"& pTipoAjuste)	
    if (not rs.eof) then armarSQLAjustePuerto = armarLineaTipoDocumento(rs.RecordCount,pTipoAjuste,pPto)
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLPlanillas()
	Dim sp_ret, rs, strSector
	
	'1ro determino el sector del usuario
	strSector=getListBossOf(session("Usuario"))
	'2do traigo todas las planillas que tenga para firmar.  
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPCTCABECERA_GET_CBTES_A_FIRMAR", session("Usuario")& "||" & strSector & "||1||0$$totalRegistros")
	if not rs.eof then armarSQLPlanillas = ArmarListaTipoDocumento( rs )	
	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuPics()
	Dim sp_ret, rs, strSector
	
	'1ro determino el sector del usuario
    strSector=getListBossOf(session("Usuario"))
	'2do traigo todas las planillas que tenga para firmar.  
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZAJUSTES_GET_CBTES_A_FIRMAR", session("Usuario") & "||1||1||0$$totalRegistros")		
	if not rs.eof then armarSQLAjuPics = ArmarListaTipoDocumento( rs )	
	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuCEC()
	Dim sp_ret, rs, strSector
	
	'1ro determino el sector del usuario
    strSector=getListBossOf(session("Usuario"))
	'2do traigo todas las planillas que tenga para firmar.  
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZAJUSTES_GET_CBTES_A_FIRMAR", session("Usuario") & "||0||1||0$$totalRegistros")		
	if not rs.eof then armarSQLAjuCEC = ArmarListaTipoDocumento( rs )	
	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAjuCTC()
	Dim sp_ret, rs, strSector
	
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACTCAJUSTES_GET_CBTES_A_FIRMAR", session("Usuario") & "||1||0$$totalRegistros")
	if not rs.eof then armarSQLAjuCTC = ArmarListaTipoDocumento( rs )	
	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLPics()
	Dim sp_ret, rs, strSector
		
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_CBTES_A_FIRMAR", session("Usuario") & "||1||1||0$$totalRegistros")
	if not rs.eof then armarSQLPics = ArmarListaTipoDocumento( rs )	
		
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLCEC()
	Dim sp_ret, rs, strSector
		
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLCTZCABECERA_GET_CBTES_A_FIRMAR", session("Usuario") & "||0||1||0$$totalRegistros")
	if not rs.eof then armarSQLCEC = ArmarListaTipoDocumento( rs )	
		
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLVales(pTipo)

	Dim strSqlBase,conn,rs,rol,secuencia
	Dim myIn,rsRegistros,usrKr,secuencia1,secuencia2
	Dim tipoVale,estado, usuarioEspecial
	
	usuarioEspecial = ""	
	if (gRolFirmaAlmacenes = FIRMA_ROL_AUDITOR) then usuarioEspecial = VS_AUDIT_USER
	if (gRolFirmaAlmacenes = FIRMA_ROL_SUP_PUERTO) then usuarioEspecial = VS_PORT_SUPERVISOR_USER
	if (gRolFirmaAlmacenes = FIRMA_ROL_RESP_PUERTO) then usuarioEspecial = VS_PORT_GERENTE_USER
	if (gRolFirmaAlmacenes = FIRMA_ROL_DIRECTOR) then usuarioEspecial = DIRECTOR_USER
	'Obtengo los permisos del usuario
	strSQL = "select * from TBLREGISTROFIRMAS where CDUSUARIO = '" & session("Usuario") & "'"	
	Call executeQueryDB(DBSITE_SQL_INTRA, rsRegistros, "OPEN", strSQL)
	myIn = "0,"
	if (not rsRegistros.EoF) then
		if (rsRegistros("AJARROYO")      = 1) then myIn = myIn & DIV_ARROYO & ","
		if (rsRegistros("AJTRANSITO")    = 1) then myIn = myIn & DIV_TRANSITO & ","
		if (rsRegistros("AJPIEDRABUENA") = 1) then myIn = myIn & DIV_PIEDRABUENA & ","
		if (rsRegistros("AJEXPORTACION") = 1) then myIn = myIn & DIV_EXPORTACION & ","
	end if
	myIn = left(myIn,len(myIn)-1) ' le saco la ultima coma
	rtrn=""
	if ((Len(myIn) > 1) or (usuarioEspecial <> "")) then		
		'Puede firmar ajustes y sus anulaciones		
		'Determino el estado del comprobante
		tipoVale = "'"& pTipo & "'"
		if (Left(pTipo,1) <> PREFIJO_ANULACION) then
			estado = ESTADO_ACTIVO
		else
			estado = ESTADO_ANULACION			
		end if
		strSqlBase = 		 	  " select vf.idVALE iddocumento, vc.CDVALE tipo, 0 idpedido, vc.IDOBRA IDOBRA, vc.NRVALE CDDOCUMENTO, C.cdobra CDOBRA, 0 AS CDPEDIDO, 0 AS IDPROVEEDOR, '' DSPROVEEDOR, '' PTO  "
		strSqlBase = strSqlBase & " from (Select * from TBLVALESFIRMAS where HKEY is null or HKEY = '') vf "
		strSqlBase = strSqlBase & " inner join (Select * from TBLVALESCABECERA where ESTADO = " & estado & " and CDVALE in (" & tipoVale & ")) vc on vc.IDVALE = vf.IDVALE "
		strSqlBase = strSqlBase & " inner join TBLALMACENES alm on alm.IDALMACEN = vc.IDALMACEN "
		strSqlBase = strSqlBase & " inner join TBLDIVISIONES div on div.IDDIVISION = alm.IDDIVISION "
		strSqlBase = strSqlBase & "	LEFT JOIN TBLDATOSOBRAS C on C.IDOBRA = vc.IDOBRA " 		
		
		'Solo se muestran para firmar los vales que no esten firmados, que el usuario tenga permiso de firma y siempre y cuando no haya realizado alguna de las firmas faltantes.
		strSqlBase = strSqlBase & " where (vf.CDUSUARIO = '"&session("Usuario")&"') "		
		strSqlBase = strSqlBase & " or ("										
		'								Me aseguro que el usuario no figure en niguna otra posición firmada o no.
		strSqlBase = strSqlBase & " 	vf.IDVALE not in (  select distinct(IDVALE) "
		strSqlBase = strSqlBase & " 							from TBLVALESFIRMAS "
		strSqlBase = strSqlBase & " 							where CDUSUARIO = '"&session("Usuario")&"')" 		
		
		strSqlBase = strSqlBase & " 	and ("		
		if (usuarioEspecial = VS_PORT_GERENTE_USER) then 
			strSqlBase = strSqlBase & "             (vf.SECUENCIA = " & VS_FIRMA_GERENTE & " and vf.CDUSUARIO = '"& VS_NO_USER &"' and div.IDDIVISION in ("&myIn&")) "
		else
			if (usuarioEspecial = DIRECTOR_USER) then 
				strSqlBase = strSqlBase & " 		    (vf.SECUENCIA = " & VS_FIRMA_DIRECTOR & " and vf.CDUSUARIO = '"& usuarioEspecial &"')"
			else
				strSqlBase = strSqlBase & " 		    (vf.SECUENCIA = " & VS_FIRMA_COORD_AUDIT & " and vf.CDUSUARIO = '"& usuarioEspecial &"')"
			end if	
		end if
		strSqlBase = strSqlBase & "         ))"
		rtrn = strSqlBase			
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", rtrn)	
		if not rs.eof then armarSQLVales = ArmarListaTipoDocumento( rs )		
	end if	
End Function
'------------------------------------------------------------------------------------------------
Function armarSQLAfes()
	Dim strSQL,isAuditor
	
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLDATOSAFE_GET_CBTES_A_FIRMAR", session("Usuario")&"||1||0$$totalRegistros")
    if not rs.eof then armarSQLAfes = ArmarListaTipoDocumento( rs )

End Function
'------------------------------------------------------------------------------------------------
'armarSQLCierreAlmacenes: Arma el sql de las firmas pendientes de los cierres contables.
'						  Este tipo de documento tiene una página donde muestra todos los cierres juntos(por division)
'						  que se tiene para firmar, es decir que para identificar a un cierre necesito solo la division,
'						  por este motivo en la SQL se asigna al IdDocumento el valor de la division
Function armarSQLCierreAlmacenes()
	Dim strSQL
	if ((gRolFirmaAlmacenes = FIRMA_ROL_RESP_CONTADURIA)or(gRolFirmaAlmacenes = FIRMA_ROL_RESP_PUERTO)) Then
		call executeProcedureDb(DBSITE_SQL_INTRA, rsRegistros, "TBLREGISTROFIRMAS_GET_BY_PARAMETERS", "0||"& Session("Usuario") & "||")
		myIn = "0,"
		if (not rsRegistros.EoF) then
			if (rsRegistros("ASARROYO")      = 1) then myIn = myIn & DIV_ARROYO & ","
			if (rsRegistros("ASTRANSITO")    = 1) then myIn = myIn & DIV_TRANSITO & ","
			if (rsRegistros("ASPIEDRABUENA") = 1) then myIn = myIn & DIV_PIEDRABUENA & ","
			if (rsRegistros("ASEXPORTACION") = 1) then myIn = myIn & DIV_EXPORTACION & ","
		end if
		myIn = left(myIn,len(myIn)-1)
		strSQL = "SELECT DISTINCT cab.IDDIVISION iddocumento, " &_
				 "		 '"& AUTH_TYPE_CCN &"' AS tipo," &_
				 "		 0 as idpedido, " &_
				 "		 0 as IDOBRA, " &_				 
				 "       Cast(DIV.DSDIVISION as varchar(30)) + '-' + Cast(cab.anio as varchar(10)) + '/' + Cast(cab.mes as varchar(10)) CDDOCUMENTO, "&_ 				 
				 "		 '' as CDOBRA, " &_
				 "		 '' as CDPEDIDO, " &_
				 "		 0 as IDPROVEEDOR, " &_
				 "		 '' as DSPROVEEDOR, " &_
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
				 "WHERE cab.ESTADO = '"& TIPO_CIERRE_PROVISORIO &"'" &_
				 "		AND CAB.IDDIVISION IN ("& myIn &") "&_
				 "		AND fir.SECUENCIA = "& gRolFirmaAlmacenes
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
		if not rs.eof then armarSQLCierreAlmacenes = ArmarListaTipoDocumento( rs )
	end if	
End Function
'------------------------------------------------------------------------------------------------
'Funcion: ArmarListaTipoDocumento(pRs, pTipo) : Se encarga de ir armando una lista de los registro del recordset, cada campo esta separado
'									   por "|" y cada registro es separado por ";"
'Parametros: 
'		    - pRs:   es el recordset de la SQL
'			- pTipo: es el Tipo de Documento 
'Fecha: 12/09/2012   -  CNA
Function ArmarListaTipoDocumento(pRs)
	Dim	listOfDocumento, cdPedido, cdObra, idProveedor, dsProveedor, idPedido, idObra		
	
	While (not pRs.eof) 	    
		cdPedido = getRSValue(pRs, "CDPEDIDO", "Sin Pedido")		
		cdObra = getRSValue(pRs, "CDOBRA", "Sin Obra")		
		idProveedor = getRSValue(pRs, "IDPROVEEDOR", 0)		
		idPedido = getRSValue(pRs, "IDPEDIDO", 0)		
		idObra = getRSValue(pRs, "IDOBRA", 0)		
		dsProveedor = getDescripcionProveedor(idProveedor)
		listOfDocumento = listOfDocumento & Trim(pRs("TIPO")) & "|" & Trim(cdObra) & "|" & idObra & "|" & Trim(cdPedido) &"|" & idPedido &"|" & idProveedor  &"|" & Trim(dsProveedor) &"|" & pRs("IDDOCUMENTO") &"|" & Trim(pRs("CDDOCUMENTO")) &"|" & Trim(getDSDocumentoFirmar(pRs("TIPO"))) & "|"&  pRs("PTO") & "|0;"
		pRs.MoveNext()	
	wend
	ArmarListaTipoDocumento = listOfDocumento	
	
End function
'------------------------------------------------------------------------------------------------
Function armarLineaTipoDocumento(pCantidad,pTipo,pPto)	
	armarLineaTipoDocumento = pTipo & "||0||0|0||0||" & Trim(getDSDocumentoFirmar(Cstr(pTipo))) &"|"& pPto &"|"& pCantidad &";"
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
'---------------------------------------------------------------------------------------------
'Función:	
'				getDocumentoFirmarDireccion	
' Autor: 	
'				CNA - Ajaya Nahuel
' Fecha: 	
'				26/05/2014
' Objetivo:
'				Devuelve todos los tipos de documentos que puede tener la Direccion para firmar
' Parametros:
'				-
' Devuelve:
'				Array con los Tipos de Documentos 
'--------------------------------------------------------------------------------------------------
Function getDocumentoFirmarDireccion()
	Dim vTipoDocumentoDireccion()
	Redim vTipoDocumentoDireccion(4)
	vTipoDocumentoDireccion(0) = AUTH_TYPE_PCP
	vTipoDocumentoDireccion(1) = AUTH_TYPE_PIC
	vTipoDocumentoDireccion(2) = AUTH_TYPE_AIC
	vTipoDocumentoDireccion(3) = AUTH_TYPE_AFE
	getDocumentoFirmarDireccion = vTipoDocumentoDireccion
End Function
'------------------------------------------------------------------------------------------------
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim rsFirmasPendientes, rs, conn, strSQL, cdObra, paginaActual, mostrar, idProv, dsProv, origen,pTipo,vTipoDocumento,vTipoDocumentoDireccion
Dim gRolFirmaCompras, gRolFirmaAlmacenes, gRolFirmaProveedores, gRolFirmaPoseidon

origen = GF_PARAMETROS7("origen","",6)
call addParam("origen", origen, params)
pTipo = GF_PARAMETROS7("Tipo","",6)
call addParam("Tipo", pTipo, params)
accion = GF_PARAMETROS7("accion","",6)
call addParam("accion", accion, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10


GP_ConfigurarMomentos

If(accion = ACCION_PROCESAR)then
    listOfDocumento = ""
    if (session("Usuario") <> "") then	
        gRolFirmaCompras = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
        gRolFirmaAlmacenes = getRolFirma(session("Usuario"), SEC_SYS_ALMACENES)
        gRolFirmaProveedores = getRolFirma(session("Usuario"), SEC_SYS_PROVEEDORES)
        gRolFirmaPoseidon = getRolFirma(session("Usuario"), SEC_SYS_POSEIDON)		
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
			case AUTH_TYPE_AEC: '  AJS PIC	'	
			    listOfDocumento = armarSQLAjuCEC()
		    case AUTH_TYPE_AFE: ' AFE	'
			    listOfDocumento = armarSQLAfes()
		    case AUTH_TYPE_VRS: ' RECASIFICACION STOCK , RECASIFICACION STOCK ANULACION,AJUSTE STOCK , AJUSTE STOCK ANULACION '			
			    listOfDocumento = armarSQLVales(CODIGO_VS_RECLASIFICACION_STOCK)
			    'listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_RECLASIFICACION_STOCK_X)
			    listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_AJUSTE_STOCK)
			    'listOfDocumento = listOfDocumento & armarSQLVales(CODIGO_VS_AJUSTE_STOCK_X)
			case AUTH_TYPE_AJD: ' AJS DRAFT SURVEY (ARROYO, TRANSITO, PIEDRABUENA)'
                listOfDocumento = armarSQLAjustePuerto(TERMINAL_ARROYO, AUTH_TYPE_AJD)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_TRANSITO, AUTH_TYPE_AJD)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_PIEDRABUENA, AUTH_TYPE_AJD)
			case AUTH_TYPE_AJC: ' AJS CALIDAD (ARROYO, TRANSITO, PIEDRABUENA)'
				listOfDocumento = armarSQLAjustePuerto(TERMINAL_ARROYO, AUTH_TYPE_AJC)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_TRANSITO, AUTH_TYPE_AJC)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_PIEDRABUENA, AUTH_TYPE_AJC)
			case AUTH_TYPE_AJM: ' AJS MANIPULEO (ARROYO, TRANSITO, PIEDRABUENA)'
				listOfDocumento = armarSQLAjustePuerto(TERMINAL_ARROYO, AUTH_TYPE_AJM)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_TRANSITO, AUTH_TYPE_AJM)
                listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_PIEDRABUENA, AUTH_TYPE_AJM)
			case AUTH_TYPE_CCN: 
				listOfDocumento = armarSQLCierreAlmacenes()
            case AUTH_TYPE_AJV: ' AJS MERMA VOLATIL (ARROYO, TRANSITO, PIEDRABUENA)'
				listOfDocumento = armarSQLAjustePuerto(TERMINAL_ARROYO, AUTH_TYPE_AJV)
				listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_TRANSITO, AUTH_TYPE_AJV)
				listOfDocumento = listOfDocumento & armarSQLAjustePuerto(TERMINAL_PIEDRABUENA, AUTH_TYPE_AJV)	
            case AUTH_TYPE_APP: 'Ajuste Partida Presupuestaria
                listOfDocumento = armarSQLAjustePartidaPresupuestaria()
	    end select
	end if
	if (Len(listOfDocumento) > 0) then listOfDocumento = left(listOfDocumento,Len(listOfDocumento)-1)
	Response.Write listOfDocumento
	response.end
else
	vTipoDocumento = getDocumentoFirmar()
	'Si es un Coordinador de Puertos se cargará un nuevo vector con los tipos de documentos que tiene para firmar la Direccion
	if (getRolFirma(session("Usuario"), SEC_SYS_COMPRAS) = FIRMA_ROL_GTE_COMPRAS) then vTipoDocumentoDireccion = getDocumentoFirmarDireccion()
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Autorizaciones Pendientes</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css">
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
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/framework.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">

	var ch = new channel();	
	var imgTipo;
	var onclickTipo;
	var titleTipo;	
	var registrosAcumulados = 0;	
	var arrTipoDomuento = new Array();
	var arrTipoDomuentoDireccion = new Array();
	var vRegs;
	var vTipoDocumento = new Array();
	var vTipoDocumentoDireccion = new Array();
	function bodyOnLoad() {		
		vRegs = new Paginacion("paginacion");
		var tb = new Toolbar('toolbar', 5, 'images/compras/');
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		tb.addButtonREFRESH("Recargar", "refresh()");		
		<% if (origen <> ALMACENES) then%>
			tb.addButton("Quote_purchase-16x16.png", "Ped. Precio", "irPedidos()");		
			tb.addButton("direct_Purchase-16x16.png", "Compra Directa", "irDirecta()");
		<% end if %>
		tb.addButton("see_all-16x16.png", "Todas las Autorizaciones", "irMaestro()");
		tb.draw();				
		//Cargo el vector con los tipos de documento.
		<% for i=0 to UBound(vTipoDocumento)-1 %>
			vTipoDocumento.push('<% =vTipoDocumento(i) %>');
		<% next %>		
		//Hago el primer llamado para iniciar la cadena de carga
		var tipo = vTipoDocumento.pop();		
		document.getElementById("msjCargando").innerHTML= getDescripcionTipo(tipo);		
		ch.bind("comprasAutorizaciones.asp?Tipo=" + tipo + "&accion=<%=ACCION_PROCESAR%>","CallBack_getAutorizaciones()");
		ch.send();		
	}
		
	var hayDatos = false;	
	var hayDatosDireccion = false;		
	
	function MostrarListaTipoDocumento(isLast) {
		while (arrTipoDomuento.length > 0) {
			var linea = String(arrTipoDomuento.splice(0,1));			
			var vals = linea.split("|");
			hayDatos = true;
			agregarLineaAutorizaciones(vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],vals[6],vals[7],vals[8],vals[9],vals[10],"",vals[11]);
		}
		document.getElementById("msjCargando").innerHTML="";		
	}
		
	
	function CallBack_getAutorizaciones(){
		var isLast = false;		
		//Lanzo la busqueda del siguiente tipo de documento		
		if (vTipoDocumento.length > 0) {
			var tipo = vTipoDocumento.pop();
			document.getElementById("msjCargando").innerHTML= getDescripcionTipo(tipo);				
			ch.bind("comprasAutorizaciones.asp?Tipo=" +  tipo + "&accion=<%=ACCION_PROCESAR%>","CallBack_getAutorizaciones()");
			ch.send();
		} else {
			isLast = true;
			document.getElementById("msjCargando").innerHTML="";			
			<% if (getRolFirma(session("Usuario"), SEC_SYS_COMPRAS) = FIRMA_ROL_GTE_COMPRAS) then %>		
			<%    for i=0 to UBound(vTipoDocumentoDireccion)-1 %>
					vTipoDocumentoDireccion.push('<% =vTipoDocumentoDireccion(i) %>');					
			<%	  next %>		
				  var tipoDir = vTipoDocumentoDireccion.pop();
				  document.getElementById("msjCargandoPendiente").innerHTML= getDescripcionTipo(tipoDir);
				  ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" + tipoDir + "&accion=<%=ACCION_PROCESAR%>&origen=2","getAutorizacionesDireccion_CallBack()");
				  ch.send();
			<% end if %>
		} 		
		//Proceso los resultados obtenidos del servidor		
		var rtrn = ch.response();		
		if (rtrn.length > 0) {		
			var arr = rtrn.split(";");
			arrTipoDomuento = arrTipoDomuento.concat(arr);
			MostrarListaTipoDocumento(isLast);
		}
		if ((isLast) && (!hayDatos)) {
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
		
	
	function getDescripcionTipo(pTipo){		
		var dsTipo;
		if(pTipo == "<%=AUTH_TYPE_PCP%>") dsTipo =  "Cargando... PCP";	
		if(pTipo == "<%=AUTH_TYPE_CTC%>") dsTipo =  "Cargando... Contratos";	
		if(pTipo == "<%=AUTH_TYPE_PIC%>") dsTipo =  "Cargando... PIC";
		if(pTipo == "<%=AUTH_TYPE_CEC%>") dsTipo =  "Cargando... CEC";
		if(pTipo == "<%=AUTH_TYPE_AIC%>") dsTipo =  "Cargando... Ajuste de PIC";
		if(pTipo == "<%=AUTH_TYPE_AEC%>") dsTipo =  "Cargando... Ajuste de CEC";
		if(pTipo == "<%=AUTH_TYPE_AFE%>") dsTipo =  "Cargando... AFE";
		if(pTipo == "<%=AUTH_TYPE_VRS%>") dsTipo =  "Cargando... Vales";
		if(pTipo == "<%=AUTH_TYPE_AJD%>") dsTipo =  "Cargando... Ajuste de Draft Survey";
		if(pTipo == "<%=AUTH_TYPE_AJC%>") dsTipo =  "Cargando... Ajuste de Calidad";
		if(pTipo == "<%=AUTH_TYPE_AJM%>") dsTipo =  "Cargando... Ajuste de Manipuleo";
		if(pTipo == "<%=AUTH_TYPE_CCN%>") dsTipo =  "Cargando... Cierres de Almacenes";
		if(pTipo == "<%=AUTH_TYPE_PVS%>") dsTipo =  "Cargando... Provisiones";
		if(pTipo == "<%=AUTH_TYPE_AJV%>") dsTipo =  "Cargando... Merma Volatil";
		if(pTipo == "<%=AUTH_TYPE_APP%>") dsTipo =  "Cargando... Ajuste Partida Presupuestaria";
		return dsTipo;
	}
	
	/* funcion:    agregarLineaAutorizaciones()
	 *			   Se encarga de crear las lineas de la tabla con sus respectivos registros 
	 * Parametros: Todos los paramtros pasadosson el resultado de la consulta SQL traida por ajax, cada uno se va a mostrar en la tabla
	 */
	function agregarLineaAutorizaciones(pTipo, pCdObra, pIdObra, pCdPedido ,pIdPedido ,pIdProvEleg ,pDsProvEleg ,pIdDocumento ,pCdDocumento,pDSDocumentoFirmar, pPto,pUserDireccion,pCantidadDocumentos) {
		if (pUserDireccion != ""){
			var tblAutorizaciones = document.getElementById("MyTablaAutorizacionesPendientes");
		}
		else{
			var tblAutorizaciones = document.getElementById("MyTablaAutorizaciones");
		}
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
				setEvent(tdTipo, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");
				trTipo.appendChild(tdTipo);			
		/* OBRA */
				var tdObra  = document.createElement('td');
				tdObra.align = 'center';		
				if(pIdObra > 0) tdObra.innerHTML = pCdObra;								
				setEvent(tdObra, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");
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
				setEvent(tdPedido, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");				
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
				setEvent(tdProveedor, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");
				var strProveedor = "Sin Datos";
				if(pIdProvEleg > 0) strProveedor = pIdProvEleg +" - "+ pDsProvEleg;
				tdProveedor.innerHTML = strProveedor;
				trTipo.appendChild(tdProveedor);	
		/* IMAGEN PROVEEDOR */
				var tdImgProveedor  = document.createElement('td');
				tdImgProveedor.align = 'center';
				if(pIdPedido > 0)
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
				setEvent(tdDocumento, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");
				var strDocumento = pIdDocumento;
				if (pCantidadDocumentos == 0)	
					tdDocumento.innerHTML = pCdDocumento;					
				else
					tdDocumento.innerHTML = "Pendientes: "+ pCantidadDocumentos;
						
				trTipo.appendChild(tdDocumento);	
		/* FIRMA O USUARIO DIRECCION */					
			if(pUserDireccion == ""){
				var tdImgFirma = document.createElement('td');
				tdImgFirma.align = 'center';				
				var imgFirma = document.createElement("img");
				imgFirma.src   = "images/compras/Authorize-16x16.png";			
				imgFirma.title = "Firmar";				
				setEvent(imgFirma, FWRK_EVT_ON_CLICK, "firmar('" + pTipo + "', " + pIdDocumento + ",'" + pPto + "')");
				tdImgFirma.appendChild(imgFirma);			
				trTipo.appendChild(tdImgFirma);		
			}
			else{
				var tdUserDireccion  = document.createElement('td');
				tdUserDireccion.align = 'center';
				tdUserDireccion.innerHTML = pUserDireccion;
				trTipo.appendChild(tdUserDireccion);
			}			
			tBody.appendChild(trTipo);
		    tblAutorizaciones.appendChild(tBody);
	}
	
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
		location.href = "comprasAutorizaciones.asp"
	}
	function irMaestro() {
		location.href = "comprasAutorizacionesMaestro.asp"
	}
	function abrirVale(id) {
		window.open("almacenValePedidoPrint.asp?idVale=" + id, "_blank", "location=no,menubar=no,statusbar=no",false);
	}
	function abrirTarea(p_IdDocumento){
		window.open("geminiIssueFirma.asp?idTarea="+p_IdDocumento, "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
	}	
	function abrirAfe(id) {
		window.open("comprasAFEPrint.asp?idafe=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
	}
	function abrirContrato(id) {
		window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,statusbar=no",false);
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
	function abrirAutorizacionesProveedor(id){	
		alert("Abrir")
	}
	function firmar(tipo, id, pto) {
		switch	(tipo) {			
			case "<% =AUTH_TYPE_PCP %>":
				window.open("comprasComparativoDeOfertasFirma.asp?idPedido=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
				break;
			case "<% =AUTH_TYPE_PIC %>":
				window.open("comprasPICFirma.asp?idCotizacionElegida=" + id);
				break;
			case "<%=AUTH_TYPE_CEC%>":
				window.open("comprasPICFirma.asp?idCotizacionElegida=" + id);
			break;
			case "<% =AUTH_TYPE_AFE %>":
				window.open("comprasAFEFirma.asp?idAFE=" + id, "_blank");
				break;
			case "<% =AUTH_TYPE_AJS %>":
				window.open("almacenValesFirma.asp?idVale=" + id +"&tipo=<%=AUTH_TYPE_AJS%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);				
				break;
			case "<% =AUTH_TYPE_AIC %>":
				window.open("comprasAjustePICFirmas.asp?idAjuste=" + id);
				break;
			case "<% =AUTH_TYPE_AEC %>":
				window.open("comprasAjustePICFirmas.asp?idAjuste=" + id);
				break;
			case "<% =AUTH_TYPE_XJS %>":
				window.open("almacenValesFirma.asp?idVale=" + id +"&tipo=<%=AUTH_TYPE_XJS%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);				
				break;
			case "<% =AUTH_TYPE_VRS %>":
				window.open("almacenValesFirma.asp?idVale=" + id +"&tipo=<%=AUTH_TYPE_VRS%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);				
				break;
			case "<% =AUTH_TYPE_XRS %>":
				window.open("almacenValesFirma.asp?idVale=" + id +"&tipo=<%=AUTH_TYPE_XRS%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);			
				break;
			case "<% =AUTH_TYPE_CTC %>":
				window.open("comprasAjusteCTCFirmas.asp?idAjuste=" + id);
				break;
			case "<% =AUTH_TYPE_AJD %>":				
			    window.open("poseidonAjusteFirma.asp?idAjuste=" + id + "&Pto="+pto, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
				break;		
			case "<% =AUTH_TYPE_AJC %>":				
			    window.open("poseidonAjusteFirma.asp?idAjuste=" + id + "&Pto="+pto, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
				break;		
			case "<% =AUTH_TYPE_AJM %>":				
			    window.open("poseidonAjusteFirma.asp?idAjuste=" + id + "&Pto="+pto, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
				break;				
			case "<% =AUTH_TYPE_CCN %>":
				window.open("almacenCCN_ConsultasContables.asp?idDivision=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes",false);
				break;
		    case "<% =AUTH_TYPE_AJV %>":
		        window.open("poseidonAjusteFirma.asp?pto="+pto, "_blank", "type=fullWndow,resizable=yes,scrollbars=1,height=768,width=1024",false);
		        break;
		    case "<% =AUTH_TYPE_APP %>":
		        window.open("comprasBudgetFirma.asp", "_blank", "type=fullWndow,resizable=yes,scrollbars=1,height=768,width=1224",false);
		        break;
		}
	}		
	
	function getAutorizacionesDireccion_CallBack(){			
		var isLast = false;				
		var rtrn = ch.response();
		if (vTipoDocumentoDireccion.length > 0) {			
			var tipoDir = vTipoDocumentoDireccion.pop();
			document.getElementById("msjCargandoPendiente").innerHTML= getDescripcionTipo(tipoDir);
			ch.bind("comprasAutorizacionesMaestro.asp?Tipo=" +  tipoDir + "&accion=<%=ACCION_PROCESAR%>&origen=2","getAutorizacionesDireccion_CallBack()");
			ch.send();
		}
		else{
			isLast = true;
			document.getElementById("msjCargandoPendiente").innerHTML="";			
		}		
		if (rtrn.length > 0) {		
			var arr = rtrn.split(";");
			arrTipoDomuentoDireccion = arrTipoDomuentoDireccion.concat(arr);			
			MostrarListaTipoDocumentoDireccion(isLast);
		}		
		if ((isLast) && (!hayDatosDireccion)) {
			var tblAutorizaciones = document.getElementById("MyTablaAutorizacionesPendientes");
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
	function MostrarListaTipoDocumentoDireccion(isLast) {
		while (arrTipoDomuentoDireccion.length > 0) {
			var linea = String(arrTipoDomuentoDireccion.splice(0,1));			
			var vals = linea.split("|");			
			hayDatosDireccion = true;
			agregarLineaAutorizaciones(vals[0],vals[1],vals[2],vals[3] ,vals[4] ,vals[5]	,vals[6]	,vals[7]	  ,vals[8]		,vals[9],vals[11],vals[10], 0);			                          
			//COMPRAS AUTORIZACIONES
			/*						   pRs("TIPO"),cdObra  ,idObra  ,cdPedido , idPedido ,idProveedor ,dsProveedor ,pRs("IDDOCUMENTO") ,pRs("CDDOCUMENTO"),getDSDocumentoFirmar("TIPO"),pRs("PTO") 		
		    agregarLineaAutorizaciones(vals[0]	  ,vals[1] ,vals[2] ,vals[3]  ,vals[4]   ,vals[5]     ,vals[6]     ,vals[7]			   ,vals[8]		      ,vals[9]					   ,vals[10]	,false				  );
		    //COMPRAS AUTORIZACIONES MAESTRO
			agregarLineaAutorizaciones(vals[0]	  ,vals[1] ,vals[2] ,vals[3]  ,vals[4]   ,vals[5]     ,vals[6]     ,vals[7]			   ,vals[8]		      ,vals[9]					   ,vals[11]	,true				  );			
			//FUNCION
			agregarLineaAutorizaciones(pTipo	  ,pCdObra ,pIdObra ,pCdPedido,pIdPedido ,pIdProvEleg ,pDsProvEleg ,pIdDocumento	   ,pCdDocumento	  ,pDSDocumentoFirmar		   ,pPto		,pIsDireccion		  );													
			*/
		}
		document.getElementById("msjCargando").innerHTML="";		
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">	
	<form  id="Myform" name="Myform">
		<input type="hidden" name="accion" id="accion" value="<%= accion%>">
		<div id="toolbar"></div>		
		<br>
		<table id="MyTablaAutorizaciones" name="MyTablaAutorizaciones" align="center" width="90%" class="reg_Header">			
			<tr><td colspan="10"><div id="paginacion"></div></td></tr>										
			<tr><td colspan="10"><div id="msjCargando" align="center" class="divAutorizaciones" ></div></td></tr>		
				
			<tr class="reg_Header_nav">
				<td width="20%" style="text-align: center"><% =GF_TRADUCIR("Que se Autoriza/ Aprueba") %></td>
				<td width="15%" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Ptda.Presup.") %></td>
				<td width="15%" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Pedido") %></td>				
				<td width="20%" colspan="2"style="text-align: center"><% =GF_TRADUCIR("Adjudicado A") %></td>
				<td width="10%" style="text-align: center"><% =GF_TRADUCIR("Documento") %></td>		
				<td width="3%" style="text-align: center"><% =GF_TRADUCIR("Firma") %></td>
			</tr>								
		</table>	
		<% if (getRolFirma(session("Usuario"), SEC_SYS_COMPRAS) = FIRMA_ROL_GTE_COMPRAS) then %>		
			<br></br>
			<table id="MyTablaAutorizacionesPendientes" name="MyTablaAutorizacionesPendientes" align="center" width="90%" class="reg_Header">
				<tr><td colspan="10"><div id="msjCargandoPendiente" align="center" class="divAutorizaciones" ></div></td></tr>
				<tr class="reg_Header_nav"><td colspan="10"><%= GF_TRADUCIR("Pendientes de Dirección")%></td></tr>
				<tr class="reg_Header_nav">
					<td width="20%" style="text-align: center"><% =GF_TRADUCIR("Que se Autoriza/ Aprueba") %></td>
					<td width="15%" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Ptda.Presup.") %></td>
					<td width="15%" colspan="2" style="text-align: center"><% =GF_TRADUCIR("Pedido") %></td>
					<td width="15%" colspan="2"style="text-align: center"><% =GF_TRADUCIR("Adjudicado A") %></td>
					<td width="10%" style="text-align: center"><% =GF_TRADUCIR("Documento") %></td>
					<td width="15%" style="text-align: center"><% =GF_TRADUCIR("Usuario") %></td>
				</tr>
			</table>
		<% end if %>
	</form>	
</body>
</html>
