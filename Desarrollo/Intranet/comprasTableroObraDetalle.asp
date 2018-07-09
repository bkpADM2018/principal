<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<%

Const TODOS_LOS_MOVIMIENTOS = 0
Const SIN_OBRA = 0	

'-----------------------------------------------------------------------------------------------
' Funcion: obtenerListaEntradas
' Descripcion:
' Esta funcion se encarga de procesar los datos del formulario para mostrar el detalle
' de movimientos.
' Devuelve un RecordSet con los datos.
'-----------------------------------------------------------------------------------------------
Function obtenerListaEntradas(idObra, cdMoneda, pagina, regXpag, pIdArea, pIdDetalle)

	Dim sqlCTZ_P, sqlCTZ_D,sqlAFE,sqlFAC, sqlCTC
	Dim filtroFAC,groupFAC, filtroCTC
	Dim strSQL, rs, myWhere, firstRecord, conn
	Dim filtroPIC,camposAreaDetalle,camposAreaDetalle2,camposAreaDetalle3, camposAreaDetalle4,myGroup
	Dim filtroVales,campoAreaDetalleVales, groupVales
	
	filtroPIC = ""
	filtroCTC = ""
	filtroVales = ""
	myGroup   = ""
	groupFAC  = ""
	groupVales = ""
	camposAreaDetalle  = ""
	camposAreaDetalle2 = ""
	camposAreaDetalle3 = ""
	
	'Verifico si hay que buscar por Area y/o Detalle
	if (pIdArea <> "0") and (pIdArea <> "") then		
		filtroPIC = " and IDAREA = " & pIdArea
		filtroCTC = " and P.IDAREA = " & pIdArea
		filtroFAC = " acd7.IDAREA = " & pIdArea
		filtrosVale = " and cab.idbudgetarea = " & pIdArea
		camposAreaDetalle  = " ctzDet.idArea AS idArea, "
		camposAreaDetalle2 = " " & BUDGET_SIN_AREA & " AS idArea, "
		camposAreaDetalle3 = " acd7.idArea,"
		camposAreaDetalle4 = " P.IDAREA as idArea, "
		campoAreaDetalleVales = " cab.idbudgetarea AS idArea, "
		myGroup = ",ctzDet.idArea "
		groupFAC = ",acd7.idArea "
		groupVales = ", cab.idbudgetarea"
		if (pIdDetalle <> "0") and (pIdDetalle <> "") then			
			filtroPIC = filtroPIC & " and IDDETALLE = " & pIdDetalle
			filtroCTC = filtroCTC & " and P.IDDETALLE = " & pIdDetalle
			filtroFAC = filtroFAC & " and acd7.IDDETALLE = " & pIdDetalle	
			filtrosVale = filtrosVale &" and cab.IDBUDGETDETALLE = " & pIdDetalle			
			camposAreaDetalle = camposAreaDetalle & " ctzDet.idDetalle AS idDetalle, "
			camposAreaDetalle2 = camposAreaDetalle2 & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
			camposAreaDetalle3 = camposAreaDetalle3 & " acd7.IDDETALLE AS idDetalle, "
			camposAreaDetalle4 = camposAreaDetalle4 & " P.IDDETALLE, "
			campoAreaDetalleVales = campoAreaDetalleVales & " cab.IDBUDGETDETALLE as idDetalle, "
			myGroup = myGroup & ",ctzDet.idDetalle "
			groupFAC = groupFAC & ",acd7.IDDETALLE "
			groupVales = groupVales & ",cab.IDBUDGETDETALLE "
		else
			camposAreaDetalle  = camposAreaDetalle  & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
			camposAreaDetalle2 = camposAreaDetalle2 & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
			camposAreaDetalle3 = camposAreaDetalle3 & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
			camposAreaDetalle4 = camposAreaDetalle4 & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
			campoAreaDetalleVales = campoAreaDetalleVales & " " & BUDGET_SIN_DETALLES & " AS idDetalle, "
		end if
	else
		camposAreaDetalle  = " " & BUDGET_SIN_AREA & " AS idArea, " & BUDGET_SIN_AREA & " AS idDetalle, "
		camposAreaDetalle2 = camposAreaDetalle
		camposAreaDetalle3 = camposAreaDetalle
		camposAreaDetalle4 = camposAreaDetalle
		campoAreaDetalleVales = camposAreaDetalle
	end if
	
	tipoCambio = getTipoCambio(MONEDA_DOLAR,"")
	
	sqlCTZ_P = armarSqlPIC_P(CamposAreaDetalle,IdObra,FiltroPIC,myGroup, cdMoneda, tipoCambio)
	sqlCTZ_D = armarSqlPIC_D(CamposAreaDetalle,IdObra,FiltroPIC,myGroup, cdMoneda, tipoCambio)
	sqlFAC = armarSqlFAC(cdMoneda,CamposAreaDetalle3,IdObra,filtroFAC,groupFAC)
	sqlVALE = armarSqlVales(cdMoneda,campoAreaDetalleVales,idObra,filtrosVale,groupVales)
	sqlCTC = armarSqlCTC(cdMoneda,camposAreaDetalle4,idObra,FiltroCTC, tipoCambio)
	
	strSQL = armarSqlGeneral(sqlCTZ_P, sqlCTZ_D, sqlVALE, SqlFAC, sqlCTC)
	
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	
	call logDebug("SQL NUEVA :" & strSQL)
	
	Set obtenerListaEntradas = rs
	
End Function
'-----------------------------------------------------------------------------------------------
Function armarSqlGeneral(pSqlCTZ_P, pSqlCTZ_D, pSqlVale, pSqlFAC, pSqlCTC)
	Dim strSQL
	strSQL = "SELECT t.id,t.idobra,t.cbte,t.tipo,t.codigo, t.idProveedor,t.idpedidoasociado,t.importeVales, t.importePIC, t.importeFAC,t.fecha,t.momento,t.idarea,t.iddetalle,t.habilitado FROM ("
	strSQL = strSQL & pSqlCTZ_P	
	strSQL = strSQL & " UNION "
	strSQL = strSQL & pSqlCTZ_D
	strSQL = strSQL & " UNION "
	strSQL = strSQL & pSqlVale
	strSQL = strSQL & " UNION "
	strSQL = strSQL & pSqlFAC
	strSQL = strSQL & " UNION "
	strSQL = strSQL & pSqlCTC
	strSQL = strSQL & "     ) t         "
	strSQL = strSQL & "ORDER BY fecha "
	'call logdebug(strSQL)
	armarSqlGeneral = strSQL
End Function
'-----------------------------------------------------------------------------------------------
'Funcion que selecciona los contratos a mostrar en la cta corriente de la obra.
'Autor: Javier A. Scalisi
'Fecha: 08/11/2011
Function armarSqlCTC(pMoneda,pCamposAreaDetalle,pIdObra,pFiltroCTC, tipoCambio)
	Dim strSQL		
	
	strSQL = strSQL & "         SELECT P.idcontrato                            AS id              , "
	strSQL = strSQL & "                P.idobra                                AS idobra          , "
	strSQL = strSQL & "                ''                                      AS cbte            , "
	strSQL = strSQL & "                'CTC'                                   AS tipo            , "
	strSQL = strSQL & "                ctc.cdcontrato                          AS codigo          , "
	strSQL = strSQL & "                ctc.idproveedor                         AS idproveedor	  , "
	strSQL = strSQL & "                ctc.idpedido                            AS idpedidoAsociado, "
	strSQL = strSQL & "				   0									   AS importeVales    , "	
	strSQL = strSQL & "                CASE when P.CDMONEDA='" & pMoneda & "' then (P.IMPORTEASIGNADO-P.IMPORTEGASTADO) when P.CDMONEDA='" & MONEDA_PESO & "' then (P.IMPORTEASIGNADO-P.IMPORTEGASTADO) / " & tipoCambio & " else (P.IMPORTEASIGNADO-P.IMPORTEGASTADO) * " & tipoCambio & " end AS importePIC      , "
	strSQL = strSQL & "				   0									   AS importeFAC	  , "	
	strSQL = strSQL & "                FECHAINICIO                             AS fecha           , "
	strSQL = strSQL & "                ctc.mmtoconf                            AS momento         , "
	strSQL = strSQL & pCamposAreaDetalle
	strSQL = strSQL & "                convert(varchar, ctc.estado)            AS habilitado "
	strSQL = strSQL & "         FROM TBLOBRACONTRATOS CTC "
	strSQL = strSQL & "         INNER JOIN TBLCTCPARTIDAS P on CTC.IDCONTRATO=P.IDCONTRATO "
	strSQL = strSQL & "         where CTC.ESTADO not in (" & ESTADO_CTC_CANCELADO & ") and P.IDOBRA=" & pIdObra
	strSQL = strSQL & pFiltroCTC
	armarSqlCTC = strSQL
End Function
'-----------------------------------------------------------------------------------------------
'Se arma la SQL que trae todos los PIC en Pesos, menos aquellos que estan relacionados a un contrato.
'Los que estan relacionados a un contrato no se muestran dado que se muestra directamente el contrato.
'Autor: Javier A. Scalisi
Function armarSqlPIC_P(pCamposAreaDetalle,pIdObra,pFiltroPIC,pGroup, pMonedaConsulta,tc)
	Dim strSQL		
                
	strSQL = strSQL & "         SELECT ctz.idcotizacion                        AS id              , "
	strSQL = strSQL & "                ctz.idobra                              AS idobra          , "
	strSQL = strSQL & "                ''                                      AS cbte            , "
	strSQL = strSQL & "                ctz.tipo                                AS tipo            , "
	strSQL = strSQL & "                convert(varchar(20), ctz.idcotizacion)  AS codigo          , "
	strSQL = strSQL & "                ctz.idproveedor                         AS idproveedor	  , "
	strSQL = strSQL & "                ctz.idpedido                            AS idpedidoAsociado, "
	strSQL = strSQL & "				   0									   AS importeVales    , "	
	if (pMonedaConsulta = MONEDA_PESOS) then
	    strSQL = strSQL & "                Sum(ctzdet.importepesos)            AS importePIC      , "
	else
	    strSQL = strSQL & "                Sum(ctzdet.importedolaresfacturado + ((ctzdet.importepesos - ctzdet.importepesosfacturado )  / " & tc & " ))  AS importePIC      , "
    end if	    
	strSQL = strSQL & "				   0									   AS importeFAC	  , "	
	strSQL = strSQL & "                ctz.fechaentrega                        AS fecha           , "
	strSQL = strSQL & "                ctz.momento                             AS momento         , "
	strSQL = strSQL & pCamposAreaDetalle
	strSQL = strSQL & "                ctz.estado                              AS habilitado "
	'MODIFICACION: A LA TABLA CTZCABECERA LE AGREGO UN FILTRO PARA QUE BUSQUE SOLO LOS PIC QUE NO TENGAN CONTRATOS
	strSQL = strSQL & "         FROM   (Select case when idcontrato = 0 then 'PIC' else 'CEC' end tipo, c.* from tblctzcabecera c where idobra = " & pIdObra & " and CDMONEDA='" & MONEDA_PESOS & "') ctz"
	strSQL = strSQL & "                INNER JOIN tblctzdetalle ctzDet ON ctz.idcotizacion = ctzDet.idcotizacion"	
	strSQL = strSQL & " WHERE  ctz.ESTADO <> '" & CTZ_ANULADA & "' " & pFiltroPIC		
	strSQL = strSQL & "         GROUP BY ctz.idcotizacion,ctz.idobra ,ctz.idcotizacion, ctz.tipo,ctz.idproveedor, ctz.idpedido,ctz.fechaentrega,ctz.momento,ctz.estado" & pGroup
	
	'call logdebug(strSQL)
	armarSqlPIC_P = strSQL
End Function
'-----------------------------------------------------------------------------------------------
'Se arma la SQL que trae todos los PIC en Dolares, menos aquellos que estan relacionados a un contrato.
'Los que estan relacionados a un contrato no se muestran dado que se muestra directamente el contrato.
'Autor: Javier A. Scalisi
Function armarSqlPIC_D(pCamposAreaDetalle,pIdObra,pFiltroPIC,pGroup, pMonedaConsulta,tc)
	Dim strSQL		
                
	strSQL = strSQL & "         SELECT ctz.idcotizacion                        AS id              , "
	strSQL = strSQL & "                ctz.idobra                              AS idobra          , "
	strSQL = strSQL & "                ''                                      AS cbte            , "
	strSQL = strSQL & "                ctz.tipo                                AS tipo            , "
	strSQL = strSQL & "                convert(varchar(20), ctz.idcotizacion)  AS codigo          , "
	strSQL = strSQL & "                ctz.idproveedor                         AS idproveedor	  , "
	strSQL = strSQL & "                ctz.idpedido                            AS idpedidoAsociado, "
	strSQL = strSQL & "				   0									   AS importeVales    , "	
	if (pMonedaConsulta = MONEDA_PESOS) then
	    strSQL = strSQL & "                Sum(ctzdet.importepesosfacturado + ((ctzdet.importedolares - ctzdet.importedolaresfacturado ) * " & tc & " ))            AS importePIC      , "
	else
	    strSQL = strSQL & "                Sum(ctzdet.importedolares)          AS importePIC      , "
    end if	    
	strSQL = strSQL & "				   0									   AS importeFAC	  , "	
	strSQL = strSQL & "                ctz.fechaentrega                        AS fecha           , "
	strSQL = strSQL & "                ctz.momento                             AS momento         , "
	strSQL = strSQL & pCamposAreaDetalle
	strSQL = strSQL & "                ctz.estado                              AS habilitado "
	'MODIFICACION: A LA TABLA CTZCABECERA LE AGREGO UN FILTRO PARA QUE BUSQUE SOLO LOS PIC QUE NO TENGAN CONTRATOS
	strSQL = strSQL & "         FROM   (Select case when idcontrato = 0 then 'PIC' else 'CEC' end tipo, c.* from tblctzcabecera c where idobra = " & pIdObra & " and CDMONEDA='" & MONEDA_DOLAR & "') ctz"
	strSQL = strSQL & "                INNER JOIN tblctzdetalle ctzDet ON ctz.idcotizacion = ctzDet.idcotizacion"	
	strSQL = strSQL & " WHERE  ctz.ESTADO <> '" & CTZ_ANULADA & "' " & pFiltroPIC		
	strSQL = strSQL & "         GROUP BY ctz.idcotizacion,ctz.idobra ,ctz.idcotizacion, ctz.tipo, ctz.idproveedor, ctz.idpedido,ctz.fechaentrega,ctz.momento,ctz.estado" & pGroup

	'call logdebug(strSQL)
	armarSqlPIC_D = strSQL
End Function
'-----------------------------------------------------------------------------------------------
Function armarSqlFAC(pMoneda,pCamposAreaDetalle,pIdObra,pFiltro,pGroup)
	Dim strSQL,agregado
	Dim campoImporte
	
	campoImporte = "acd7.ImporteDolares"
	if (pMoneda=MONEDA_PESO) then campoImporte = "acd7.ImportePesos"
	
	strSQL =		  "SELECT   acd7.IDPIC  					 		AS id, "
	strSQL = strSQL & "         acd7.IDOBRA 					 		AS idobra, "
	strSQL = strSQL & "         Format (acds.succbt, '0000') + Format(acds.nrocbt, '00000000') AS cbte, "
	strSQL = strSQL & "         CASE WHEN acds.tipcbt = "& CBTE_PROVEEDORES_FAC &" then '" & PREFIX_FAC & "' when acds.tipcbt = "& CBTE_PROVEEDORES_NDB &" then '" & PREFIX_NDB & "' else '" & PREFIX_NCR & "' end	AS tipo, "
	strSQL = strSQL & "         convert(varchar(20), acd7.NroInt)       AS codigo, "
	strSQL = strSQL & "         acds.emicbt 					 		AS idproveedor, "
	strSQL = strSQL & "         0      							 		AS idpedidoAsociado, "
	strSQL = strSQL & "			0										AS importeVales, "	
	strSQL = strSQL & "			0										AS importePIC, "	
	strSQL = strSQL & "         SUM(" & campoImporte & ")*100 	 		AS importeFAC, "
	strSQL = strSQL & "         convert(varchar(10), acds.fecvto, 112)	AS fecha, "
	strSQL = strSQL & "         convert(varchar(10), acds.fecvto, 112)	AS momento, "
	strSQL = strSQL & pCamposAreaDetalle
	strSQL = strSQL & "         ''                               		AS habilitado "
	strSQL = strSQL & "FROM     (Select * from VWMEP001C where idobra =  " & pIdObra & ") acd7 "
	strSQL = strSQL & "         INNER JOIN VWCOMPROBANTES acds ON acd7.NroInt = acds.NroInt AND acds.anio=acd7.anio AND acds.mes=acd7.mes "
	strSQL = strSQL & "			INNER JOIN tblarticulos art on art.idarticulo=acd7.IDArticulo "
	strSQL = strSQL & "			INNER JOIN (Select * from tblartcategorias where tipocategoria <> '" & TIPO_CAT_IMPUESTOS & "') cat on art.idcategoria=cat.idcategoria "
	strSQL = strSQL & "WHERE "
	if (pFiltro <> "") then strSQL = strSQL & pFiltro & " and "
	strSQL = strSQL & "			acd7.IDArticulo NOT IN (" & ITEM_FONDO_REPARO_ARS & "," & ITEM_FONDO_REPARO_USD & ")"
	strSQL = strSQL & " GROUP BY acd7.NroInt,acd7.IDPIC, "
	strSQL = strSQL & "         acd7.IDOBRA, acds.emicbt, "
	strSQL = strSQL & "         acds.fecvto,acds.succbt, acds.nrocbt, acds.tipcbt " & pGroup

	'call logdebug(strSQL)
	armarSqlFAC = strSQL
End Function
'-----------------------------------------------------------------------------------------------
Function armarSqlVales(pMoneda,pCamposAreaDetalle,pIdObra,pFiltro,pGroup)
	Dim strSQL,agregado
	Dim campoImporte
	
	campoImporte = "det.VLUDOLARES"
	if (pMoneda=MONEDA_PESO) then campoImporte = "det.VLUPESOS"
	
	strSQL =          "SELECT 	det.idvale 					AS id, "
	strSQL = strSQL & "			cab.idobra					AS idobra, "
	strSQL = strSQL & "			''							AS cbte, "
	strSQL = strSQL & "			cab.cdvale					AS tipo, "
	strSQL = strSQL & "			cab.nrvale					AS codigo, "
	strSQL = strSQL & "			0							AS idProveedor, "
	strSQL = strSQL & "			0							AS idpedidoAsociado, "
	strSQL = strSQL & "			SUM("& campoImporte &"*existencia) AS importeVales, "
	strSQL = strSQL & "			0							AS importePIC, "
	strSQL = strSQL & "			0							AS importeFAC, "	
	strSQL = strSQL & "			cab.momento					AS fecha, "
	strSQL = strSQL & "			cab.momento					AS momento, "
	strSQL = strSQL & pCamposAreaDetalle
	strSQL = strSQL & "			''							AS habilitado "
	strSQL = strSQL & "FROM 	tblvalescabecera cab "
	strSQL = strSQL & "	   		INNER JOIN tblvalesdetalle det "
	strSQL = strSQL & "				ON cab.idvale = det.idvale "
	strSQL = strSQL & "WHERE  cab.idobra = " & pIdObra
	strSQL = strSQl & " AND   cab.estado = "&ESTADO_ACTIVO&" "
	strSQL = strSQL & pFiltro
	strSQL = strSQL & " GROUP  BY det.idvale,cab.idobra,cab.nrvale,cab.momento,cab.cdvale  " & pGroup

	'call logdebug(strSQL)
	armarSqlVales = strSQL
End Function
'-----------------------------------------------------------------------------------------------
' Function: getTotales
' Descripcion:
' Esta funcion se encarga de calcular la suma total de los CTZ/PIC, los AFEs y el total general
' de los movimientos filtrados
' Los valores son devueltos por referencia.
'-----------------------------------------------------------------------------------------------
Sub getTotales(pCdMoneda, pIdObra, pIdArea, pIdDetalle, byref pSumaPIC, byref pSumaFAC, byref pSumaVALES) 	
	
	Dim rs
	
	Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", pIdObra & "||" & pIdArea & "||" & pIdDetalle)		
	if (not rs.eof) then
	    if (pCdMoneda = MONEDA_DOLAR) then
	        pSumaPIC = Cdbl(rs("IMPORTEPICDOLARES"))
            pSumaVALES = Cdbl(rs("IMPORTEVALESDOLARES"))	        
        else
            pSumaPIC = Cdbl(rs("IMPORTEPICPESOS"))
            pSumaVALES = Cdbl(rs("IMPORTEVALESPESOS"))            
        end if        
    else
        pSumaPIC = 0
        pSumaVALES = 0
	end if	
	'pSumaPIC = calcularGastosObra(pCdMoneda, pIdObra, pIdArea, pIdDetalle, false)
	pSumaFAC = calcularGastosFacturados(pIdObra,pIdArea,pIdDetalle,"","",pCdMoneda)
	'pSumaVALES = obtenerTotalValesObra(pIdObra,pIdArea,pIdDetalle,"",pCdMoneda)

End sub
'-----------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'-------------------------------------INICIO PAGINA------------------------------------------------
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Call comprasControlAccesoCM(RES_OBR)

Dim registros
Dim sumaPIC, sumaPICTotal, sumaFAC,sumaVALE
Dim idArea,idDetalle,myaux
Dim presupuesto
Dim myValue,myClass,mySelect,aux
Dim linkFac, folder,totalRows
Dim accion, importeFacturado
dim totalAFEs, flagAFEs, rutaProcedencia
presupuesto = 0
accion = GF_PARAMETROS7("accion","",6)
idObra = GF_PARAMETROS7("idObra","",6)
cdMoneda = GF_PARAMETROS7("cdMoneda","",6)
if cdMoneda = "" then cdMoneda = MONEDA_DOLAR
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
myaux = GF_PARAMETROS7("AreaDetalle","",6)
flagAFEs = false

'compruebo los filtros area y detalle. 
if (myaux <> "" and myaux <> ",") then
	myaux  = split(myaux,",")
	idArea = CInt(myaux(0))
	if (idArea <> TODOS_LOS_MOVIMIENTOS) then 
		idDetalle = CInt(myaux(1))
	end if
end if

Call GP_ConfigurarMomentos()
Set registros = obtenerListaEntradas(idObra, cdMoneda, paginaActual, mostrar, idArea, idDetalle)

Call setupPaginacion(registros, paginaActual, mostrar)
Set rsObra = obtenerListaObras(idObra, "", "", "", "")
if (rsObra.eof) then
	response.redirect "comprasAccesoDenegado.asp"
end if

lineasTotales = registros.RecordCount

if (isInversion(idObra)) then 
	presupuesto = totalizarAFESObra(cdMoneda, idObra, idarea, iddetalle, false)
	if presupuesto <> 0 then flagAFEs = true
end if
'Response.Write "(" & presupuesto & ")"
if presupuesto = 0 then
	presupuesto = calcularCostoEstimadoObra(cdMoneda,idObra,idarea,iddetalle)
end if

saldo = presupuesto

%>

<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" type="text/css" href="css/main.css">
    <link rel="stylesheet" href="css/Header.css">
    <style>
	    input[type=text] { text-align:right;}
  	    label { padding: 0; marign: 0; display: block; }
  	    textarea { width: 100%; border: 0px; padding: 0px; }
  	    .the-fix { }
	    .celda {
		    border-radius:8px 8px 8px 8px;
	    }
	    table.exp {
    	    border-collapse: separate;
    	    border-spacing: 4px 0px;
	    }
    </style>
	<script type="text/javascript" src="scripts/paginar.js"></script>
	<script type="text/javascript">
		function bodyOnLoad()
		{
			<%	if (not registros.eof) then		%>								
					var pgn = new Paginacion("paginas");
					pgn.paginar(<% =paginaActual %>, <% =registros.recordcount %>, <% =mostrar %>, 50, "comprasTableroObraDetalle.asp?idObra=<% =idObra %>&areaDetalle=<%= idArea & "," & idDetalle%>&cdMoneda=<%=cdMoneda%>");
			<%	end if %>
			parent.iFrameOnLoad();
		}

		
		function abrirPIC(id) {
			window.open("comprasPICPrint.asp?idCotizacionElegida=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
		}
		
		function abrirCEC(id) {
			abrirPIC(id);
		}
		
		function abrirCTC(id) {
			window.open("comprasCTC.asp?idContrato=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
		}
		
		function abrirFAC(pLink) {
			window.open(pLink, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);
		}
		
		function abrirNCR(pLink) {
			abrirFAC(pLink);
		}
		
		function abrirAFEPrint(id) {
			window.open("comprasAFEPrint.asp?idAFE=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700",false);		
		}
		
		function abrirObra(id) {
			window.open("comprasPropObra.asp?idObra=" + id, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=300,width=400",false);
		}
		
		function abrirVALE(id){
			window.open("almacenValePedidoPrint.asp?idVale="+id,"_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=800",false);
		}
	</script>
</head>
<body onLoad="bodyOnLoad()">
	<form id="myForm" action="comprasTableroObraDetalle.asp">
	<input type="hidden" id="ruta" name="ruta" value="<%=rutaProcedencia%>">
	<input type="hidden" id="idObra" name="idObra" value="<%=idObra%>">
	<input type="hidden" id="cdMoneda" name="cdMoneda" value="<%=cdMoneda%>">
	<div class="col66"></div>
    <br>
    <table width="320" align="center" border="1" bordercolor="#999999">
        <thead>
		    <tr>
    	    <td align="left">DETALLES DE MOVIMIENTOS PARA:</td>
	        </tr>
	    </thead>
        <tr>
            <td align="center" height="40">
                <select name="AreaDetalle" id="AreaDetalle" size="1" onChange="javascript:submit()">
                <%  if (idArea = "") then 
                        idArea = TODOS_LOS_MOVIMIENTOS
                        idDetalle = TODOS_LOS_MOVIMIENTOS
                    end if                        
                %>																
			        <option value =<%=TODOS_LOS_MOVIMIENTOS%>><%=GF_TRADUCIR("Completo")%></option>
			         <% Set rs = leerBudget(idObra)						
					    while not rs.eof 
				            myValue = rs("IDAREA") & "," & rs("IDDETALLE")
						    if (rs("IDDETALLE")=0) then 
						        myClass =  "class='titulo' " 
						    else 
						        myClass = ""
						    end if
						    if (cint(idArea)=rs("IDAREA")) and (cint(iddetalle) = rs("IDDETALLE")) then 
						        mySelect = "selected='selected'"
						    else 
						        mySelect = ""
						    end if %>
					        <option value="<%=myValue %>" <%=myClass%> <%=mySelect%> >
						    <%  if (rs("IDDETALLE")<>0) then
						            response.write "&nbsp;&nbsp;&nbsp;&nbsp;" & rs("IDDETALLE") & " - " &  rs("DSBUDGET")
						        else
								    response.write rs("IDAREA") & " - " & rs("DSBUDGET")
							    end if%>
					        </option>
						    <%rs.movenext
						    wend%>
			    </select>
            </td>
	    </tr>
    </table>
    
    <table class="datagrid" width="95%" align="center">
        <thead>
            <tr>
                <th colspan="3" align="center">FORMULARIO</th>
                <th rowspan="2" align="center">FECHA</th>
                <th rowspan="2" align="center">PRESUPUESTO</th>
                <th colspan="3" align="center">GASTOS</th>
                <th rowspan="2" align="center">SALDO</th>
            </tr>
            <tr>
                <td align="center">MINUTA</td>
                <td align="center">COMPROBANTE</td>
                <td align="center">PROVEEDOR</td>
                <td align="center">VALES</td>
                <td align="center">PEDIDOS</td>
                <td align="center">FACTURADOS</td>
            </tr>
        </thead>
        <tbody>
            <tr>
				<td colspan="3"  style="font-weight:bold; color:#000;text-align: left" height="20px">
					<%  if flagAFEs then 
						    Response.Write GF_TRADUCIR("AFEs")
					    else
						    Response.Write GF_TRADUCIR("OBR") & ":&nbsp;<a href='javascript:abrirObra(" & rsObra("idObra") & ");'>" & rsObra("cdObra") & "</a>"
					    end if %>
				</td>
				<td style="text-align: center"><% =GF_FN2DTE(rsObra("FECHAINICIO")) %></td>
				<td style="text-align: right"><%=GF_EDIT_DECIMALS(presupuesto,2)%></td>
				<td></td>
				<td></td>
				<td style="text-align: center">&nbsp;</td>
				<td style="text-align: right"><%=GF_EDIT_DECIMALS(presupuesto,2)%></td>				
			</tr>	
         <%	reg=0
	        if (not registros.eof) then
			    saldo = presupuesto 		
			    while ((not registros.eof) and (reg < mostrar))
    				reg=reg+1
				    importeFacturado = CDbl(registros("ImporteFAC")) %>
                    <tr <% if (registros("habilitado")= AFE_ANULADO) then %> class="reg_header_rejected" <%end if%> >
						<td style="text-align: left">
							<% 
							if (CDbl(registros("ImporteFAC")) > 0)then 							
								if cstr(registros("tipo")) = PREFIX_NCR then importeFacturado = cdbl(registros("ImporteFAC")) * -1								
								Response.write registros("tipo") & ": " & registros("CODIGO")
							else 									
									if (CDbl(registros("ImporteVales")) > 0)then 
										Response.write GF_TRADUCIR(registros("tipo")) & ":<a href='javascript:abrirVALE(" & registros("id") & ");'>&nbsp;" & registros("CODIGO") & "</a>"
									else
										Response.write GF_TRADUCIR(registros("tipo")) & ":<a href='javascript:abrir" & registros("tipo") & "(" & registros("id") & ");'>&nbsp;" & registros("CODIGO") & "</a>"
									end if
							end if %>
                        </td>	
						<td style="text-align: center">
							<%
								if (Trim(CStr(registros("cbte"))) <> "") then
									if (CDbl(registros("cbte"))<> 0) then
										response.write GF_EDIT_CBTE(registros("cbte"))
									else
										response.write "&nbsp;"
									end if
								else
									response.write "&nbsp;"
								end if
							%>                </td>
						<td>&nbsp;&nbsp;<% =registros("idProveedor") %> - <% =getDescripcionProveedor(CLng(registros("idProveedor"))) %></td>
						<td style="text-align: center"><% =GF_FN2DTE(left(registros("FECHA"),8)) %></td>
						<td style="text-align: center">&nbsp;</td>
						<td style="text-align: right"><%=GF_EDIT_DECIMALS(registros("ImporteVales"),2)%></td>
						<td style="text-align: right"><%=GF_EDIT_DECIMALS(registros("ImportePIC"),2)%></td>
						<td style="text-align: right"><%=GF_EDIT_DECIMALS(importeFacturado,2)%></td>
						<%
							saldo = CDbl(saldo) - CDbl(registros("ImporteVales")) 
							saldo = CDbl(saldo) - CDbl(registros("ImportePIC"))							
						%>
						<td width="35" style="text-align: right"><%=GF_EDIT_DECIMALS(saldo,2)%></td>
					</tr>									
		    <%  sumaVALE = sumaVALE + CDbl(registros("ImporteVales"))
				SumaPIC = SumaPIC   + CDbl(registros("ImportePIC"))
				SumaFAC = SumaFAC   + importeFacturado
				registros.MoveNext()				
			wend %>
	    <%end if%>
        </tbody>
        <thead>
            <tr>
        	    <td colspan="2" align="center">TOTAL PÁGINA</td>
                <td align="center">&nbsp;</td>
                <td align="center">&nbsp;</td>
                <td align="right"><%=GF_EDIT_DECIMALS(presupuesto, 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(sumaVALE),2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(SumaPIC), 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(SumaFAC), 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS((CDbl(presupuesto) - (CDbl(SumaPIC)+CDbl(SumaVALE)) ), 2)%></td>
            </tr>
        </thead>
        <tbody>
            <tr colspan="9" align="center">
                <td bgcolor="#FFF" height="10px">&nbsp;</td>
            </tr>
        </tbody>
        <% 'Se calculan los totales para la tabla.
		    Call getTotales(cdMoneda, idObra, idArea, idDetalle, sumaPIC, sumaFAC, sumaVALE) %>
         <thead>
            <tr>
        	    <td colspan="2" align="center">TOTAL GENERAL</td>
                <td align="center">&nbsp;</td>
                <td align="center">&nbsp;</td>
                <td align="right"><%=GF_EDIT_DECIMALS(presupuesto, 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(sumaVALE),2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(SumaPIC), 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS(Cdbl(SumaFAC), 2)%></td>
                <td align="right"><%=GF_EDIT_DECIMALS((CDbl(presupuesto) - (CDbl(SumaPIC)+CDbl(SumaVALE)) ), 2)%></td>
            </tr>
        </thead>
        <tfoot>
		    <tr><td colspan="10" height="40px"><div id="paginas"></div></td></tr>
	    </tfoot>
    </table>
</form>
</body>
</html>