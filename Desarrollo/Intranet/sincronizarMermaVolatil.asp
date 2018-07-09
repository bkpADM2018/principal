<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<%
'---------------------------------------------------------------------------------------------------------------------------------------
Function obtenerCamionesPorFecha(pFechaDesde, pFechaHasta, pPto)
    Dim strSQL
    strSQL = " SELECT A.IDCAMION, A.CDPRODUCTO, B.NUCARTAPORTE CARTAPORTE, ((YEAR(A.DTCONTABLE)*10000) + (MONTH(A.DTCONTABLE)*100) + DAY(A.DTCONTABLE)) as DTCONTABLE, CDESTADO"&_
             "      FROM HCAMIONES A "&_
             "      INNER JOIN HCAMIONESDESCARGA B ON A.IDCAMION = B.IDCAMION AND A.DTCONTABLE = B.DTCONTABLE "&_      
             "      LEFT JOIN MERMAVOLATIL C on B.DTCONTABLE=C.DTCONTABLE and B.IDCAMION=C.IDTRANSPORTE and TIPOTRANSPORTE=" & TIPO_TRANSPORTE_CAMION &_
             "WHERE CDESTADO in (6, 8) and A.DTCONTABLE >= '"& GF_FN2DTCONTABLE(pFechaDesde) &"' and A.DTCONTABLE <= '"& GF_FN2DTCONTABLE(pFechaHasta) &"' and C.IDTRANSPORTE is Null"
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    Set obtenerCamionesPorFecha = rs
End function
'---------------------------------------------------------------------------------------------------------------------------------------
Function obtenerVagonesPorFechaH(pFechaDesde, pFechaHasta, pPto)
    Dim strSQL
    strSQL = " SELECT CDVAGON, CDPRODUCTO, (NUCARTAPORTESERIE + LEFT(B.NUCARTAPORTE, 8)) CARTAPORTE, B.NUCARTAPORTE, ((YEAR(DTCONTABLEVAGON)*10000) + (MONTH(DTCONTABLEVAGON)*100) + DAY(DTCONTABLEVAGON)) as DTCONTABLE, CDESTADO "&_ 
             "      FROM HVAGONES B"&_  
             "      LEFT JOIN MERMAVOLATIL C on B.DTCONTABLEVAGON=C.DTCONTABLE and B.CDVAGON=C.IDTRANSPORTE and TIPOTRANSPORTE=" & TIPO_TRANSPORTE_VAGON &_
             " WHERE CDESTADO in (6, 8) and DTCONTABLEVAGON >= '"& GF_FN2DTCONTABLE(pFechaDesde) &"' and DTCONTABLEVAGON <= '"& GF_FN2DTCONTABLE(pFechaHasta) &"' and C.IDTRANSPORTE is Null  "             
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    Set obtenerVagonesPorFechaH = rs
End function
'---------------------------------------------------------------------------------------------------------------------------------------
Function obtenerVagonesPorFechaD(pFechaDesde, pFechaHasta, pPto)
    Dim strSQL
    strSQL = " SELECT CDVAGON, CDPRODUCTO, (NUCARTAPORTESERIE + LEFT(B.NUCARTAPORTE, 8)) CARTAPORTE, B.NUCARTAPORTE, ((YEAR(DTCONTABLEVAGON)*10000) + (MONTH(DTCONTABLEVAGON)*100) + DAY(DTCONTABLEVAGON)) as DTCONTABLE, CDESTADO "&_ 
             "      FROM VAGONES B"&_  
             "      LEFT JOIN MERMAVOLATIL C on B.DTCONTABLEVAGON=C.DTCONTABLE and B.CDVAGON=C.IDTRANSPORTE and TIPOTRANSPORTE=" & TIPO_TRANSPORTE_VAGON &_
             "WHERE CDESTADO in (6, 8) and DTCONTABLEVAGON >= '"& GF_FN2DTCONTABLE(pFechaDesde) &"' and DTCONTABLEVAGON <= '"& GF_FN2DTCONTABLE(pFechaHasta) &"' and C.IDTRANSPORTE is Null  "             	
    Call executeQueryDb(pPto, rs, "OPEN", strSQL)
	Set obtenerVagonesPorFechaD = rs
End function
'---------------------------------------------------------------------------------------------------------------------------------------
Function cargarMermaVolatil(pIdTransporte, pCdProducto, pCtaPte, pNuCtaPrte, pDtcontable, pTipoTransporte, pPto)
    Dim rs,kiloMerma

    Call logMig.info("Analizando si tiene merma volatil")
    if (pTipoTransporte = TIPO_TRANSPORTE_CAMION) then
        Call logMig.info("-->Ejecutando HCAMIONESDESCARGA_GET_MERMAVOLATIL_CALCULAR --> Fecha: "& GF_FN2DTCONTABLE(pDtcontable) &" | Producto: "& pCdProducto &" | IdCamion: "&pIdTransporte&" | Carta Porte: "& pCtaPte )
        Call executeSP_Puertos(rs, pPto, "HCAMIONESDESCARGA_GET_MERMAVOLATIL_CALCULAR", GF_FN2DTCONTABLE(pDtcontable) &"||"& GF_FN2DTCONTABLE(pDtcontable) &"||"& pCdProducto &"||"& pIdTransporte &"||"& pCtaPte)
    else
        Call logMig.info("-->Ejecutando HVAGONES_GET_MERMAVOLATIL_CALCULAR --> Fecha: "& GF_FN2DTCONTABLE(pDtcontable) &" | Producto: "& pCdProducto &" | CdVagon: "&pIdTransporte&" | Carta Porte: "& pNuCtaPrte )
        Call executeSP_Puertos(rs, pPto, "HVAGONES_GET_MERMAVOLATIL_CALCULAR", GF_FN2DTCONTABLE(pDtcontable) &"||"& GF_FN2DTCONTABLE(pDtcontable) &"||"& pCdProducto &"||"& pIdTransporte &"||"& pNuCtaPrte)
    end if
    
    if (not rs.Eof) then
        'Si el camion/vagon tiene un porcentaje de ratio asignado lo agrego a la tabla Merma volatil
		kiloMerma = Round((Cdbl(rs("RATIO")) * Cdbl(rs("PESO")))/100)		
        strSQL = "INSERT INTO MERMAVOLATIL VALUES('"& GF_FN2DTCONTABLE(pDtcontable) &"','"& pIdTransporte &"', " & pTipoTransporte & ",'"& Trim(pCtaPte) &"',"& kiloMerma &")"
        Call GF_BD_Puertos(pPto, rs, "EXEC", strSQL)    
        Call logMig.info("Agrega merma volatil: "& strSQL)        
    end if
End function
'---------------------------------------------------------------------------------------------------------------------------------------
Function procesarMermaVolatilPuerto(pPto, pFechaDesde, pFechaHasta)
    Dim rsCam,rsVag
    procesarMermaVolatilPuerto = "OK"
    Call logMig.info("------------------------------------------------- INICIA "& pPto &" -------------------------------------------------")   
	'Obtengo los camiones descargados de una determinada fecha
	Set rsCam = obtenerCamionesPorFecha(pFechaDesde, pFechaHasta, pPto)
	Call logMig.info("Cantidad de camiones descargados: "& rsCam.RecordCount)
	while (not rsCam.Eof )
		'Por cada camion traigo el procentaje de ratio de la merma volatil
		Call cargarMermaVolatil(rsCam("IDCAMION"), rsCam("CDPRODUCTO"), rsCam("CARTAPORTE"), "", rsCam("DTCONTABLE"), TIPO_TRANSPORTE_CAMION, pPto)
		rsCam.MoveNext()
	wend

	Set rsVag = obtenerVagonesPorFechaD(pFechaDesde, pFechaHasta, pPto)
	Call logMig.info("Cantidad de vagones descargados (D): "& rsVag.RecordCount)
	while (not rsVag.Eof )
		Call cargarMermaVolatil(rsVag("CDVAGON"), rsVag("CDPRODUCTO"), rsVag("CARTAPORTE"), rsVag("NUCARTAPORTE"), rsVag("DTCONTABLE"), TIPO_TRANSPORTE_VAGON, pPto)
		rsVag.MoveNext()
	wend
	
	Set rsVag = obtenerVagonesPorFechaH(pFechaDesde, pFechaHasta, pPto)
	Call logMig.info("Cantidad de vagones descargados (H): "& rsVag.RecordCount)
	while (not rsVag.Eof )
		Call cargarMermaVolatil(rsVag("CDVAGON"), rsVag("CDPRODUCTO"), rsVag("CARTAPORTE"), rsVag("NUCARTAPORTE"), rsVag("DTCONTABLE"), TIPO_TRANSPORTE_VAGON, pPto)
		rsVag.MoveNext()
	wend
	
	Call logMig.info("------------------------------------------------- FINALIZA "& pPto &" -------------------------------------------------")    
End function
'************************************************************************************************************************************************************
'*********************************************************          COMIENZO DE LA PAGINA           *********************************************************
'************************************************************************************************************************************************************
Dim fecha,logMig, fechaHasta, myRet

Call GP_CONFIGURARMOMENTOS()

fechaDesde = GF_PARAMETROS7("fd", 0, 6)
if (fechaDesde = 0 ) then fechaDesde = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")
fechaHasta = GF_PARAMETROS7("fh", 0, 6)
if (fechaHasta = 0 ) then fechaHasta = Left(session("MmtoDato"), 8)


Set logMig = new classLog
Call startLog(HND_FILE+HND_VIEW,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "SINCRONIZAR_MERMA_VOLATIL_"& left(session("MmtoDato"),8)
Call logMig.info("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* INICIA SINCRONIZACION -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
Call logMig.info("Periodo: "& GF_FN2DTE(fechaDesde) & " - " & GF_FN2DTE(fechaHasta))

myRet = TERMINAL_ARROYO & ":" & procesarMermaVolatilPuerto(TERMINAL_ARROYO, fechaDesde, fechaHasta)

myRet = myRet & "|" & TERMINAL_PIEDRABUENA & ":" & procesarMermaVolatilPuerto(TERMINAL_PIEDRABUENA, fechaDesde, fechaHasta)

myRet = myRet & "|" & TERMINAL_TRANSITO & ":" & procesarMermaVolatilPuerto(TERMINAL_TRANSITO, fechaDesde, fechaHasta)

Call GP_ENVIAR_MAIL("SINCRONIZAR MERMA VOLATIL - OK",myRet,"scalisij@toepfer.com","scalisij@toepfer.com")

Call logMig.info("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* FINALIZA SINCRONIZACION -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")    
%>
<html>
    <body>
	    <form method="post" action="sincronizarMermaVolatil.asp" name="frmSincro" id="frmSincro" target="ifrm">
    	    <input type="hidden" name="fecha" id="fecha" value="<% =fecha %>" />
	    </form>
    </body>
</html>