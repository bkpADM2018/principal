<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosObras.asp"-->
<!--#include file="../Includes/procedimientosCTC.asp"-->
<!--#include file="../Includes/procedimientosCTZ.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<%
'Esta function obtiene la obra de mantenimiento de un determinado periodo
Function obtenerObraMantenimientoPorPeriodo(pFechaInicio, pFechaFin, pPto)
    Dim strSQL,rs,strObra
    strObra = ""
    strSQL = "SELECT IDOBRA "&_
             "FROM TOEPFERDB.TBLDATOSOBRAS "&_
             "WHERE TIPOGASTO in ('"& OBRA_TIPO_MANT_TRIM &"', '" & OBRA_TIPO_MANT_ANUAL & "') "&_
             "  AND FECHAINICIO >= '"& pFechaInicio &"'"&_
             "  AND (CASE WHEN FECHAAJUSTADA = 0 THEN FECHAFIN ELSE FECHAAJUSTADA END) <= '"& pFechaFin &"'"&_
             "  AND IDDIVISION = (SELECT IDDIVISION FROM TOEPFERDB.TBLDIVISIONES WHERE CDDIVISIONABR = '"& getLetraPuerto(pPto) &"')"&_
             " ORDER BY IDOBRA DESC "
    Call executeQuery(rs, "OPEN", strSQL)
    if (not rs.Eof) then strObra = rs("IDOBRA")
    obtenerObraMantenimientoPorPeriodo = strObra
End Function
'--------------------------------------------------------------------------------------------
Function getTotalAlmacen(pObra, pTipoCambio)
    Dim strSQL
    strSQL = "SELECT case when Sum(T1.gasto) is null then 0 else Sum(T1.gasto) end as almacen "&_
             "FROM ( "&_
             "      SELECT cab.idobra, "&_
             "              SUM(det.existencia*(det.vlupesos/"&pTipoCambio&")) gasto "&_
             "      FROM toepferdb.tblvalescabecera cab  "&_
             "      INNER JOIN toepferdb.tblvalesdetalle det  "&_
	         "          ON cab.idvale = det.idvale  "&_
             "      WHERE cab.idobra in ("& pObra &") "&_
	         "          AND cab.estado = "& ESTADO_ACTIVO &_
	         "          AND cab.fecha <= "& left(session("MmtoDato"),8) &_
             "      GROUP BY cab.idobra ) T1"
    Call executeQuery(rs, "OPEN", strSQL)
    getTotalAlmacen = 0
    if (not rs.Eof) then getTotalAlmacen = Cdbl(rs("almacen"))
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim g_strPuerto,anio,idObra,fechaInicio,fechaFinal,gTipoCambio
   
g_strPuerto = GF_Parametros7("pto","",6)
anio = GF_PARAMETROS7("anio", 0, 6)
if (Cdbl(anio) = 0) then anio = Year(Now())

fechaInicio = anio & "0101"
fechaFinal  = anio & "1231"

strValue = ""
idObra = obtenerObraMantenimientoPorPeriodo(fechaInicio,fechaFinal,g_strPuerto)
strSQL = "Select * from TOEPFERDB.TBLBUDGETObras where idobra = " & idObra
Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
gTipoCambio = rs("TIPOCAMBIO")
if (idObra <> "") then
    g_Presupuesto   = calcularCostoEstimadoObra(MONEDA_DOLAR, idObra, 0, 0)
    g_Comprometido  = calcularGastosObra(MONEDA_DOLAR, idObra, 0, 0, False)
    g_Pagado        = calcularGastosFacturados(idObra,0, 0, "", "", MONEDA_DOLAR)
    g_Almacen       = getTotalAlmacen(idObra,gTipoCambio)
    g_Saldo         = Cdbl(g_Presupuesto) - Cdbl(g_Comprometido) - Cdbl(g_Almacen)
    strValue = idObra & STRING_DELIMITER & GF_EDIT_DECIMALS(g_Presupuesto,2) & STRING_DELIMITER & GF_EDIT_DECIMALS(g_Comprometido,2) & STRING_DELIMITER & GF_EDIT_DECIMALS(g_Pagado,2) & STRING_DELIMITER & GF_EDIT_DECIMALS(g_Almacen,2) & STRING_DELIMITER & GF_EDIT_DECIMALS(g_Saldo,2)
end if
Response.Write strValue
%>