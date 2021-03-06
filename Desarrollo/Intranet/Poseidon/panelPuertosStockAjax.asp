<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<%
'---------------------------------------------------------------------------------------------------------------------------------
Function armarSQLStock(pFechaDesde, pFechaHasta, pPto, pCdCliente, pCuitCliente)
    Dim strSQL, fechaDesde, fechaHasta,rs
    strSQL = "SELECT TOP 5 A.KILOS, A.CDPRODUCTO, B.DSPRODUCTO "&_
             "FROM ( SELECT cc.cdproducto, sum( cc.vlsaldoinicial + cc.vlcredito - cc.vldebito ) as Kilos "&_
	         "     FROM excuentcorrientes cc "&_
	         "     WHERE cc.dtcontable >= '"& pFechaDesde &"' and cc.dtcontable <= '"& pFechaHasta &"' "&_
             "          AND ( cc.vlsaldoinicial + cc.vlcredito - cc.vldebito ) <> 0 " 
	if (not isToepfer(pCdCliente)) then		 
		strSQL = strSQL & " and cdcliente in (Select CDCLIENTE from clientes where NUCUIT = '" & pCuitCliente & "') "
	end if
	strSQL = strSQL & " GROUP BY cc.CDPRODUCTO ) A "&_
             " LEFT JOIN PRODUCTOS B ON B.CDPRODUCTO = A.CDPRODUCTO "&_
             "ORDER BY A.KILOS DESC "             
	Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	Set armarSQLStock = rs
End Function
'--------------------------------------------------------------------------------------------
Function convertTnToBushel(pTonelada, pCdProducto)
    Dim strSQL,rsBus,auxTn
    strSQL = "SELECT FACTOR "&_
             "FROM TOEPFERDB.TBLUNIDADESCONVERSION "&_
             "WHERE IDUNIDADDEST = (SELECT IDUNIDAD "&_
             "                      FROM TOEPFERDB.TBLUNIDADES "&_
             "                      WHERE CDUNIDAD LIKE '%"& STRING_DELIMITER & pCdProducto & STRING_DELIMITER &"%')"
    Call executeQuery(rsBus, "OPEN", strSQL)
    convertTnToBushel = 0
    if (not rsBus.Eof) then convertTnToBushel = GF_EDIT_DECIMALS(Round(Cdbl(pTonelada)*Cdbl(rsBus("FACTOR")),0),0)
End Function
'--------------------------------------------------------------------------------------------
Function getTotalStock(pFechaDesde,pFechaHasta,pUnidad,pPuerto, pCdCliente, pCuitCliente)
    Dim rsSto, strStock, auxPeso, dsProducto
    strStock = ""
    Set rsSto = armarSQLStock(pFechaDesde,pFechaHasta,pPuerto, pCdCliente, pCuitCliente)
    if not rsSto.Eof then
        while (not rsSto.Eof)
            Select Case Cdbl(pUnidad)
                Case TIPO_PESO_KILO
                    auxPeso = GF_EDIT_DECIMALS(Cdbl(rsSto("KILOS")),0)                    
                Case TIPO_PESO_TONELADA
                    auxPeso = GF_EDIT_DECIMALS(Round(Cdbl(rsSto("KILOS"))/1000),0) & " Tn"
                Case TIPO_PESO_BUSHEL
                    auxPeso = convertTnToBushel(Cdbl(rsSto("KILOS"))/1000,rsSto("CDPRODUCTO"))
            End Select
            'Por medio del Codigo de Producto del Puerto obtengo la descripci�n en Ingl�s situada en la tabla de Buenos Aires
            dsProducto = rsSto("DSPRODUCTO")
            strStock = strStock & dsProducto & ":" & auxPeso & STRING_DELIMITER
            rsSto.MoveNext()
        wend
        strStock = left(strStock,len(strStock)-1)
    End If
    getTotalStock = strStock
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim g_strPuerto,fechaDesde,fechaHasta,unidad,strValue

g_strPuerto = GF_Parametros7("pto","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
fechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
unidad = GF_PARAMETROS7("unidad", 0, 6)

if (fechaDesde = "") then fechaDesde = Year(Now()) &"-"& GF_nDigits(Month(Now()),2) &"-"& GF_nDigits(Day(Now()),2) 
if (fechaHasta = "") then fechaHasta = Year(Now()) &"-"& GF_nDigits(Month(Now()),2) &"-"& GF_nDigits(Day(Now()),2)

strValue = getTotalStock(fechaDesde,fechaHasta,unidad,g_strPuerto, session("KCOrganizacion"), session("CuitOrganizacion"))
Response.Write strValue

%>