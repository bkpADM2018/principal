<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->

<%
'--------------------------------------------------------------------------------------------
Function getCuposAsignados(pPto, pFechaDesde, pCuitCliente)
    Dim rs, ret, strSQL
    'Cupos asignados para Hoy
	strSQL="Select count(*) CANT from CODIGOSCUPO where FECHACUPO=" & pFechaDesde & " and ESTADO > 0"    
	if (CDbl(pCuitCliente) > 0) then strSQL= strSQL & " and CUITCLIENTE='" & pCuitCliente & "'"    
    Call executeQueryDb(pPto, rs, "OPEN", strSQL)
    ret = 0
    if (not rs.eof) then ret = CLng(rs("CANT"))
    getCuposAsignados = ret
End Function
'--------------------------------------------------------------------------------------------
Function getIngresadosHoy(pPto, pFechaDesde, pCuitCliente)
    Dim rs, ret, strSQL
    'Ingresados con cupos asignados para Hoy
    strSQL="Select count(*) CANT from CAMIONES C inner join CAMIONESDESCARGA CD on C.IDCAMION=CD.IDCAMION where CDESTADO<>" & CAMIONES_ESTADO_BAJA
	if (CDbl(pCuitCliente) > 0) then strSQL= strSQL & " and CDCLIENTE in (Select CDCLIENTE from CLIENTES where NUCUIT='" & pCuitCliente & "')"    
    Call executeQueryDb(pPto, rs, "OPEN", strSQL)    
    ret = 0
    if (not rs.eof) then ret = CLng(rs("CANT"))
    getIngresadosHoy = ret
End Function
'--------------------------------------------------------------------------------------------
Function getTotalCupos(pFechaDesde,pFechaHasta,pCdProducto,pPuerto, pCuitCliente)
    Dim rsCup
    getTotalCupos = getCuposAsignados(pPuerto, pFechaDesde, pCuitCliente) & STRING_DELIMITER & getIngresadosHoy(pPuerto, pFechaDesde, pCuitCliente)
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim g_strPuerto,fechaDesde,fechaHasta,cdProducto,strValue
Dim cuitCliente
   
g_strPuerto = GF_Parametros7("pto","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
fechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)

if (fechaDesde = "") then fechaDesde = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2) 
if (fechaHasta = "") then fechaHasta = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)

cuitCliente = 0
if (not IsToepfer(session("KCOrganizacion"))) then cuitCliente = session("CuitOrganizacion")

strValue = getTotalCupos(fechaDesde,fechaHasta,cdProducto,g_strPuerto, cuitCliente)
Response.Write strValue

%>

