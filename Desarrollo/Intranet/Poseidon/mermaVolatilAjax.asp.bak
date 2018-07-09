<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosTitulos.asp"-->
<!--#include file="../includes/procedimientosSQL.asp"-->
<!--#include file="../includes/procedimientosLog.asp"-->
<%
Function getSelectProducto(pPto) 
    Dim rsProductos %>
    <select id="cdProducto" name="cdProducto" style="width:95%;" >
        <option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
    	<%  strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
	        call GF_BD_Puertos (pPto, rsProductos, "OPEN",strSQL)
		    while not rsProductos.eof %>
				<option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
		    <%	rsProductos.movenext
			wend %>
       </select>
<%
End Function
'----------------------------------------------------------------------------------------------------------------------
Function eliminarMermaVolatil(pPto,pFecha,pCdProducto,pCdCliente,pCdSilo)
    Dim strSQL
    strSQL = "DELETE FROM DBO.TBLREGLASMERMAVOLATIL "&_
             "WHERE DTCONTABLE = '"& GF_FN2DTCONTABLE(pFecha) &"'"&_
             "  AND CDPRODUCTO = "& pCdProducto &_
             "  AND CDCLIENTE = "& pCdCliente &_
             "  AND CDSILO = '"& pCdSilo &"'"
    Call GF_BD_Puertos(pPto, rs, "EXEC", strSQL) 
    Set logMig = new classLog
    Call startLog(HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
    logMig.fileName = "MERMA_VOLATIL_"& pPto & "_"& left(session("MmtoDato"),8)
    Call logMig.info("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* INICIA TRANSACCION -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
    Call logMig.info("ELIMINA MERMA VOLATIL:")
    Call logMig.info("--> FECHA: "& GF_FN2DTE(pFecha))
    Call logMig.info("--> PRODUCTO: "& getDsProducto(pCdProducto))
    Call logMig.info("--> CLIENTE: "& getDsCliente(pCdCliente))
    Call logMig.info("--> SILO: "& pCdSilo)
End function
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_cdProducto,accion,g_cdSilo,g_cdCliente,g_fecha,logMig

accion       = GF_Parametros7("accion","",6)
g_strPuerto  = GF_Parametros7("Pto","",6)
g_cdProducto = GF_PARAMETROS7("cdProducto",0,6)
g_fecha      = GF_PARAMETROS7("fecha",0,6)
g_cdCliente  = GF_PARAMETROS7("cdCliente",0,6)
g_cdSilo     = GF_PARAMETROS7("cdSilo","",6)


SELECT CASE accion
    case ACCION_VISUALIZAR
        Call getSelectProducto(g_strPuerto)
    case ACCION_BORRAR
        Call eliminarMermaVolatil(g_strPuerto,g_fecha,g_cdProducto,g_cdCliente,g_cdSilo)
END SELECT

%>