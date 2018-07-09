<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->

<%
dim idArticulo, idCotizacion
dim strSQL, rs

idCotizacion = GF_Parametros7("idCotizacion",0,6)
idArticulo = GF_Parametros7("idArticulo",0,6)

'Se elimina el ajuste
strSQL = "Update TBLCTZDETALLE set IMPORTEPESOSCREDITO=0, IMPORTEDOLARESCREDITO=0 where IDCOTIZACION=" & idCotizacion & " and IDARTICULO=" & idArticulo
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
%>