<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<%
dim myIdArticulo, myIdAlmacen
dim strSQL, rs, conn, rtrn

myIdArticulo = GF_Parametros7("idArticulo",0,6)
myIdAlmacen = GF_Parametros7("idAlmacen",0,6)

rtrn = 0
strSQL="Select existencia from TBLARTICULOSDATOS TAD inner join TBLARTICULOS TA on TAD.IDARTICULO=TA.IDARTICULO  where TAD.IDALMACEN=" & myIdAlmacen & " and TA.IDARTICULO=" & myIdArticulo
'Response.writge strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs2, oConn, "OPEN", strSQL)
if not rs.eof then rtrn = rs("EXISTENCIA")
Call executeQueryDB(DBSITE_SQL_INTRA, rs2, oConn, "CLOSE", strSQL)
Response.Write rtrn
%>