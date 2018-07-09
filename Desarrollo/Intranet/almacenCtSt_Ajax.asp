<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Const ESTADO_PENDIENTE = 0
'-----------------------------------------------------------------------------------------
Function CambiarEstadoControlStockCab(pIdControl, pEstado)
	Dim strSQL,rs,oConn
	strSQL = "Update TBLCSTKCABECERA set IDESTADO=" & pEstado & " WHERE IDCONTROL = "& pIdControl	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end function
'*****************************************************************************************
'******************************** COMIENZO DE LA PAGINA **********************************
'*****************************************************************************************
Dim accion,idControl_new,myIdAlmacen

myIdAlmacen = GF_PARAMETROS7("IdAlmacen",0,6)
idControl_new = GF_PARAMETROS7("IdControl",0,6)
accion = GF_PARAMETROS7("accion","",6)
'------------------------------------------------------------------
Select case accion
	Case ACCION_BORRAR
		Call CambiarEstadoControlStockCab(idControl_new, ESTADO_BAJA)
	Case ACCION_ACTIVAR
		Call CambiarEstadoControlStockCab(idControl_new, ESTADO_ACTIVO)
	Case else 
		'si no es ni grabar ni borrar entonces devuelve una lista del resultado 
		'de Control de Stock, que ya esta cargado		
		if idControl_new > 0 then
			Set rs = leerDetallesCtSt(idControl_new,myIdAlmacen, true, "")	
			listOfCtSt=";"
			while (not rs.eof) 		
				vStock = rs("STOCKSISTEMA")
				if isNull(rs("STOCKSISTEMA"))then vStock = 0
				listOfCtSt = listOfCtSt & rs("idarticulo")&"|"&rs("dsarticulo")&"|"&rs("cdinterno")&"|"&vStock&"|"&rs("abreviatura")&";"
				rs.MoveNext()	
			wend
			listOfCtSt = left(listOfCtSt,Len(listOfCtSt)-1)
			Response.Write listOfCtSt		
		end if			
End Select
response.end		

%> 