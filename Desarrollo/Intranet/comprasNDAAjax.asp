<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<%
'Esta pagina es la encargada de la administracion de una Nota de Pedido, la cual recibe por medio del canal una determinada accion y en base a eso '
'realiza la actualizacion o la guarda'
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 23/03/2012
' Objetivo:
'			Actualiza el momento del ultimo envio que realizo una NDA y el usuario que la envio'
' Parametros:
'			[int]	pidPedido 
'			 [int] pIdCotizacion
' Devuelve:
'			---	
' Modificaciones:
'			18/10/2012 - CNA
'--------------------------------------------------------------------------------------------------
Function updateMmtoNotaAceptacion(pIdPedido,pIdCotizacion,pIdNDA)
	dim strSQL, rs, varNm
	strSQL = "Update TBLNOTACEPTACION set MMTOENVIO =" & session("MmtoSistema") &",CDUSRENVIO ='"& session("Usuario") &"' where IDNDA =" &pIdNDA
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 23/03/2012
' Objetivo:
'			Graba en la tabla Notaceptacion los registros y devuelve el IdNDA con el que se guardo
' Parametros:
'			[int]	pidPedido
'			[str]	pDsMensaje
'			[int]	pIdCotizacion
' Devuelve:
'			[int]	idNDA
' Modificaciones:
'			18/10/2012 - CNA
'--------------------------------------------------------------------------------------------------
Function guardarNotaAceptacion(pidPedido, pIdCotizacion, pDsMensaje)
	dim strSQL, cnn, rs
	strSQL = "INSERT INTO tblnotaceptacion (idpedido, idcotizacion, dsmensaje, CDUSRCARGA,CDUSRENVIO, MMTOENVIO, MOMENTO) "
	strSQL = strSQL & " VALUES ('"& pidPedido &"','"& pIdCotizacion &"','"& pDsMensaje &"','"& session("Usuario") &"','"& session("Usuario") &"','"& session("MmtoSistema") &"','"& session("MmtoSistema") &"')"	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = "SELECT MAX(IDNDA) as NDA from tblnotaceptacion "	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.eof)then Response.Write rs("NDA")
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 
' Objetivo:
'			Devuelve una lista de proveedores de un pedido 
' Parametros:
'			[int]	pidPedido
' Devuelve:
'			[int]	idNDA
' Modificaciones: -
'NOTA: ESTA FUNCION SE SITUA EN ESTA PAGINA DEBIDO A QUE ES LLAMADA DESDE AJAX Y UTILIZADA SOLO POR NOTA DE ACEPTACION
'--------------------------------------------------------------------------------------------------
Function getMailProvPedidoAjax(pidPedido, idCotizacion)
	Dim strSQL,listOfProveedores,myEmail
	strSQL = " SELECT distinct ctz.idproveedor,prov.NOMEMP as dsempresa,mail.email "
	strSQL = strSQL & "  FROM TBLCTZCABECERA ctz INNER JOIN [Database].[dbo].MET001A prov "
	strSQL = strSQL & "		 ON ctz.idproveedor = prov.NROEMP "
	strSQL = strSQL & "  LEFT JOIN TBLMAILSCOMPRAS mail ON mail.idempresa = prov.NROEMP "
	if (pidPedido > 0) then
		strSQL = strSQL & "  WHERE ctz.idpedido = " & pidPedido	
	else
		strSQL = strSQL & "  WHERE ctz.idcotizacion = " & idCotizacion	
	end if
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof		
		listOfProveedores = listOfProveedores & Trim(rs("dsempresa")) & "|" & Trim(rs("email")) &";"
		rs.MoveNext()
	wend	
	if (Len(listOfProveedores) > 0) then listOfProveedores = left(listOfProveedores,Len(listOfProveedores)-1)
	Response.Write listOfProveedores
End Function
'--------------------------------------------------------------------------------------------------
Dim  idPedido,accion,dsMensaje,idNDA

idPedido= GF_PARAMETROS7("idPedido",0,6)
idCotizacion = GF_PARAMETROS7("idCotizacion",0,6)
idNDA = GF_PARAMETROS7("IdNDA",0,6)
accion = GF_PARAMETROS7("accion",0,6)
dsMensaje = GF_PARAMETROS7("dsMensaje","" ,6)

dsMensaje = replace(dsMensaje, "|A|", "&")
dsMensaje = replace(dsMensaje, "|B|", "+")

'Establezco las claves de la tabla de acuerdo a las hiptesis asumidas.
if (idPedido <> 0)and(idCotizacion = 0) then idCotizacion=0

if(accion = ACCION_ACTUALIZAR_NDA)then 
	Call updateMmtoNotaAceptacion(idPedido,idCotizacion,idNDA)	
	response.end
end if
if ((accion = ACCION_IMPRIMIR_NDA) or (accion = ACCION_ENVIAR_NDA)) then 	
	call guardarNotaAceptacion(idPedido,idCotizacion,dsMensaje)
	response.end
else
	call getMailProvPedidoAjax(idPedido, idCotizacion)
	Response.end
end if


















%>