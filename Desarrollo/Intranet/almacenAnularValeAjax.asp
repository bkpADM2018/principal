<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<%
dim myIdVale, myIdPM, myCdVale
myIdVale = GF_Parametros7("idVale",0,6)
myIdAlmacen = GF_Parametros7("idAlmacen",0,6)
myIdPM = GF_Parametros7("idPM",0,6)
myCdVale = GF_Parametros7("cdVale","",6)
dim strSQL, rs, conn, rtrn, newIdVale, bOK
newIdVale = 1
bOK = true

'Controlar que se pueda sacar el stock correspondiente a la anulacion
if ((myCdVale = CODIGO_VS_ENTRADA) or (myCdVale = CODIGO_VS_DEVOLUCION) or (myCdVale = CODIGO_VS_RECEPCION) or (myCdVale = CODIGO_VS_AJUSTE_STOCK) or (myCdVale = CODIGO_VS_RECLASIFICACION_STOCK)) then 
	bOK = false
	if (puedeQuitarStock(myIdVale, myCdVale)) then bOK = true
end if

if bOK then
	

	'Crear vale de anulacion - Cabecera
	nroVale= getNumeracionVale(myIdAlmacen)
	strSQL = "INSERT INTO TBLVALESCABECERA (CDVALE,FECHA,CDSOLICITANTE,IDALMACEN,NRVALE,IDOBRA,CDUSUARIO,MOMENTO,PARTIDAPENDIENTE,IDBUDGETAREA,IDBUDGETDETALLE,ESTADO) " & _
			"(SELECT 'X' + SUBSTRING(CDVALE,2,4)," & left(session("MmtoDato"),8) & ",CDSOLICITANTE,IDALMACEN,'" & nroVale & "',IDOBRA,'" & session("Usuario") & "'," & session("MmtoDato") & ",PARTIDAPENDIENTE,IDBUDGETAREA,IDBUDGETDETALLE, " & ESTADO_ANULACION & " FROM TBLVALESCABECERA where IDVALE=" & myIdVale & ")"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = "Select IDVALE as IDVALE from TBLVALESCABECERA where NRVALE='" & nroVale & "'"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		newIdVale = rs("IDVALE")
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)

		'Modificar estado de vale
		strSQL = "UPDATE TBLVALESCABECERA SET ESTADO = " & ESTADO_BAJA & " where IDVALE=" & myIdVale
		'Response.write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
			
		'Crear vale de anulacion - Detalle
		strSQL = "INSERT INTO TBLVALESDETALLE(IDVALE, IDARTICULO, CANTIDAD, EXISTENCIA, SOBRANTE, VLUPESOS, VLUDOLARES) (SELECT " & newIdVale & ", IDARTICULO, CANTIDAD, EXISTENCIA, SOBRANTE, VLUPESOS, VLUDOLARES FROM TBLVALESDETALLE WHERE IDVALE=" & myIdVale & ")"
		'Response.write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)

		'Registrar relacion de anulacion
		strSQL = "INSERT INTO TBLVALESRELACIONES VALUES(" & myIdVale & "," & newIdVale & ")"
		'Response.write strSQL
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)

		'Analizar si se debe actualizar stock y/o saldo del pm
		select case myCdVale
			case CODIGO_VS_SALIDA, CODIGO_VS_PRESTAMO, CODIGO_VS_TRANSFERENCIA
				call ActualizarSaldoPM(myIdPM, myIdVale, "+")
				call ActualizarStockA(myIdVale, "EXISTENCIA", "+")
				call ActualizarStockA(myIdVale, "SOBRANTE", "+")
				if myCdVale = CODIGO_VS_SALIDA then call ActualizarPrecios(newIdVale, CODIGO_VS_SALIDA_X)
			case CODIGO_VS_ENTRADA
				call ActualizarStockA(myIdVale, "SOBRANTE", "-")
			case CODIGO_VS_DEVOLUCION, CODIGO_VS_RECEPCION
				call ActualizarStockA(myIdVale, "EXISTENCIA", "-")		
				call ActualizarStockA(myIdVale, "SOBRANTE", "-")		
				if myCdVale = CODIGO_VS_RECEPCION then call ActualizarPrecios(newIdVale, CODIGO_VS_RECEPCION_X)
			case CODIGO_VS_AJUSTE_PEDIDO
				call ActualizarSaldoPM(myIdPM, myIdVale, "+")
			case CODIGO_VS_AJUSTE_STOCK
				call ActualizarStockA(myIdVale, "EXISTENCIA", "-")
				call ActualizarStockA(myIdVale, "SOBRANTE", "-")
				'Si es ajuste de stock puede que el vale anulado sea de un control de stock. Se debe liberar el control tambien.
				strSQL = "Update TBLCSTKCABECERA set IDRESULTADO=0 where IDRESULTADO=" & myIdVale
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
			case CODIGO_VS_RECLASIFICACION_STOCK
				call ActualizarStockA(myIdVale, "EXISTENCIA", "-")
				call ActualizarStockA(myIdVale, "SOBRANTE", "-")
				call ActualizarPrecios(newIdVale, CODIGO_VS_RECLASIFICACION_STOCK_X)
			case CODIGO_VS_AJUSTE_VALE
				call ActualizarPrecios(newIdVale, CODIGO_VS_AJUSTE_VALE_X)
		end select

		Call GenerarFirmasXjs(myIdVale,newIdVale)
		
	end if
end if
'-----------------------------------------------------------------------------------------
sub GenerarFirmasXjs(pIdVale,pNuevoIdVale)
	Dim rs,oConn,strSQL

    'Las firmas deben ser las mismas que el del vale original, solo aquellos que ya hayana firmado, salvo por el responsable que es el que anula el vale.
    strSQL = "Insert into TBLVALESFIRMAS Select " & pNuevoIdVale & ", SECUENCIA, CDUSUARIO, FECHAFIRMA, HKEY  from TBLVALESFIRMAS where IDVALE=" & pIdVale & " and HKEY is not null" 
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    'Blanqueo las firmas.
	strSQL = "Update TBLVALESFIRMAS set FECHAFIRMA=null, HKEY=null where IDVALE=" & pNuevoIdVale
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
    'El responsable es quien cargó el vale.
	strSQL = "Update TBLVALESFIRMAS set CDUSUARIO='" & session("Usuario") & "', FECHAFIRMA=" & session("MmtoDato") & ", HKEY='" & A_MANO & "' where IDVALE=" & pNuevoIdVale & " and SECUENCIA=" & VS_FIRMA_RESPONSABLE
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------
sub ActualizarSaldoPM(idPM, idVale, signo)
	strSQL = "UPDATE PMD " & _
			 "	 SET PMD.SALDO = PMD.SALDO " & signo &  _
			" (SELECT CANTIDAD FROM TBLVALESDETALLE WHERE IDVALE=" & idVale & _
			" AND IDARTICULO=PMD.IDARTICULO) " & _
			" FROM TBLPMDETALLE PMD WHERE PMD.IDPEDIDO=" & idPM & _
			" AND EXISTS(SELECT idarticulo FROM TBLVALESDETALLE WHERE IDVALE=" & idVale & " AND IDARTICULO=PMD.IDARTICULO)"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)			
end sub
'-----------------------------------------------------------------------------------------
sub ActualizarStockA(idVale, columna, signo)
'JUNTAR EXISTENCIA Y SOBRANTE
	strSQL ="	UPDATE ART1 " & _
			"		SET ART1." & columna & " = ART1." & columna & signo &  _
			"	( " & _
			"		SELECT " & columna & " FROM TBLVALESDETALLE " & _
			"			WHERE IDVALE=" & idVale & " AND IDARTICULO=ART1.IDARTICULO" & _
			"	) " & _
			"	FROM TBLARTICULOSDATOS ART1 " & _ 
			"	WHERE ART1.IDARTICULO IN (SELECT IDARTICULO FROM TBLVALESDETALLE " & _
			"	WHERE IDVALE=" & idVale & ") and ART1.IDALMACEN=(SELECT IDALMACEN FROM TBLVALESCABECERA " & _
			"	WHERE IDVALE=" & idVale & ")"
			'Response.Write strSQL
call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)			
end sub
'-----------------------------------------------------------------------------------------
function puedeQuitarStock(pIdVale, pCdVale)
dim rs, oConn, strSQL, bOK, rtrn
puedeQuitarStock = false
bOK = true
strSQL = "SELECT DET.EXISTENCIA EXISTENCIA_VALE, DET.SOBRANTE SOBRANTE_VALE, ART.IDARTICULO, ART.DSARTICULO, DAT.EXISTENCIA EXISTENCIA_ALMA, DAT.SOBRANTE SOBRANTE_ALMA FROM TBLVALESCABECERA CAB INNER JOIN TBLVALESDETALLE DET ON CAB.IDVALE=DET.IDVALE INNER JOIN TBLARTICULOSDATOS DAT ON DET.IDARTICULO=DAT.IDARTICULO AND CAB.IDALMACEN=DAT.IDALMACEN INNER JOIN TBLARTICULOS ART ON DET.IDARTICULO=ART.IDARTICULO WHERE CAB.IDVALE= " & pIdVale
'Response.Write strSQL
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
while not rs.eof
		if pCdVale = CODIGO_VS_AJUSTE_STOCK then 
			if cdbl(rs("EXISTENCIA_VALE")) < 0 or cdbl(rs("SOBRANTE_VALE")) < 0 then bOK = false
		end if	
		if bOK then
			if (cdbl(rs("EXISTENCIA_VALE")) > cdbl(rs("EXISTENCIA_ALMA"))) then
				'rtrn = "No se puede quitar '" & rs("EXISTENCIA_VALE") & "' unidades de EXISTENCIA al articulo '" & rs("IDARTICULO") & " - " & rs("DSARTICULO") & "' debido a que actualmente en stock solo se cuenta con '" & rs("EXISTENCIA_ALMA") & "' unidades."
				rtrn = GF_Traducir("Imposible anular el " & pCdVale & " - " & pIdVale & " !<BR>El articulo '" & rs("IDARTICULO") & " - " & rs("DSARTICULO") & "' presenta una cantidad insuficiente de unidades.<BR>Se prentende anular '" & rs("EXISTENCIA_VALE") & "' unidades y en stock hay solo '" & rs("EXISTENCIA_ALMA") & "' - EXISTENCIA.")
			end if
			if (cdbl(rs("SOBRANTE_VALE")) > cdbl(rs("SOBRANTE_ALMA"))) then
				rtrn = rtrn & "<BR><BR>Se prentende anular '" & rs("SOBRANTE_VALE") & "' unidades y en stock hay solo '" & rs("SOBRANTE_ALMA") & "' - SOBRANTE."
			end if
			if rtrn <> "" then
				Response.Write rtrn
				exit function
			end if	
		end if
	rs.movenext
wend	
call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
puedeQuitarStock = true
end function
'-----------------------------------------------------------------------------------------

%>