<%
'secuencia que le corresponde a cada firmante en la tabla de firmas de las planillas
Const PCP_FIRMA_RESPONSABLE 		= 0
Const PCP_FIRMA_MIEMBRO1	 		= 1
Const PCP_FIRMA_MIEMBRO2	 		= 2
Const PCP_FIRMA_MIEMBRO3	 		= 3
Const PCP_FIRMA_GTE_PUERTO 	 		= 4
Const PCP_FIRMA_GTE_SECTOR 	 		= 5
Const PCP_FIRMA_GTE_COMPRAS	 		= 6
Const PCP_FIRMA_SUP_PUERTOS	 		= 7
Const PCP_FIRMA_DIRECCION	 		= 8

Const PCP_TYPE_PURCHASE_SMALL   = 1
Const PCP_TYPE_PURCHASE_MEDIUM  = 2
Const PCP_TYPE_PURCHASE_LARGE   = 3

'---------------------------------------------------------------------------------------------
function addPCPItems(pIdPedido, pNroSobre, pIdProveedor, pCaracteristicas, pImporte, pMoneda, pCondPago, pFecEntrega)
'on error resume next
dim strSQL, rs, conn, rsIns, connIns, auxFec
addPCPItems = true
	if (InStr(1,pFecEntrega,"/") > 0) then
		auxFec = right(pFecEntrega,4) & mid(pFecEntrega,4,2) & left(pFecEntrega,2)
	else
		auxFec = "0"
	end if	
	strSQL="Select * from TBLPCPDETALLE where idPedido = " & pIdPedido & " and IdProveedor=" & pIdProveedor
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if rs.eof then
		strSQL = "Insert Into TBLPCPDETALLE VALUES(" & pIdPedido & "," & pNroSobre & "," & pIdProveedor & ",'" & pCaracteristicas & "'," & pImporte & ",'" & pMoneda & "','" & pCondPago & "'," & auxFec & ")"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
	else
		strSQL = "Update TBLPCPDETALLE set NroSobre=" & pNroSobre & ", Caracteristicas='" & pCaracteristicas & "',Importe=" & pImporte & ",CDMONEDA='" & pMoneda & "' ,CondPago='" & pCondPago & "', FecEntrega = " & auxFec & " where idPedido = " & pIdPedido & " and IdProveedor=" & pIdProveedor 
        Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "UPDATE", strSQL)
	end if	
	'Response.Write strSQL
	'Call GF_BD_COMPRAS(rs, conn, "CLOSE", strSQL)
'if err.number > 0 then addPCPItems = false
end function
'---------------------------------------------------------------------------------------------
function addPCPCabecera(pIdPedido, pComentarios, pObservaciones)
'on error resume next
dim strSQL, rs, conn, rsIns, connIns,myObs,mycoment
addPCPCabecera = true

	strSQL="Select * from TBLPCPCABECERA where idPedido = " & pIdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	'Se reemplaza el caracter de enter por <br>
	mycoment = replace(pComentarios,chr(10),ENTER_SYMBOL)
	mycoment = replace(mycoment,"'","*")
	if (left(mycoment,4) = ENTER_SYMBOL) then	mycoment = mid(mycoment,5,len(mycoment))		
	myObs = replace(pObservaciones,chr(10),ENTER_SYMBOL)
	myObs = replace(myObs,"'","*")
	if (left(myObs,4) = ENTER_SYMBOL) then	myObs = mid(myObs,5,len(myObs))		
	if rs.eof then
		strSQL = "Insert Into TBLPCPCABECERA VALUES(" & pIdPedido & ",'" & mycoment & "','" & myObs & "')"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
	else
		strSQL = "Update TBLPCPCABECERA set COMENTARIOS='" & mycoment & "', OBSERVACIONES='" & myObs & "' where idPedido = " & pIdPedido
        Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "UPDATE", strSQL)
	end if	
	'Response.Write strSQL
	'Call GF_BD_COMPRAS(rs, conn, "CLOSE", strSQL)
'if err.number > 0 then addPCPCabecera = false
end function
'---------------------------------------------------------------------------------------------
Function adminPCPFirmas(pIdPedido, pSecuencia, pCdUsuario)
	Dim rs, strSQL, conn
	
	'Cargo los nuevos firmantes
	strSQL="Select * from TBLPCPFIRMAS where IDPEDIDO=" & pIdPedido & " and SECUENCIA=" & pSecuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then
		if (pCdUsuario <> "") then 
            strSQL="Insert into TBLPCPFIRMAS(IDPEDIDO, SECUENCIA, CDUSUARIO) values (" & pIdPedido & ", " & pSecuencia & ", '" & pCdUsuario & "')"		
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
        end if
	else
		if (pCdUsuario <> "") then
			strSQL="Update TBLPCPFIRMAS set CDUSUARIO='" & pCdUsuario & "', FECHAFIRMA=null, HKEY=null where IDPEDIDO=" & pIdPedido & " and SECUENCIA= " & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		else
			strSQL="Delete from TBLPCPFIRMAS where IDPEDIDO=" & pIdPedido & " and SECUENCIA=" & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	end if
End Function
'------------------------------------------------------------------------------------------
Function puedeModificarPlanilla(idPedido, estado)
	Dim rs, conn, strSQL, rtrn
	rtrn = false
	if ((idPedido > 0) and (checkControlPCT())) then
		if ((estado >= ESTADO_PCT_ABIERTO) and (estado <= ESTADO_PCT_APROBADO)) then
			strSQL = strSQL &	"SELECT IDCOTIZACION FROM TBLCTZCABECERA "
			strSQL = strSQL &	" WHERE ESTADO<> '" & CTZ_ANULADA & "' AND IDPEDIDO = " & idPedido
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (rs.eof) then
				rtrn = true
			end if
		end if
	end if
	puedeModificarPlanilla = rtrn
End Function
'-----------------------------------------------------------------------------------------------
Function getPCPAuthorizationType(pImporte, pMoneda)
    Dim limiteFirmaSupervisor, limiteFirmaDireccion
    Dim myImporte, unidadCD
    
    'Se traen los limites de firma para uso y controles.
    limiteFirmaSupervisor = CDbl(getValorNorma("VLPCPSP"))
    limiteFirmaDireccion = CDbl(getValorNorma("VLPCPDR"))
    unidadCD = getUnidadNorma("VLPCPSP")
        
    'Se transforma el importe a la moneda de la regal de auditoria. (Importe con 2 decimales!)
    myImporte = pImporte    
	if (pMoneda <> unidadCD) then
		if (pMoneda = MONEDA_PESO) then	
			myImporte = round(myImporte / getTipoCambio(MONEDA_DOLAR, ""),2)
		else
			myImporte = round(myImporte * getTipoCambio(MONEDA_DOLAR, ""),2)
		end if
	end if	
	
	'Controlo el importe contra los limites
	if (Cdbl(myImporte) < CDbl(limiteFirmaSupervisor)) then getPCPAuthorizationType = PCP_TYPE_PURCHASE_SMALL
	if ((Cdbl(myImporte) >= Cdbl(limiteFirmaSupervisor)) and (myImporte < limiteFirmaDireccion)) then getPCPAuthorizationType = PCP_TYPE_PURCHASE_MEDIUM
	if (Cdbl(myImporte) >= Cdbl(limiteFirmaDireccion)) then getPCPAuthorizationType = PCP_TYPE_PURCHASE_LARGE

End Function    
%>