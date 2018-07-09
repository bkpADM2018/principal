<!--#include file="Includes/procedimientosUnificador.asp"-->
<%

Dim strSQL, rs


Function convertir(pTabla, pCampo, pClave)
	Call convertir4(pTabla, pCampo, pClave, "", "", "")
End Function

Function convertir2(pTabla, pCampo, pClave, pClave2)
	Call convertir4(pTabla, pCampo, pClave, pClave2, "", "")
End Function

Function convertir3(pTabla, pCampo, pClave, pClave2, pClave3)
	Call convertir4(pTabla, pCampo, pClave, pClave2, pClave3, "")
End Function

Function convertir4(pTabla, pCampo, pClave, pClave2, pClave3, pClave4)

	response.write "--************************************************--<br>"
	response.write "--	CONVERSION TABLA: " & pTabla & ", Campo: " & pCampo & "<br>"
	response.write "--************************************************--<br>"	
	strSQL = "Select " & pClave & " Clave, " 
	if (pClave2 <> "") then strSQL= strSQL & pClave2 & " Clave2, " 
	if (pClave3 <> "") then strSQL= strSQL & pClave3 & " Clave3, " 
	if (pClave4 <> "") then strSQL= strSQL & pClave4 & " Clave4, " 
	strSQL= strSQL & " nroemp Viejo, adm Nuevo from " & pTabla & " B inner join [Database].dbo.TMPNOBORRARMETANTERIOR A on A.nroemp = B." & pCampo 	
	strSQL= strSQL & " where B.IDORDER in (11075, 11076, 11077, 11078, 11079, 11080, 11081, 11082, 11083, 11084, 11085, 11086, 11087, 11088,  11089,  11090,  11091,  11092,  11093,  11094,  11095,  11096,   11097)"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	strSQL = "Select count(*) TOTAL from " & pTabla
	Call executeQueryDb(DBSITE_SQL_INTRA, rsC, "OPEN", strSQL)
	response.write "Registros a Convertir: " & rs.RecordCount & " de " & rsC("TOTAL") & " <br>"
	contar = 0
	while (not rs.eof)
		strSQL="Update " & pTabla & " Set " & pCampo & " = " & rs("Nuevo") & " where " & pClave & " = " & rs("Clave") 
		if (pClave2 <> "") then strSQL= strSQL & " and  " & pClave2 & " = " & rs("Clave2") 
		if (pClave3 <> "") then strSQL= strSQL & " and  " & pClave3 & " = " & rs("Clave3") 
		if (pClave3 <> "") then strSQL= strSQL & " and  " & pClave4 & " = " & rs("Clave4") 		
		strSQL= strSQL & "; -- Viejo: " & rs("Viejo") & ";"
		response.write strSQL & "<br>"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		contar = contar + 1
		rs.MoveNext()
	wend
	response.write "Ejecutadas OK: " & contar & "<br>"
	response.write "-- FIN --<br>"
End Function

'Call convertir("TBLCTZCABECERA", "IDPROVEEDOR", "IDCOTIZACION")
'Call convertir("TBLMAILSCOMPRAS", "IDEMPRESA", "IDEMPRESA")
'Call convertir("TBLOBRACONTRATOS", "IDPROVEEDOR", "IDCONTRATO")
'Call convertir2("TBLPCPDETALLE", "IDPROVEEDOR", "IDPEDIDO", "IDPROVEEDOR")
'Call convertir("TBLPCTCABECERA", "IDPROVEEDOR", "IDPEDIDO")
'Call convertir("TBLPCTCOTIZACIONES", "IDPROVEEDOR", "IDCOTIZACION")
'Call convertir2("TBLPCTPROVEEDORES", "IDPROVEEDOR", "IDPEDIDO", "IDPROVEEDOR")
'Call convertir2("TBLPROVEEDORESARCHIVOS", "IDPROVEEDOR", "IDPROVEEDOR", "SECUENCIA")
'Listo! - Call convertir("TBLPROVEEDORESCD", "IDPROVEEDOR", "IDPROVEEDOR")
'Call convertir("TBLREMCABECERA", "IDPROVEEDOR", "IDREMITO")
Call convertir("TBLSMORDER", "IDRESPONSABLECOMPANY", "IDORDER")
%>
