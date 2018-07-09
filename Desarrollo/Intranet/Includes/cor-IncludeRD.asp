<%
sub GF_OBTENER_RESUMEN(p_rs, ByRef p_orkr,ByRef p_arkr,ByRef rec, ByRef apli, ByRef stock, ByRef id)

Dim oConn,strSQL,rs
Dim KRUsuario,ORKR,ARKR

'Obtengo el KR del usuario.
GF_MGC "SG",session("Usuario"),KRUsuario,""
'Obtengo las recepciones.
strSQL="Select * from cor_ResumenCorredor where CORKR=" & KRUsuario
strSQL= strSQL & " and MmtoDato <=" & ConvertirDate2Segundos(date() & " 23:59:59") 
strSQL= strSQL & " and MmtoDato >=" & ConvertirDate2Segundos(date() & " 00:00:00") 
strSQL= strSQL & " and Tipo=1 and ORKR=" & p_rs("ORKR") & " and ARKR=" & p_rs("ARKR")
'Response.Write strSQL
GF_BD_CONTROL rs,oConn,"OPEN",strSQL
if not(rs.EOF) then
	rec=rs("intCantidad")
else
	rec=0
end if
'Obtengo las aplicaciones.
strSQL="Select * from cor_ResumenCorredor where CORKR=" & KRUsuario
strSQL= strSQL & " and MmtoDato <=" & ConvertirDate2Segundos(date() & " 23:59:59") 
strSQL= strSQL & " and MmtoDato >=" & ConvertirDate2Segundos(date() & " 00:00:00") 
strSQL= strSQL & " and Tipo=2 and ORKR=" & p_rs("ORKR") & " and ARKR=" & p_rs("ARKR")
'Response.Write strSQL
GF_BD_CONTROL rs,oConn,"OPEN",strSQL
if not(rs.EOF) then
	apli=rs("intCantidad")	
else
	apli=0
end if
'Obtengo los datos de los clientes.
P_orkr=p_rs("ORKR")
P_arkr=p_rs("ARKR")
stock=p_rs("intCantidad")
id=p_rs("HeadID")
'Response.Write ConvertirDate2Segundos(GF_VerFechaDato()) 
end sub
'---------------------------------------------------------
sub GF_OBTENER_STOCKS(ByRef rs)

Dim oConn,strSQL,KRUsuario

'Obtengo el KR del usuario.
GF_MGC "SG",session("Usuario"),KRUsuario,""
'Obtengo el stock
strSQL="Select distinct ORKR,ARKR,intCantidad,HeadID from cor_ResumenCorredor_MmtoDato where CORKR=" & KRUsuario
strSQL= strSQL & " and MmtoDato <=" & ConvertirDate2Segundos(GF_VERFECHADATO()) 
strSQL= strSQL & " and Tipo=3"
'response.Write strSQL
'Response.Write ConvertirDate2Segundos(GF_VerFechaDato()) 
GF_BD_CONTROL rs,oConn,"OPEN",strSQL

end sub
'------------------------------------------------------------
sub GF_RECEPCIONES_SS(byRef rs)

Dim oConn,strSQL,strSQLLista
Dim KRUsuario

'Obtengo el KR del usuario.
GF_MGC "SG",session("Usuario"),KRUsuario,""
'Armo la lista de productos que ya se listo
strSQLLista="Select distinct ARKR from cor_ResumenCorredor_MmtoDato where CORKR=" & KRUsuario
strSQLLista= strSQLLista & " and MmtoDato <=" & ConvertirDate2Segundos(GF_VERFECHADATO()) 
strSQLLista= strSQLLista & " and Tipo=3"
'Obtengo las recepciones.
strSQL="Select * from cor_ResumenCorredor where CORKR=" & KRUsuario
strSQL= strSQL & " and MmtoDato <=" & ConvertirDate2Segundos(date() & " 23:59:59") 
strSQL= strSQL & " and MmtoDato >=" & ConvertirDate2Segundos(date() & " 00:00:00") 
strSQL= strSQL & " and Tipo=1 and not ARKR in (" & strSQLLista & ")" 
'Response.Write strSQL
GF_BD_CONTROL rs,oConn,"OPEN",strSQL

end sub
'-------------------------------------------------------------
sub GF_APLICACIONES_SS(p_rs, ByRef P_strDSOrganizacion,ByRef P_strDSProducto,ByRef rec, ByRef apli)

Dim oConn,strSQL,rs
Dim KRUsuario,ORKR,ARKR

'Obtengo el KR del usuario.
GF_MGC "SG",session("Usuario"),KRUsuario,""
'Obtengo las aplicaciones.
strSQL="Select * from cor_ResumenCorredor where CORKR=" & KRUsuario
strSQL= strSQL & " and MmtoDato <=" & ConvertirDate2Segundos(date() & " 23:59:59") 
strSQL= strSQL & " and MmtoDato >=" & ConvertirDate2Segundos(date() & " 00:00:00") 
strSQL= strSQL & " and Tipo=2 and ORKR=" & p_rs("ORKR") & " and ARKR=" & p_rs("ARKR")
'Response.Write strSQL
GF_BD_CONTROL rs,oConn,"OPEN",strSQL
if not(rs.EOF) then
	apli=rs("intCantidad")	
else
	apli=0
end if
'Obtengo los datos de los clientes.
GF_MGC "","",p_rs("ORKR"),P_strDSOrganizacion
GF_MGC "","",p_rs("ARKR"),P_strDSProducto
rec=p_rs("intCantidad")
'Response.Write ConvertirDate2Segundos(GF_VerFechaDato()) 
end sub
'-------------------------------------------------------------
sub GF_OBTENER_DETALLE_STO(ByRef rs, P_ID)

Dim strSQL,oConn

strSQL="Select * from cor_DetalleStock where HeadID=" & P_ID
GF_BD_CONTROL rs,oConn,"OPEN",strSQL

end sub
'-------------------------------------------------------------
sub GF_DETALLE_STOCK(rs,ByRef intRecibo,ByRef dteFecha,ByRef strOrigen,ByRef Zar,ByRef Hum,ByRef Cal,ByRef intKgDesc,ByRef intPendiente,ByRef intCarta,ByRef Chasis,ByRef Acoplado,ByRef Cosecha)

Dim KCCamion

intRecibo=rs("Recibo")
dteFecha=left(ConvertirSegundos2Date(rs("Fecha")),10)
strOrigen=rs("Origen")
Zar=GF_EDIT_DECIMALS(cDbl(rs("Zar")), 2)
Hum=GF_EDIT_DECIMALS(cDbl(rs("Hum")), 2)
Cal=GF_EDIT_DECIMALS(cDbl(rs("Cal")), 2)
intKgDesc=rs("KgDesc")
intPendiente=rs("Pendiente")
intCarta=rs("CartaPorte")
GF_MGC "",KCCamion,rs("Camion"),""
Chasis=GF_DT1("READ","CACHASIS","","","CA",KCCamion)
Acoplado=GF_DT1("READ","CAACOP","","","CA",KCCamion)
Cosecha=rs("Cosecha")
end sub
'---------------------------------------------------------
sub GF_DETALLE_RECIBO(P_RECIBO,dteFecha,strOrigen,Zar,Hum,Cal,intKgDesc,intCarta,Chasis,Acoplado,strCosecha,strDSOrganizacion,strDSProducto)					

Dim strSQL, oConn, rs, rs2

strSQL= "Select * from cor_DetalleStock where Recibo=" & P_RECIBO
GF_BD_CONTROL rs,oConn,"OPEN",strSQL
'Completo los campos a devolver
GF_DETALLE_STOCK rs,P_RECIBO,dteFecha,strOrigen,Zar,Hum,Cal,intKgDesc,"",intCarta,Chasis,Acoplado,strCosecha
strSQL ="Select * from cor_ResumenCorredor where HeadID=" & rs("HeadID")
GF_BD_CONTROL rs2,oConn,"OPEN",strSQL
GF_MGC "","",rs2("ORKR"),strDSOrganizacion
GF_MGC "","",rs2("ARKR"),strDSProducto

end sub
'---------------------------------------------------------
%>