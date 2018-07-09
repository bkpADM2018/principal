<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosmail.asp"-->
<!--#include file="../Includes/procedimientosunificador.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="../Includes/procedimientossql.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/includeGeneracionArchivos.asp"-->
<!--#include file="interfacturas.asp"-->
<% 
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim idProveedor, dsProveedor, accion, dtDesde, dtHasta, strPath
dim strSQL, rs, strWhere, conFSO, arch, aux_nroFactura, flagDatos
dim fecCbte, fecVto, tipoCbte, tipoCbteAB, ptoVta, nroFactura, cuitCorredor, cuitVendedor, cuitEmisor, ctoToepfer, ctoCorredor
dim subTotalGrav, subTotalNoGrav, tasaIVA, porIVA, percIVA, percIIBB, total, nroCAE, fecCAE, ctaPte, kgDescarga
dim kgMerma, porHumedad, precGasto, tipoGasto, moneda


idProveedor = GF_PARAMETROS7("idProveedor", 0, 6)
dsProveedor = GF_PARAMETROS7("dsProveedor", "", 6)
dtDesde = GF_PARAMETROS7("dtDesde", "", 6)
dtHasta = GF_PARAMETROS7("dtHasta", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)

myUsuario = session("Usuario")
if not isToepfer(session("KCOrganizacion")) then myUsuario = FAC_USER_WEB

Call executeSP(rs, "TFFL.TF100F1_GET_FACTURAS_BY_PARAMETERS", FAC_AUTORIZADA & "||" & GF_DTE2FN(dtDesde) & "||" & GF_DTE2FN(dtHasta) & "||" & idProveedor & "||" & myUsuario & "||" & SEC_SYS_FACTURACION)

Set confile = createObject("Scripting.FileSystemObject")

strPath = server.MapPath("..") & "\temp\facturas-" & idProveedor & ".txt"

'Response.Write strPath
Set fich = confile.CreateTextFile(strPath) 

flagDatos = false
while not rs.eof
		fecCbte = verNull(rs("FECHACBTE"),"N")
		fecVto = verNull(rs("FechaVto"),"N")
		tipoCbte = verNull(rs("TipoCbte"),"N")
		tipoCbteAB = verNull(rs("TipoABCbte"),"T")
		ptoVta = verNull(rs("PtoVta"),"N")
		nroFactura = verNull(rs("NroCbte"),"N")
		cuitCorredor = verNull(rs("CuitCor"),"N")
		cuitVendedor = verNull(rs("CuitVen"),"N")
		cuitEmisor = CUIT_TOEPFER
		ctoToepfer = verNull(rs("CONTRATO"),"T")
		ctoCorredor = verNull(rs("CtoCorredor"),"T")
		moneda = MONEDA_PESO
		if (rs("CdMoneda") <> MONEDA_PESO) then moneda = MONEDA_DOLAR
		subTotalGrav = verNull(rs("SubTotalGravado"),"N")
		subTotalNoGrav = verNull(rs("NoGravado"),"N")
		tasaIVA = cDbl(verNull(rs("TasaIVA"),"N"))*100
		porIVA = verNull(rs("IVA"),"N")
		percIVA = verNull(rs("PercepcionIVA"),"N")
		percIIBB = verNull(rs("PercepcionIIBB"),"N")
		total = verNull(rs("Total"),"N")
		tipoCambio = cDbl(verNull(rs("TipoCambio"),"N"))*1000
		nroCAE = trim(verNull(rs("NroCAE"),"N"))
		fecCAE = verNull(rs("FechaCAE"),"N")
		ctaPte = replace(verNull(rs("CtaPte"),"N"),"-","")
		kgDescarga = verNull(rs("KilosDescarga"),"N")
		kgMerma = verNull(rs("KilosMerma"),"N")
		porHumedad = cDbl(verNull(rs("PorcentajeHumedad"),"N"))*100
		precGasto = cDbl(verNull(rs("PrecioGasto"),"N"))*100
		tipoGasto = verNull(rs("TipoGasto"),"N")
		aux = string(tFECHA - len(fecCbte)," ") & fecCbte
		aux = aux & string(tFECHA - len(fecVto),"0") & fecVto
		aux = aux & string(tTCBT - len(tipoCbte),"0") & tipoCbte
		aux = aux & string(tTCBT - len(tipoCbteAB)," ") & tipoCbteAB
		aux_nroFactura = GF_nDigits(ptoVta,4) & GF_nDigits(nroFactura,8)
		aux = aux & string(tNUM - len(trim(aux_nroFactura)),"0") & trim(aux_nroFactura)
		aux = aux & string(tCUIT - len(cuitCorredor),"0") & cuitCorredor
		aux = aux & string(tCUIT - len(cuitVendedor),"0") & cuitVendedor
		aux = aux & string(tCUIT - len(cuitEmisor),"0") & cuitEmisor
		aux = aux & string(tCtos - len(ctoToepfer)," ") & ctoToepfer
		aux = aux & string(tCtos - len(ctoCorredor)," ") & ctoCorredor
		aux = aux & cdMoneda
		aux = aux & string(tIMP - len(subTotalGrav),"0") & subTotalGrav
		aux = aux & string(tIMP - len(subTotalNoGrav),"0") & subTotalNoGrav
		aux = aux & string(tPORC - len(tasaIVA),"0") & tasaIVA
		aux = aux & string(tIMP - len(porIVA),"0") & porIVA
		aux = aux & string(tIMP - len(percIVA),"0") & percIVA
		aux = aux & string(tIMP - len(percIIBB),"0") & percIIBB
		aux = aux & string(tIMP - len(total),"0") & total
		aux = aux & string(tCambio - len(tipoCambio),"0") & tipoCambio
		aux = aux & string(tCAE - len(nroCAE),"0") & nroCAE
		aux = aux & string(tFECHA - len(fecCAE),"0") & fecCAE
		aux = aux & string(tCTAPTE - len(ctaPte),"0") & ctaPte
		aux = aux & string(tKILOS - len(kgDescarga),"0") & kgDescarga
		aux = aux & string(tKILOS - len(kgMerma),"0") & kgMerma
		aux = aux & string(tPORC - len(porHumedad),"0") & porHumedad
		aux = aux & string(tIMP - len(precGasto),"0") & precGasto
		aux = aux & string(tTIPO - len(tipoGasto),"0") & tipoGasto
		fich.WriteLine(aux) 		
		flagDatos = true
	rs.MoveNext()
wend

fich.close() 
Set confile = nothing
Set fich = nothing
if (flagDatos) then
	if (accion = ACCION_BACH) then
		dirMail = getMailFacturacionProveedores(idProveedor, FACTURACION_LISTA_MAIL_ARCHIVO)
		Response.Write "Enviar de " & SENDER_FACTURACION & " a " & dirMail	
	    if (dirMail <> "") then	
            auxAsunto  = GF_TRADUCIR(getDescripcionProveedor(CD_TOEPFER) & " - Facturación periodo " & dtDesde & " al " & dtHasta)
		    auxMensaje = "Se adjunta el archivo con las facturas emitidas en el período " & dtDesde & " al " & dtHasta & "." &vbcrlf&vbcrlf &_
		                 "Atentamente."&vbcrlf&vbcrlf&_
					     "Departamento de Tesoreria"&vbcrlf& getDescripcionProveedor(CD_TOEPFER)&vbcrlf&"Tel (011) 4317-0000"
	        Call GP_ENVIAR_MAIL_ATTACHMENT(auxAsunto, auxMensaje, SENDER_FACTURACION, dirMail, strPath)			
		end if
	else
		Call Descargar(strPath)
	end if
end if    
'---------------------------------------------------------------------------------------------------------------------------------
function verNull(pValue,pTipo)
dim rtrn 
rtrn = trim(pValue)
	if isNull(rtrn) then 
		rtrn = " "
		if pTipo = "N" then	rtrn = "0"
	end if	
	verNull = rtrn
end function

%>
<html>
<head>
<% if not flagDatos then %>
    <script>parent.sinDatos();</script>
<% end if %>    
</head>
</html>
