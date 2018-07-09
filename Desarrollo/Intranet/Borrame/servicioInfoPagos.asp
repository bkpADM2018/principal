<!--#include file="includes/procedimientosUnificador.asp"-->
<!--#include file="includes/procedimientospdf.asp"-->
<!--#include file="includes/procedimientosformato.asp"-->
<!--#include file="includes/procedimientosAS400.asp"-->
<!--#include file="includes/procedimientosRetenciones.asp"-->
<!--#include file="includes/procedimientosTraducir.asp"-->
<!--#include file="includes/procedimientosMG.asp"-->
<!--#include file="includes/procedimientosfechas.asp"-->
<!--#include file="includes/procedimientosMail.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<!--#include File="includes/cor-includePC.ASP" -->
<%
Function generarRetenciones(prov, fd, fh)

    Dim strSQL, rsRet, ret, filename, rtrn

    strSQL="Select MTCODE CodigoDetalle, MTFPAG FechaPago, MTNRET NroRetencion from tesfl.tes134f1 " &_
           " where MTFPAG >= " & fd & " and MTFPAG <= " & fh &_
           " and MTCODE in ('B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'P', 'Q', 'R', 'T', 'U') and MTNPRO = " & prov       
    Call GF_BD_AS400_2(rsRet, con, "OPEN", strSQL)
    ret = ""
    while (not rsRet.eof)
        filename = "RET_" & prov & "_" & rsRet("CodigoDetalle") & "_" & rsRet("NroRetencion") & "_" & rsRet("FechaPago") & ".pdf"    
        myRetPrint = server.mappath(".") & "\Temp\" & filename        
        Set Gbl_oPDF = GF_createPDF(myRetPrint)        
        rtrn = false
	    Select Case rsRet("CodigoDetalle")
		       Case "C": rtrn = GF_Retencion_C(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "E": rtrn = GF_Retencion_E(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "B": rtrn = GF_Retencion_B(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "H": rtrn = GF_Retencion_H(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "D": rtrn = GF_Retencion_D(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "G": rtrn = GF_Retencion_G(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "J": rtrn = GF_Retencion_J(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "I": rtrn = GF_Retencion_I(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "K": rtrn = GF_Retencion_K(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "L": rtrn = GF_Retencion_L(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "M": rtrn = GF_Retencion_M(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
		       Case "P": rtrn = GF_Retencion_P(rsRet("NroRetencion"), rsRet("CodigoDetalle"), rsRet("FechaPago"))
	    End Select	
	    if (rtrn) then	
	        if (ret <> "") then ret = ret & ";"
            ret = ret & myRetPrint
        end if            
	    Call GF_closePDF(Gbl_oPDF)
        rsRet.MoveNext()
    wend
    generarRetenciones = ret
    
End Function
'*********************************************************************************
Function getCuerpoMail(pProveedor, pList, pFecha)
    if (pList = "") then
        getCuerpoMail = "Hay pagos y retenciones del proveedor " & pProveedor & "-" & getDescripcionProveedor(pProveedor) & vbcrlf &_
	                    " que no pudieron ser enviadas ya que no tenemos mails. Se adjuntan los comprobantes."        
    else
          getCuerpoMail = "Señor/es: "& vbcrlf & getDescripcionProveedor(pProveedor) & vbcrlf & vbcrlf &_
             "Nos dirijimos a Usted a efectos de hacerle llegar adjuntos los comprobantes de retenciones y el resumen de pagos correspondiente al día : " & GF_FN2DTE(pFecha) & vbcrlf & vbcrlf &_					              
             "Atentamente."&vbcrlf&vbcrlf&_
             "Departamento de Tesoreria"&vbcrlf& getDescripcionProveedor(CD_TOEPFER)&vbcrlf&"Tel (011) 4317-0000"    
    end if	             
End Function
'*********************************************************************************    
Function getAsunto(pProveedor, pList, pFecha)
    if (pList = "") then
        getAsunto =  "PAGOS - El proveedor no tiene direcciones definidas: " & getDescripcionProveedor(pProveedor)
    else        
        getAsunto = getDescripcionProveedor(CD_TOEPFER) & " - Pagos del día " & GF_FN2DTE(pFecha)
    end if
End Function
'*********************************************************************************   
Function cargarMails(pProv)
    Dim vecMails(100), ret, i, finMails
        
    Call obtenerMailRetenciones(pProv, vecMails)
    ret = ""
    i=0
    finMails = false
    while (not finMails)
        if ((not isnull(vecMails(i))) and (vecMails(i) <> "")) then
            ret = ret & vecMails(i) & ";"            
        else
            finMails = true
        end if
        i = i + 1
        if (i=100) then finMails = true
    wend 
    cargarMails = ret
End Function
'*********************************************************************************   
Function checkProcessLock(pData)
    Dim fso, myFile, myFilename, myData
    
    checkProcessLock = false
    myFilename = server.MapPath(".") & "\servicioInfoPagos.lck"    
    Set fso = CreateObject("scripting.filesystemobject")        
	if (fso.FileExists(myFilename)) then	    
	    Set myFile = fso.OpenTextFile(myFilename,1,false)
        if (not myFile.AtEndOfStream) then
	        myData = myfile.ReadLine
	        if (myData <> pData) then checkProcessLock = true	        
        else
            checkProcessLock = true
        end if	        
	    myFile.Close
	    Set myFile = Nothing
	else
	    checkProcessLock = true
	end if
	Set fso = Nothing
End Function
'*********************************************************************************   
Function lockProcess(pData)
    Dim fso, myFile, myFilename
    
    myFilename = server.MapPath(".") & "\servicioInfoPagos.lck"
    Set fso = CreateObject("scripting.filesystemobject")    
	if (fso.FileExists(myFilename)) then fso.DeleteFile(myFilename)
    Set myFile = fso.CreateTextFile(myFilename, true)
    myFile.WriteLine(pData)
    myFile.Close
	Set myFile = Nothing
	Set fso = Nothing
	
End Function
'*********************************************************************************   
dim p_fechaDesde, p_fechaHasta, p_filename, p_prove
Dim myProvs, myAttach, myLck

'On Error Resume Next

Call GP_CONFIGURARMOMENTOS()

myLck = false
p_fechaDesde = GF_Parametros7("fd", "", 6)
if (p_fechaDesde = "") then 
    p_fechaDesde = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")
    myLck = true
end if        
p_fechaHasta = GF_Parametros7("fh", "", 6)
if (p_fechaHasta = "") then p_fechaHasta = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")
p_prove = GF_Parametros7("prov", "", 6)

if (checkProcessLock(p_fechaDesde)) then
    '---------------
    'Primero verifico que todos los proveedores esten en el TES960 (esto se hace ya que el proceso de retenciones requiere los datos migrados en ese archivo).
    'Eliminar este control si se deja de utilizar el archivo.
    strSQL="Select * from (Select DISTINCT MTPRET from TESFL.TES134F1 where MTFPAG >= " & p_fechaDesde & " and MTFPAG <= " & p_fechaHasta & " and MTCODE in ('B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'P', 'Q', 'R', 'T') "
    if (p_prove <> "") then   strSQL = strSQL & " and MTNPRO = " & p_prove
    strSQL= strSQL & ") P where not EXISTS (Select * from TESFL.TES960F1 where (WCPRET=MTPRET or WCNPRO=MTPRET))"
    Call executeQuery(rsAux, "OPEN", strSQL)
    if (not rsAux.eof) then
        auxMensaje = ""
        while (not rsAux.eof)
            auxMensaje = auxMensaje & rsAux("MTPRET") & vbcrlf                
            response.Write rsAux("MTPRET") & "<br>"        
            rsAux.MoveNext()
        wend    
        strDestinatario = "scalisij@toepfer.com"   
        Call GP_ENVIAR_MAIL("PAGOS - FALTAN DATOS DE PROVEEDORES" , auxMensaje, SENDER_TESORERIA, strDestinatario)
        response.Write "Faltan Proveedores en archivo TES960!"
        response.end
    end if
    '---------------
    'Se toman todos los proveedores que tuvieron pagos, si se paso uno por parametro, solo tomará ese.
    strSQL="Select MTNPRO from tesfl.tes134f1 " &_
           " INNER JOIN PROVFL.ACDSREL0 ON MTNING = DSQFNB " &_
           " where MTFPAG >= " & p_fechaDesde & " and MTFPAG <= " & p_fechaHasta
    if (p_prove <> "") then   strSQL = strSQL & " and MTNPRO = " & p_prove
    strSQL = strSQL & " and  MTCODE in ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'P', 'Q', 'R', 'T', 'U') "
    strSQL = strSQL & " group by MTNPRO"
    Call executeQuery(myProvs, "OPEN", strSQL)
    while (not myProvs.eof)
        myAttach = ""
        '----------------------------------------------------------------
        'Reporte de Pagos del día.
        p_filename = "PAGOS_" & myProvs("MTNPRO") & "_" & p_fechaDesde & ".pdf"    
        myReporte = server.mappath(".") & "\Temp\" & p_filename
        if (myAttach <> "") then myAttach = myAttach & ";"
        myAttach = myAttach & myReporte
        Set oPDF = GF_createPDF(myReporte)    
        Call PDFGirarHoja(90)
        Call GF_setPDFMODE(PDF_FILE_MODE)
        Call GF_generarPagosImpresion(oPDF,  myProvs("MTNPRO"), p_fechaDesde, p_fechaHasta)
        Call GF_closePDF(oPDF)
        '----------------------------------------------------------------
        'Cbtes de Retención.
        retAttach = generarRetenciones(myProvs("MTNPRO"), p_fechaDesde, p_fechaHasta)
        if (retAttach <> "") then
            if (myAttach <> "") then myAttach = myAttach & ";"
            myAttach = myAttach & retAttach
        end if        
        '----------------------------------------------------------------
        'Se envía por mail la info.
        strDestinatario = cargarMails(myProvs("MTNPRO"))
        
	    auxAsunto = getAsunto(myProvs("MTNPRO"), strDestinatario, p_fechaDesde)
	    auxMensaje = getCuerpoMail(myProvs("MTNPRO"), strDestinatario, p_fechaDesde)
        'Agrego los mails de Toepfer para control interno.
        strDestinatario = strDestinatario & cargarMails(CD_TOEPFER)
        'strDestinatario = "scalisij@toepfer.com"
        Call GP_ENVIAR_MAIL_ATTACHMENT(auxAsunto , auxMensaje, SENDER_TESORERIA, strDestinatario, myAttach)
        
        myProvs.MoveNext()
    wend
    if (myLck) then Call lockProcess(p_fechaDesde)
end if
response.Write "--- FIN ---"
%>

