<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->

<%
Const CONV_KEY_PUERTO = "PUERTO"
Const CONV_KEY_PRODUCTO = "PRODUCTO"

Dim dicConv

Function obtenerCuposTerceros(pFechaDesde, pFechaHasta, pCliente, pPto)
    Dim strSQL, rs
    
    strSQL="Select * from CODIGOSCUPO CC " &_
            " inner join CLIENTES CL on CC.CUITCLIENTE=CL.NUCUIT " &_
            " where FECHACUPO >= " & pFechaDesde &_
            " and   CL.CDCLIENTE=" & pCliente
            if (pFechaHasta <> "") then strSQL = strSQL & " and   FECHACUPO <= " & pFechaHasta             
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)	
	Set obtenerCuposTerceros = rs
End Function
'---------------------------------------------------------------------------------------------------------
Function cargarTablaConversion(pCuitCliente, pPto)
    
    Dim strSQL, rs, ret, auxkey, auxval
    
    logMig.info("cargarTablaConversion - Inicia")
    
    Set dicConv = createObject("Scripting.Dictionary")    
    ret = false    
    strSQL="Select * from TBLCONVERSIONES where NUCUITCLIENTE='" & pCuitCliente& "'"
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)    
    while (not rs.eof)
        auxkey = rs("TIPODATO") & "_" & rs("CDPROPIO")
        auxval = rs("CDTERCERO")
        dicConv.Add auxkey, auxval
        ret = true
        rs.MoveNext()
    wend    
    cargarTablaConversion = ret
    
    logMig.info("cargarTablaConversion - Fin")
    
End Function
'---------------------------------------------------------------------------------------------------------
Function convertir(pTipo, pCodigoPropio)
    Dim rtrn
    
    rtrn=""
    if (dicConv.Exists(pTipo & "_" & pCodigoPropio)) then        
        rtrn = dicConv.Item(pTipo & "_" & pCodigoPropio)
    end if
    convertir = rtrn    
End Function
'---------------------------------------------------------------------------------------------------------
Function armarregistroDatos(myRs, pFilename)
    Dim fs, myFile, myPuerto, contador
    
    'Se abre el archivo de datos    
    logMig.info("Inicializando archivo de datos: " & pFilename)
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(pFilename) Then  Call fs.deleteFile(pFilename, true)
    Set myFile = fs.OpenTextFile(pFilename, 2, true)
    logMig.info("Archivo listo para trabajar.")        
    
    myPuerto = getNumeroPuerto(g_strPuerto)
    contador = 0
    while (not myRs.eof)       
        '1.- Código de Cupo (Alfanumérico. máx 11 posiciones)
        registro = myRs("CODIGOCUPO") & "|"
        '2.- Fecha del Cupo (Formato AAAAMMDD)
        registro = registro & myRs("FECHACUPO") & "|"
        '3.- Código de Producto (Según tabla ADM)
        registro = registro & GF_nDigits(convertir(CONV_KEY_PRODUCTO, myRs("cdProducto")), 3) & "|"
        '4.- Código de Puerto (Según tabla ADM)         
        registro = registro & GF_nDigits(convertir(CONV_KEY_PUERTO, myPuerto), 2)
        
        myFile.WriteLine registro        
        contador = contador + 1
        myRs.MoveNext()        
    wend    
    
    myFile.Close()
    
    Set myFile = Nothing
    Set fs = Nothing
    
    logMig.info("Se regsitraron " & contador & " cupos.")
End Function
'---------------------------------------------------------------------------------------------------------
Function enviarMailCuposTerceros(pPto, pFecha, pFileAttachment,idProveedor)
    Dim strBody, strSubject,fs,auxFileAtt, myLista
    
    
    'myLista = MAIL_TASK_INFO_LIST & getLetraPuerto(pPto)
    'strSubject = "Descargas " & getNombrePuerto(pPto) & " del " & GF_FN2DTE(pFecha)      
    strDestinosMail =  getStringMailsProveedor(PROV_ID_ADM)    
    if ((pFileAttachment <> "") and (strDestinosMail <> "")) then        
        logMig.info(" Enviando mail a " & strDestinosMail)    
        strBody = "Se envia adjunto el archivo con los cupos asigandos a partir del " & GF_FN2DTE(pFecha) & " para " & getNombrePuerto(pPto)
        if (pPto = TERMINAL_PIEDRABUENA) then strBody = strBody & vbcrlf & "IMPORTANTE: Los códigos de cupo de Terminal Piedrabuena son condicionales hasta completarse el proceso de nominación correspondiente."
        Call GP_ENVIAR_MAIL_ATTACHMENT("Alfred Toepfer S.R.L - Cupos Asignados - " & getNombrePuerto(pPto), strBody, SENDER_CUPOS_BA, strDestinosMail, pFileAttachment)
        'Call SendMail(TASK_POS_DESCARGA_TERCEROS, myLista, strSubject, strBody, pFileAttachment &";"& pFileAttachmentXLS)
    else
        logMig.info(" ERROR - No se puede enviar el mail, no hay archivo y/o no hay mails definidos. Destinatarios: " & strDestinosMail)    
    end if
End Function
'---------------------------------------------------------------------------------------------------------
'                                   ***** COMIENZA PAGINA *****
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Dim fechaDesde, fechaHasta, cliente, logMig, strSQL, rsDescarga, registro, dicAcond, g_strPuerto, myFilename


fechaDesde = GF_PARAMETROS7("fd", "", 6)
fechaHasta = GF_PARAMETROS7("fh", "", 6)
cliente = GF_PARAMETROS7("cl", "", 6)
g_strPuerto = GF_PARAMETROS7("pto", "", 6)

Call GP_ConfigurarMomentos()
session("usuario") = "SYNC"

if (fechaDesde = "") then fechaDesde = GF_DTEADD(Left(session("MmtoDato"), 8), 1, "D")
if (fechaHasta = "") then fechaHasta = "20991231"

dtDesde = GF_FN2DTCONTABLE(fechaDesde)
dtHasta = GF_FN2DTCONTABLE(fechaHasta)

myFilename = server.MapPath(".\Temp") & "\CUPOS_" & g_strPuerto & "_" & cliente & "_" & fechaDesde & ".txt"

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "EXPORTACION_CUPOS_"& Ucase(g_strPuerto) &"_" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)

logMig.info("--------------------- INCIANDO EXPORTACION ------------------------")	
logMig.info(" ---> PUERTO       : " & g_strPuerto)
logMig.info(" ---> FECHA DESDE  : " & dtDesde)
logMig.info(" ---> FECHA HASTA  : " & dtHasta)
logMig.info(" ---> CLIENTE      : " & cliente)
logMig.info("-------------------------------------------------------------------")	

Set rsDescarga = obtenerCuposTerceros(fechaDesde, "", cliente, g_strPuerto) 

if (not rsDescarga.eof) then
    
    'Se carga la tabla de conversiones para el cliente
    Call cargarTablaConversion(rsDescarga("CUITCLIENTE"), g_strPuerto)      
    
    'Se arma el registro de datos.
    Call armarregistroDatos(rsDescarga, myFilename)
            
    'obtengo la ruta donde se guarda el reporte(dentro del raiz actisaintra carpeta temp)    
    logMig.info("Ruta del archivo en disco: " & myFilename)
    
    'Se envia por mail.
    Call enviarMailCuposTerceros(g_strPuerto, fechaDesde, myFilename, cliente)       
else
	logMig.info("No se encontraron cupos para exportar")
end if

logMig.info("--------------------------- FIN PROCESO ---------------------------")	
%>