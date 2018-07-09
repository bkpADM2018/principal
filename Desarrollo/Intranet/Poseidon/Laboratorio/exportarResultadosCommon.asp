<%
Const FIELD_TOKEN = ";"
Const KEY_TOKEN = "|"
Const MODO_MUESTRA = "M"
Const MODO_FECHA = "F"

Dim g_FechaDesde, g_FechaHasta, rs, conn, arch, fs,logMig,g_Pto,fname
Dim oDiccGruposEnsayosCamion,oDiccGruposEnsayos,strNamePathCabecera,strNamePathDetalle,strNamePathCuenta,muestraComercial,muestraBiotecnologia
Dim oDiccGruposEnsayosVagon,flagUTE,pathTempExp,auxProductoProteina, lstMuestras, modo, strNamePathError
'-----------------------------------------------------------------------------------------
'Se crea el archivo de un Segmento (por Fecha)
Function createFileSegment(pName)
	Dim fs, fadm
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(pName) Then  Call fs.deleteFile(pName, true)
	Set fadm = fs.CreateTextFile(pName)
	Set fs = nothing
	Set fadm = nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------
'Genera una cadena de datos necesarios para imprimir el reporte
Function imprimirDatosDetalle(pAceptacion,pCdProducto,pDtContable,pCoordinador,pCoordinado,pProducto,pCorredor,pEntregador,pVendedor,pProcedencia,pIdTransporte,pTipoTransporte,pChapa,pCartaPorte,pNeto,pMuestra)
    Dim auxIdentificacion
    if (CInt(pTipoTransporte) = TIPO_TRANSPORTE_CAMION) then
        auxIdentificacion = pChapa
    else
        auxIdentificacion = pIdTransporte
    end if                    
    imprimirDatosDetalle = pTipoTransporte & KEY_TOKEN &_
                           pAceptacion & KEY_TOKEN &_
                           pCdProducto & KEY_TOKEN &_
                           pDtContable & FIELD_TOKEN &_
                           pCoordinador & FIELD_TOKEN &_
                           pCoordinado & FIELD_TOKEN &_
                           pProducto & FIELD_TOKEN &_
                           pCorredor & FIELD_TOKEN &_
                           pEntregador & FIELD_TOKEN &_
                           pVendedor & FIELD_TOKEN &_
                           pProcedencia & FIELD_TOKEN &_
                           auxIdentificacion & FIELD_TOKEN &_
                           pCartaPorte & FIELD_TOKEN &_
                           pNeto & FIELD_TOKEN &_
                           pMuestra
End function 
'-----------------------------------------------------------------------------------------------------------------------
'--------------------------------------------COMIENZO DE LA PAGINA------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------

Set oDiccGruposEnsayos		  = createObject("Scripting.Dictionary")
Set oDiccGruposEnsayosCamion  = createObject("Scripting.Dictionary")

g_Pto		= GF_PARAMETROS7("pto", "", 6)
accion		= GF_PARAMETROS7("accion", "", 6)
fechaDesdeD = GF_PARAMETROS7("fecContableDS", "", 6)
fechaDesdeM = GF_PARAMETROS7("fecContableMS", "", 6)
fechaDesdeA = GF_PARAMETROS7("fecContableAS", "", 6)
lstMuestras = GF_PARAMETROS7("muestra", "", 6)
modo 		= GF_PARAMETROS7("modo", "", 6)
muestraComercial = GF_PARAMETROS7("muestraComercial", "", 6)
if ((muestraComercial = "on") or (modo = MODO_MUESTRA)) then muestraComercial = TIPO_AFIRMACION
muestraBiotecnologia  = GF_PARAMETROS7("muestraBiotecnologia", "", 6)
if ((muestraBiotecnologia = "on") or (modo = MODO_MUESTRA)) then muestraBiotecnologia = TIPO_AFIRMACION
'---------------------------------------------------------------------------------------------------
valParameterPath = Server.MapPath(".") & "\Archivos\Solicitudes"
'---------------------------------------------------------------------------------------------------
flagUTE = false
if (g_Pto = TERMINAL_PIEDRABUENA) then flagUTE = true

strNamePathCabecera = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_CABECERA
strNamePathReporte  = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_REPORTE
strNamePathError = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_ERROR
pathTempExp = valParameterPath & "\" & CAMARA_EXPORT_TEMP_REPORTE &"_"& g_Pto &".TXT"
if (not flagUTE) then
    strNamePathDetalle  = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_ANALISIS
    strNamePathCuenta   = valParameterPath &"\"& CAMARA_EXPORT_FILENAME_CUENTAYORDEN
end if


Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
fileNameLogExp = "EXPORTACION_CAMARA_"& Ucase(g_Pto) &"_" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)
logMig.fileName = fileNameLogExp

if(accion = ACCION_PROCESAR) then
    Call createFileSegment(strNamePathCabecera)
	Call createFileSegment(strNamePathError)
	Call createFileSegment(pathTempExp)
    if (not flagUTE) then
        Call createFileSegment(strNamePathDetalle)
	    Call createFileSegment(strNamePathCuenta)
    'else
    '    Call createFileSegment(pathTempExp)
    '    Call createFileSegment(strNamePathReporte)
    end if		
	logMig.info("--------------------- INCIANDO EXPORTACION ANALISIS DE CAMARA ------------------------")	
	logMig.info(" ---> USUARIO	: " & session("usuario"))
	logMig.info(" ---> PUERTO	: " & g_Pto)
	logMig.info(" ---> MODO		: " & modo)
end if


%>

