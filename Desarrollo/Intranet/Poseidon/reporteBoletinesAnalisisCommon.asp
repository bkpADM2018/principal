<%
'El valor de estas Constantes separan los datos guardados en el Archivo.
Const FIELD_TOKEN    = ",*," ' Separador de Campos
Const SECTOR_TOKEN   = ";*;" ' Separa la Cabecera, detalle y Totales	
Const DETAIL_TOKEN	 = "&*&" ' Separa los datos del Detalle
Const VALUE_TOKEN    = "="	 
'-----------------------------------------------------------------------------------------
'Se crea el archivo de un Segmento (por Fecha)
Function createFileSegment(pName)
	Dim fs, fadm
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	strPath = Server.MapPath(pName)
	If fs.FileExists(strPath) Then  Call fs.deleteFile(strPath, true)	
	Set fadm = fs.CreateTextFile(strPath)	
	Set fs = nothing
	Set fadm = nothing	
End Function
'-----------------------------------------------------------------------------------------
'Obtiene la Descripcion de un grado
'	Grado 1 : 1
'	Grado 2 : 2
'	Grado 3 : 3
'	Grado 4 o mas : FE
Function getDsGrado(pGrado)
	Dim rtrn	
	rtrn = "FE"	
	Select case (pGrado)
		case GRADO_CAMARA_1:
			rtrn = "Grado 1"
		case GRADO_CAMARA_2:
			rtrn = "Grado 2"
		case GRADO_CAMARA_3:
			rtrn = "Grado 3"
	End Select
	getDsGrado = rtrn
End function
'-----------------------------------------------------------------------------------------------------------------------
Dim g_Pto,fname,g_Coordinador,g_Coordinado,g_Producto,g_FechaDesdeD,g_FechaDesdeM,g_FechaDesdeA,g_FechaHastaD,g_FechaHastaM,g_FechaHastaA,g_Sticker, g_Certificado, g_Calador,stringBoletines
Dim g_Grado, valCore, valGrado, g_FechaDesde, g_FechaHasta,valAcptacionCamara,oDiccAnalisis,auxGrado,key, valorGrado,netoAnterior,barraAnterior,fechaAnterior,recepcionAnterior
Dim turnoAnterior,productoAnterior,certificadoAnterior,empresaAnterior,clienteAnterior,corredorAnterior,vendedorAnterior,ctaPteAnterior,oDiccDetalle,dicRubros, arrRubros, cdRubro
Dim g_MaxSQBoletines,rs, conn, arch, fs, strPath,arrAlignDetalle,arrTitulosCabecera,arrTitulosDetalle,arrTitulosTotal,fileCode,arrTitulosCompletos, arrTitulosCamion, arrTitulosVisteo

'Inicializo titulos para claves de datos y titulos de reporte.
arrTitulosCabecera = Array("Fecha Descarga", "Certificado", "Sticker", "Producto" , "Coordinador/Coordinado", "Corredor/Vendedor", "Carta Porte")		
arrTitulosDetalle  = Array("Recepcion", "Turno", "Kg Neto", "Ensayo", "Resultado", "Calador", "Tipo", "Bon/Rebaja")
arrTitulosTotal    = Array("Total Neto", "Grado Analisis")
arrAlignDetalle    = Array("right", "right", "right", "left", "right", "center", "center", "right")

g_Pto		  = GF_PARAMETROS7("pto", "", 6)
g_Coordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
g_Coordinado  = GF_PARAMETROS7("cdCoordinado", "", 6)
g_Producto    = GF_PARAMETROS7("cmbCdProducto", 0, 6)
g_FechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
g_FechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
g_FechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
g_FechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
g_FechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
g_FechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
g_Sticker	  = GF_PARAMETROS7("sticker", "", 6)
g_Calador	  = GF_PARAMETROS7("cdCalador", "", 6)
g_Certificado = GF_PARAMETROS7("certificado", "", 6)
g_Grado		  = GF_PARAMETROS7("grado", 0, 6)
fileCode	  = GF_PARAMETROS7("fileCode", "", 6)
maxSegment	  = GF_PARAMETROS7("maxSegment", 0, 6)
accion		  =	 GF_PARAMETROS7("accion", "", 6)
g_strPuerto	  = g_Pto

if(accion <> ACCION_PROCESAR) then
	strName =  "Temp/BOLETINES_ANALISIS" & fileCode & ".txt"
	Call createFileSegment(strName)
end if

auxGrado = g_Grado
if auxGrado > GRADO_CAMARA_3 then auxGrado = GRADO_CAMARA_FE

%>
