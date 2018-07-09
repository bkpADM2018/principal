<%
'El valor de estas Constantes separan los datos guardados en el Archivo.
Const FIELD_TOKEN    = ",*," ' Separador de Campos
Const SECTOR_TOKEN   = ";*;" ' Separa la Cabecera, detalle y Totales	
Const DETAIL_TOKEN	 = "&*&" ' Separa los datos del Detalle
Const TOTAL_TOKEN	 = "&|&" ' Separa el total
Const VALUE_TOKEN    = "="	 ' Separa el valor del campo
Const REPORTE_CAMIONES = "REPORTE_CAMIONES"
Const REPORTE_VAGONES  = "REPORTE_VAGONES"
Const REPORTE_TOTAL_CAMIONES = "TOTAL_CAMIONES"
Const REPORTE_TOTAL_VAGONES  = "TOTAL_VAGONES"
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
'-------------------------------------------------------------------------------------------------------------------
Dim fname,conn, arch, fs, strPath,arrTitulosCamiones,arrTitulosVagones,arrTitulosRubro,fileCode
Dim g_cdUsuario,g_dsUsuario,g_cdCoordinado,g_dsCoordinado,g_cdProducto,g_fechaDesdeD,g_fechaDesdeM,g_fechaDesdeA
Dim g_fechaHastaD,g_fechaHastaM,g_fechaHastaA,g_cdCorredor,g_dsCorredor,g_cdVendedor,g_dsVendedor,g_chkCamiones
Dim g_chkVagones,g_cdAceptacion,g_cdRubro,g_minimo,g_maximo,g_chkPromediar,g_accion,g_chkResumen
Dim g_fechaDesde,g_fechaHasta,strNameCam,strNameVag,arrRubro,dicTest,arrTitulosTotales,g_AbRubro,g_VlRubro,g_DsRubro

'Inicializo titulos para claves de datos y titulos de reporte.
arrTitulosCamiones = Array("Fecha", "Camion", "Carta Porte", "Patente" , "Coordinado", "Corredor", "Vendedor", "Procedencia", "Producto", "Usuario", "Terminal", "Calidad", "Grado", "Neto sin Merma", "Neto con Merma", "Kg Merma", "Bruto", "Tara")
arrTitulosVagones  = Array("Fecha", "Operativo", "Carta Porte", "Vagon" , "Coordinado", "Corredor", "Vendedor", "Procedencia", "Producto", "Usuario", "Terminal", "Calidad", "Grado", "Neto sin Merma", "Neto con Merma", "Kg Merma", "Bruto", "Tara","Proteina")
arrTitulosRubros   = Array("Codigo Rubro", "Abr. Rubro", "Valor","Desc. Rubro")
arrTitulosTotales  = Array("Abreviatura", "Descripcion", "Prom. Pond.")

'POSICIONES                   0 |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 |  10|  11|  12|  13|  14|  15|  16| 17
arrPositionTitulos  = array( 840, 780, 710, 640, 580, 480, 345, 210, 70 , 835, 760, 650, 580, 480, 363, 240, 150, 80)
arrWidthTitulos     = array( 60 , 60 , 60 , 60 , 80 , 120, 120, 120, 40 , 60 , 60 , 60 , 90 , 60 , 60 , 50 , 55 , 55)

g_Pto		  = GF_PARAMETROS7("pto", "", 6)
g_accion	= GF_PARAMETROS7("accion", "", 6)
g_cdUsuario = GF_PARAMETROS7("cdUsuario", "", 6)
g_dsUsuario = GF_PARAMETROS7("dsUsuario", "", 6)
g_cdCoordinado = GF_PARAMETROS7("cdCoordinado", 0, 6)
g_dsCoordinado = GF_PARAMETROS7("dsCoordinado", "", 6)
g_cdProducto  = GF_PARAMETROS7("cmbCdProducto", 0, 6)
g_FechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
g_FechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
g_FechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
g_FechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
g_FechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
g_FechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
g_cdCorredor = GF_PARAMETROS7("cdCorredor", 0, 6)
g_dsCorredor = GF_PARAMETROS7("dsCorredor", "", 6)
g_cdVendedor = GF_PARAMETROS7("cdVendedor", 0, 6)
g_dsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)
g_chkCamiones = GF_PARAMETROS7("chkCamiones", 0, 6)
g_chkVagones  = GF_PARAMETROS7("chkVagones", 0, 6)
g_cdAceptacion  = GF_PARAMETROS7("cdAceptacion", "", 6)
g_cdRubro  = GF_PARAMETROS7("cmbRubro", 0, 6)
g_minimo   = GF_PARAMETROS7("minimo", "", 6)
g_maximo   = GF_PARAMETROS7("maximo", "", 6)
g_chkPromediar  = GF_PARAMETROS7("chkPromediar", 0, 6)
g_chkResumen    = GF_PARAMETROS7("chkResumen" , 0, 6)
fileCode	  = GF_PARAMETROS7("fileCode", "", 6)
maxSegment	  = GF_PARAMETROS7("maxSegment", 0, 6)
g_strPuerto	  = g_Pto

if(g_accion <> ACCION_PROCESAR) then	
	strName =  "Temp/REPORTE_CALADOR" & fileCode & ".txt"
	Call createFileSegment(strName)	
end if

%>
