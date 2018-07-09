<%
Const VISTEO_CALADA		= 2
Const FIELD_TOKEN		= ";"
Const VALUE_TOKEN		= "="
Const NO_DATA			= "XXX=0"
Const TOTAL_COLUMNAS    = 10
Const PREFIX_MUESTRA	="ZZMUESTRA"

'*** Variables globales de datos administrativos
Dim g_MaxSQCalada, g_Puerto, g_fechaDesde, g_fechaHasta, g_Producto, g_Vendedor, g_Corredor, g_Entregador, g_Cliente, g_cPorte, g_idCamion

'-----------------------------------------------------------------------------------------
'Se leen los datos administrativos requeridos entre etapas salvados en archivo por una generacion parcial anterior.
'Si no existe se devuelve los datos inicializados con valores default.
' pDic			: Default--> Diccionario creado pero vacio.
Function initDatosAdministrativos(ByRef pDic, pStrPath)
	Dim fs, fadm, txtLine, arrData
	
	Set pDic =Server.CreateObject("Scripting.Dictionary")
		
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	g_MaxSQCalada = 0

	if (fs.FileExists(pStrPath)) then
		Set fadm = fs.OpenTextFile(pStrPath, 1)	
		'Se leen los datos parametricos en el orden en que son grabados
		g_MaxSQCalada	= CInt(Trim(fadm.ReadLine()))
		g_Puerto		= Trim(fadm.ReadLine())
		g_fechaDesde	= Trim(fadm.ReadLine())
		g_fechaHasta	= Trim(fadm.ReadLine())
		g_Producto		= CInt(Trim(fadm.ReadLine()))
		g_Vendedor		= CInt(Trim(fadm.ReadLine()))
		g_Corredor		= CInt(Trim(fadm.ReadLine()))
		g_Entregador	= CInt(Trim(fadm.ReadLine()))
		g_Cliente		= CInt(Trim(fadm.ReadLine()))
		g_cPorte		= Trim(fadm.ReadLine())
		g_idCamion		= CInt(Trim(fadm.ReadLine()))
		g_Estado		= CInt(Trim(fadm.ReadLine()))
		g_Transporte    = CInt(Trim(fadm.ReadLine()))
		'Recupero los datos del diccionario de rubros.
		while (not fadm.AtEndOfStream)
			txtLine = fadm.ReadLine()
			arrData = Split(txtLine, VALUE_TOKEN)
			if (not pDic.Exists(cdRubro)) then pDic.Add arrData(0), Trim(arrData(1))
		wend
	end if
		
	Set fs = nothing
	Set fadm = nothing
		
End Function
'-----------------------------------------------------------------------------------------
'Se almacenan en archivo los datos necesarioa a pasar entre etapa y etapa y entre generaciones parciales.
Function saveDatosAdministrativos(pDic, pStrPath)
	Dim fs, fadm
	
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(pStrPath) Then  Call fs.deleteFile(pStrPath, true)
	Set fadm = fs.CreateTextFile(pStrPath)
	Set fs = nothing
	
	'Primero salvo la maxima secuencia de calada.
	fadm.WriteLine(g_MaxSQCalada)
	fadm.WriteLine(g_Puerto)
	fadm.WriteLine(g_fechaDesde)
	fadm.WriteLine(g_fechaHasta)
	fadm.WriteLine(g_Producto)
	fadm.WriteLine(g_Vendedor)
	fadm.WriteLine(g_Corredor)
	fadm.WriteLine(g_Entregador)
	fadm.WriteLine(g_Cliente)
	fadm.WriteLine(g_cPorte)
	fadm.WriteLine(g_idCamion)
	fadm.WriteLine(g_Estado)
	fadm.WriteLine(g_Transporte)
	'Segundo salvo los datos del diccionario de rubros.
	For each k in pDic.Keys
		fadm.WriteLine(k & VALUE_TOKEN & pDic(k))
	Next

	fadm.Close()
	
	Set fadm = nothing
	
End Function
'-----------------------------------------------------------------------------------------
Function getNombreApellidoCalada(pcdusername) 
	dim strSQL, rs, rtrn
	strSQL= "Select * from WFPROFESIONAL inner join MG on idProfesional=MG_KR where MG_KC='" & pcdusername & "' order by NOMBRE"			
	Call GF_BD_CONTROL(rs,conn,"OPEN",strSQL)	
	if (not rs.eof) then
		rtrn = rs("Nombre")
	else
		rtrn = pcdusername
	end if
	getNombreApellidoCalada = rtrn
end Function
'-----------------------------------------------------------------------------------------
Dim idcamion,nuCartaPorte,myWhere, strSQL, separaFecha1,separaFecha2, fileCode, rsDatos
Dim cdProducto,cdVendedor,dsVendedor,cdCorredor,dsCorredor,cdCliente,dsCliente,cdEntregador,dsEntregador
Dim RPT_Division, RPT_Month, RPT_Year, RPT_Filtro, RPT_accion, strPathAdm, rsMuestras, cdMuestra
Dim conn, filename, ultimaLinea, stringCamion, arch, fs, strPath, arrAlignCamion, arrAlignVisteo,arrTitulosExcel,arrAlignExcel
Dim dicRubros, arrRubros, cdRubro, arrTitulosCompletos, arrTitulosCamion, arrTitulosVagon, arrTitulosVisteo, g_Transporte

'Inicializo titulos para claves de datos y titulos de reporte.
arrTitulosCamion = Array("FECHA", "CAMION", "CARTA PORTE", "PRODUCTO", "CHASIS", "ACOPLADO", "CLIENTE", "CORREDOR", "VENDEDOR", "ENTREGADOR", "KILOS NETOS", "MERMA")
arrAlignCamion	= Array("center","center","center","left","center","center","left","center","center","left", "Right", "Right")

arrTitulosVagon = Array("FECHA", "OPERATIVO", "VAGON", "CARTA PORTE", "PRODUCTO", "CLIENTE", "CORREDOR", "VENDEDOR", "ENTREGADOR", "KILOS NETOS", "MERMA")
arrAlignVagon	= Array("center","center","center","center", "left","left","left","left","left","Right","Right")

arrTitulosVisteo = Array("F.VISTEO", "USUARIO", "PC TERMINAL")		
arrAlignVisteo	= Array("center","left","center")


'Se reciben los parametros.
g_Puerto = GF_PARAMETROS7("pto", "", 6)
g_strPuerto = g_Puerto

g_Producto = GF_PARAMETROS7("cdProducto", 0, 6)	
g_Vendedor = GF_PARAMETROS7("cdVendedor", 0, 6)
g_Corredor = GF_PARAMETROS7("cdCorredor", 0, 6)
g_Cliente = GF_PARAMETROS7("cdCliente", 0, 6)
g_Entregador = GF_PARAMETROS7("cdEntregador", 0, 6)
	
fecContableD = GF_PARAMETROS7("fecContableD", "", 6)
fecContableM = GF_PARAMETROS7("fecContableM", "", 6)
fecContableA = GF_PARAMETROS7("fecContableA", "", 6)
Call GF_STANDARIZAR_FECHA(fecContableD, fecContableM, fecContableA)


fecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
fecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
fecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)
Call GF_STANDARIZAR_FECHA(fecContableDH, fecContableMH, fecContableAH)

'g_fechaDesde = fecContableA & "-" & fecContableM & "-" & fecContableD
'g_fechaHasta = fecContableAH & "-" & fecContableMH & "-" & fecContableDH
g_fechaDesde = fecContableA & fecContableM & fecContableD
g_fechaHasta = fecContableAH & fecContableMH & fecContableDH
	
g_idCamion = GF_PARAMETROS7("idcamion", 0, 6)
if (g_idCamion <> 0) then g_idCamion = GF_nDigits(g_idCamion, 10)

nuCartaPorte1 = GF_PARAMETROS7("nuCartaPorte1", "", 6)
if (nuCartaPorte1 <> "") then nuCartaPorte1 = GF_nDigits(nuCartaPorte1, 4)
nuCartaPorte2 = GF_PARAMETROS7("nuCartaPorte2", "", 6)
if (nuCartaPorte2 <> "") then nuCartaPorte2 = GF_nDigits(nuCartaPorte2, 8)

g_cPorte = nuCartaPorte1 & nuCartaPorte2
g_Estado = GF_PARAMETROS7("estado", 0, 6)

g_Transporte = GF_PARAMETROS7("transporte", 0, 6)
if (g_Transporte = 0) then g_Transporte=TIPO_TRANSPORTE_CAMION

if (g_Transporte=TIPO_TRANSPORTE_CAMION) then
    arrTitulosExcel = arrTitulosCamion
    arrAlignExcel = arrAlignCamion
else
    arrTitulosExcel = arrTitulosVagon
    arrAlignExcel = arrAlignVagon
end if
'------------------------Nuevos Filtros----------------------
fileCode = GF_PARAMETROS7("fileCode", "", 6)
'if (session("Usuario") = "JAS") then fileCode = "PGS_1343405675950"

'Establezco la ruta y el nombre del archivo a crear/leer con los datos de los visteos.
strPath = Server.mapPath("../..") & "\temp\VISTEOS_CALADA_" & fileCode & ".txt"
'Se toma la lista de rubros y se ordena de menor a mayor para armar los títulos.
strPathAdm = Server.mapPath("../..") & "\temp\DATOS_ADM_" & fileCode & ".txt"

Call initDatosAdministrativos(dicRubros, strPathAdm)
g_strPuerto = g_Puerto
dsProducto = getDsProducto(g_Producto)
%>
