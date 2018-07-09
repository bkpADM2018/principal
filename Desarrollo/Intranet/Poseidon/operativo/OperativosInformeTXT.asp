<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->
<%

Const LENGTH_TITLE_FILTER = 120
'******************************************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function

'------------------------------------------------------------------------------------------------------------
Function dibujarFiltros()	
	Dim auxOperativo
	auxOperativo = "Todos"
	if myOperativo <> "" then auxOperativo = myOperativo 
	auxCartaPorte = "Todas"
	if myCartaPorte <> "" then auxCartaPorte = myCartaPorte
	arch.writeline("Operativo: " & auxOperativo &string(LENGTH_TITLE_FILTER-Cdbl(Len("Operativo: " & auxOperativo))," ") & "Carta de Porte:" &auxCartaPorte)
	auxTurno = "Todos"
	if myTurno <> "" then auxTurno = myTurno
	auxVagon = "Todos"
	if myIdVagon <> "" then auxVagon = myIdVagon
	arch.writeline("Turno: " & auxTurno & string(LENGTH_TITLE_FILTER-Cdbl(Len("Turno: " & auxTurno))," ") & "Vagon:" &auxVagon)
	arch.writeline("Fecha Inicio desde: " & GF_FN2DTE(myFecContableDesde) & string(LENGTH_TITLE_FILTER-Cdbl(Len("Fecha desde: " & GF_FN2DTE(myFecContableDesde)))," ") & "Fecha Inicio hasta: " &GF_FN2DTE(myFecContableHasta))
	auxCoordinado = "Todos"
	if myCdCoordinado <> "" then auxCoordinado = myDsCoordinado &"-"& getDsCliente(myDsCoordinado)
	auxProducto = "Todos"
	if myCdProducto <> "" then auxProducto = myCdProducto &"-"& getDsProducto(myCdProducto)
	arch.writeline("Coordinado: " & auxCoordinado & string(LENGTH_TITLE_FILTER-Cdbl(Len("Coordinado: " & auxCoordinado))," ") & "Producto: " &auxProducto)
	auxCorredor = "Todos"
	if myCdCorredor <> "" then auxCorredor = myCdCorredor &"-"& getDsCorredor(myCdCorredor)
	auxVendedor = "Todos"
	if myCdVendedor <> "" then auxVendedor = myCdVendedor &"-"& Trim(getDsVendedor(myCdVendedor))
	arch.writeline("Corredor: " & auxCorredor & string(LENGTH_TITLE_FILTER-Cdbl(Len("Corredor: " & auxCorredor))," ") & "Vendedor: " &auxVendedor)	
	auxEstado = "Todos"
	if myEstado <> 0 then auxEstado = getDsEstadoOperativo(myEstado,pto)
	arch.writeline("Estado: "&auxEstado)
End Function 
'------------------------------------------------------------------------------------------------------------
Function drawCabeceraOperativos(pStr)
	Dim myRegistro,h,str
	str = "" 
	myRegistro = Split(pStr, ";") 
	For h = 0 To UBound(myRegistro)
		aa = Cdbl(arrLengthTitleOperativo(h)) - Cdbl(len(myRegistro(h)))
		desc = myRegistro(h)
		if (aa < 0) Then 
		'La descripcion excede el tamaño deseado, se acorta la cadena
			desc = Left(myRegistro(h),(Cdbl(arrLengthTitleOperativo(h))-2)) & ".."
			aa = 0
		end if
		str = str & desc & string(aa," ")
	Next
	arch.writeline(str)
End Function

'-----------------------------------------------------------------------------------------------------------
Function createFileReport(pName)
	Dim fs, fadm
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	strPath = server.MapPath(pName)
	If fs.FileExists(strPath) Then  Call fs.deleteFile(strPath, true)	
	Set fadm = fs.CreateTextFile(strPath)	
	Set fs = nothing
	Set fadm = nothing	
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Call GP_CONFIGURARMOMENTOS()




Dim cont,index,arrLengthTitleOperativo,arrLengthTitleVagon,arch

arrLengthTitleOperativo = Array(15,25,25,20,25,25,25,25,25)

index = 0
pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
totalVagones = 0
totalNetoAcumulado = 0
g_strPuerto = pto
pTipo = GF_PARAMETROS7("pTipo", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)
maxSegment = GF_PARAMETROS7("maxSegment", 0, 6)
totalVagones = 0
totalKilosNetos = 0
call getParametros()
flagHayResultado = false
strName =  "../Temp/OPERATIVOS_" & session("Usuario") & ".txt"
Call createFileReport(strName)
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arch = fs.OpenTextFile(strPath, 8, true)
arch.writeline("Informe de Operativos")
arch.writeline(string(270,"-"))
arch.writeline("")
Call dibujarFiltros()	
arch.writeline(string(270,"-"))
arch.writeline("")
while index <= maxSegment
	pStrPath = Server.MapPath("../Temp/OPERATIVOS_" & session("Usuario") & "_" & index & ".txt")
	if (fs.FileExists(pStrPath)) then		
		Set fadm = fs.OpenTextFile(pStrPath, 1)			
		while (not fadm.AtEndOfStream)
			if not flagHayResultado then
				arch.writeline("")	
				strItems =  "Turno" &  string(arrLengthTitleOperativo(0)-len("Turno")," ") &_
							"Operativo" & string(arrLengthTitleOperativo(1)-len("Operativo")," ") &_
							"Carta de Porte" & string(arrLengthTitleOperativo(2)-len("Carta de Porte")," ") &_
							"Fecha Inicio"  & string(arrLengthTitleOperativo(3)-len("Fecha Inicio")," ") &_
							"Coordinado" & string(arrLengthTitleOperativo(4)-len("Coordinado")," ") &_
							"Producto" & string(arrLengthTitleOperativo(5)-len("Producto")," ") &_
							"Corredor" & string(arrLengthTitleOperativo(6)-len("Corredor")," ") &_							
							"Vendedor" & string(arrLengthTitleOperativo(7)-len("Vendedor")," ") &_														
							"Estado" & string(arrLengthTitleOperativo(8)-len("Estado")," ")
				arch.writeline(strItems)
				arch.writeline(string(270,"-"))
			end if	
			txtLine = fadm.ReadLine()
			Call drawCabeceraOperativos(txtLine)			
			flagHayResultado = true			
		wend
		Set fadm = nothing
		fs.DeleteFile(pStrPath)
	end if
	index = index + 1
wend
if not flagHayResultado then arch.writeline(string(100," ") & " No se encontraron resultados")
arch.close()
Response.Write strName
%>