<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="includes/procedimientosVIC.asp"-->
<%
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
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
totalVagones = 0
totalNetoAcumulado = 0
pTipo = GF_PARAMETROS7("pTipo", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)

call getParametros()
strSQL = generarSQL()
call GF_BD_Puertos(pto, rsGeneral, "OPEN", strSQL)

'Nombre del archivo
sFileRepInfoCamara = "vagones.txt"

set ObjFile = CreateObject("Scripting.FileSystemObject")
set file = objFile.OpenTextFile(server.MapPath("../temp/") & "/vagones.txt",2,true)
if not rsGeneral.eof then
	'Escribir Encabezados
	file.writeline("")
	file.writeline("Reporte de Vagones a Camara")
	file.writeline(string(270,"-"))
	file.writeline("")	
	file.writeline("Turno" &  string((10-len("Turno"))," ") & "Fecha" & string((11-len("Fecha"))," ")) & "Coordinador" & string((22-len("Coordinador"))," ") & "Coordinado"  & string((22-len("Coordinado"))," ") & "Producto"	  & string((22-len("Producto"))," ") & "Corredor"    & string((22-len("Corredor"))," ") & "Entregador"  & string((23-len("Entregador"))," ") & "Vendedor"    & string((23-len("Vendedor"))," ") & "Localidad"   & string((23-len("Localidad"))," ") & "Vagon"       & string((7-len("Vagon"))," ") & "Carta Porte" & string((17-len("Carta Porte"))," ") & "Netos"       & string((15-len("Netos"))," ") & "Barras"      & string((11-len("Barras"))," ") & "Grado"       & string((12-len("Grado"))," ") &  "Aceptacion"       & string((18-len("Aceptacion"))," ") & "Momento Descarga"       & string((20-len("Momento Descarga"))," ")
	file.writeline("")	
	CargarGrados
	while not rsGeneral.eof
			myKilosNetos = Clng(rsGeneral("Bruto"))-Clng(rsGeneral("Tara"))
			myGradoParticular =  VerGrado (pto, rsGeneral("cdProducto"), rsGeneral("cdAceptacion"), rsGeneral("Barras"), rsGeneral("Fecha"),myIncluir)
			If myGradoParticular <> "XXX" Then
				totalVagones = totalVagones + 1
				totalNetoAcumulado = totalNetoAcumulado + myKilosNetos
				call Sumar_Totales (myKilosNetos, totalVagones)
				call SumarResumen (myGradoParticular,myKilosNetos)
			End If
            file.writeline(rsGeneral("Turno") & string((10-len(rsGeneral("Turno")))," ") & GF_FN2DTE(Left(rsGeneral("DTPESADA"),8)) & string((11-len(GF_FN2DTE(Left(rsGeneral("DTPESADA"),8))))," ") & rsGeneral("Coordinador") & string((22-len(rsGeneral("Coordinador")))," ") & rsGeneral("Coordinado") & string((22-len(rsGeneral("Coordinado")))," ") & rsGeneral("Producto") & string((22-len(rsGeneral("Producto")))," ") & rsGeneral("Corredor") & string((22-len(rsGeneral("Corredor")))," ") & rsGeneral("Entregador") & string((23-len(rsGeneral("Entregador")))," ") & rsGeneral("Vendedor") & string((23-len(rsGeneral("Vendedor")))," ") & rsGeneral("Localidad") & string((23-len(rsGeneral("Localidad")))," ") & rsGeneral("NoVagon") & string((7-len(rsGeneral("NoVagon")))," ") & rsGeneral("CartaPorte") & string((17-len(rsGeneral("CartaPorte")))," ") & GF_EDIT_DECIMALS(myKilosNetos,0) & string((15-len(GF_EDIT_DECIMALS(myKilosNetos,0)))," ") & rsGeneral("Barras") & string((11-len(rsGeneral("Barras")))," ") & " " & myGradoParticular & string((12-len(myGradoParticular))," ") & trim(rsGeneral("Aceptacion")) &  string((18-len(trim(rsGeneral("Aceptacion"))))," ") & Left(GF_FN2DTE(rsGeneral("DTPESADA")),Len(GF_FN2DTE(rsGeneral("DTPESADA")))-3) & string((20-len(Left(GF_FN2DTE(rsGeneral("DTPESADA")),Len(GF_FN2DTE(rsGeneral("DTPESADA")))-3)))," "))
		rsGeneral.movenext
	wend	
	file.writeline(string(8,"-") & STRING(194," ") & STRING(10,"-"))
	file.writeline("TOTALES" &  string((10-len("TOTALES"))," ") & totalVagones &  string((10-len(totalVagones))," ") & "VAGONES" &  string((10-len("VAGONES"))," ") &  string(172," ") & GF_EDIT_DECIMALS(totalNetoAcumulado,0))
	file.writeline(string(8,"-") & STRING(194," ") & STRING(10,"-"))
		
	'Resumen
	file.writeline("")	
	file.writeline("RESUMEN")
	file.writeline(string(40," ") & "VAGONES" &  string((26-len("VAGONES"))," ") & "KILOGRAMOS" &  string((30-len("KILOGRAMOS"))," "))
	file.writeline("Items" &  string((40-len("Items"))," ") & "Cantidad" & string((15-len("Cantidad"))," ") & "%" &  string((11-len("%"))," ")  & "Cantidad" & string((15-len("Cantidad"))," ") & "%" &  string((11-len("%"))," "))
	file.writeline("")	
	call SumarPorcentajeResumen(totalVagones, totalNetoAcumulado, totalVagonesRegistrados, totalKilosNetosRegistrados)
	For i = 0 To 13
		file.writeline(myGrado(i,1) &  string((40-len(myGrado(i,1)))," ") &  string((8-len(myGrado(i,2)))," ") & myGrado(i,2) &  string((8-len(myGrado(i,3)))," ") & myGrado(i,3) & string((18-len(GF_EDIT_DECIMALS(cdbl(myGrado(i,4)),0)))," ") & GF_EDIT_DECIMALS(cdbl(myGrado(i,4)),0) & string((9-len(myGrado(i,5)))," ") & myGrado(i,5))
	Next
	file.writeline("")	
	file.writeline("TOTAL" &  string((40-len("TOTAL"))," ") & string((8-len(totalVagonesRegistrados))," ") & totalVagonesRegistrados & string((8-len(GF_EDIT_DECIMALS(10000,2)))," ") & GF_EDIT_DECIMALS(10000,2) & string((18-len(GF_EDIT_DECIMALS(totalKilosNetosRegistrados,0)))," ") & GF_EDIT_DECIMALS(totalKilosNetosRegistrados,0) & string((9-len(GF_EDIT_DECIMALS(10000,2)))," ") & GF_EDIT_DECIMALS(10000,2))
	
end if
file.close() 
myRuta = "../temp/" & sFileRepInfoCamara
Response.Write myRuta
%>