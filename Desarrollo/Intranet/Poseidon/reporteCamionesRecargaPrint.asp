<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteCamionesRecargaCommon.asp"-->

<%
'-------------------------------------------------------------------------------------------------------------------------
Function dibujarFormato(pTitulo)
	Call GF_squareBox(oPDF,3,5,575 ,840,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("Images\kogge64.gif"),6, 840, 60, 60, 90)
	Call GF_setFont(oPDF,"ARIAL", 16 , FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF,27, 840, pTitulo, 840, PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
	Call GF_verticalLine(oPDF, 70, 10, 830)
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeVerticalText(oPDF,5, 840, GF_FN2DTE(session("MmtoSistema")), 830, PDF_ALIGN_RIGHT)
	Call GF_writeVerticalText(oPDF,15, 840, session("Usuario"), 830, PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",10,FONT_STYLE_NORMAL)
	Call GF_writeVerticalText(oPDF,580, 840, GF_TRADUCIR("Pagina") & " " & nroPagina, 830, PDF_ALIGN_RIGHT)
	call GF_setFont(oPDF,"COURIER",8,1)
end function
'-------------------------------------------------------------------------------------------------------------------------
Function armadoPDF(pto,fechaD,fechaH,pcdProducto, pDsProducto,pcdVendedor,pcdDestinatario,pcdCoordinado)
	Dim auxSMerma, cdProducto_old, flagInicio,totalMerma,totalTara,totalBruto, totalReg, i
	flagInicio = true
	Call dibujarFormato("REPORTE DE CAMIONES: RECARGAS")
	Call dibujarFiltros(pto,fechaD,fechaH,pcdProducto, pDsProducto,pcdVendedor,pcdDestinatario,pcdCoordinado)		
	i = 0 	
	if(Not rsRecarga.Eof)then
		totalReg = rsRecarga.RecordCount
		while Not rsRecarga.EoF
			i = i + 1
			auxSMerma = Cdbl(rsRecarga("Bruto")) - Cdbl(rsRecarga("Tara"))
			if px > PAGE_HEIGHT_SIZE and i <= totalReg then					
				Call nuevaPagina()
				cambioPagina = true
			end if
			if(Cdbl(rsRecarga("cdproducto")) <> cdProducto_old)then
				cdProducto_old = CDbl(rsRecarga("CDPRODUCTO"))
				if(not flagInicio)then				
					Call GF_squareBox(oPDF,pX + 2,12,10 ,825,0,"#E6E6E6", "#000000",0,PDF_SQUARE_NORMAL)		
					Call GF_writeVerticalText(oPDF, pX + 3, 837, "TOTAL", 600, PDF_ALIGN_CENTER)
					Call GF_writeVerticalText(oPDF, pX + 3, 193, GF_EDIT_DECIMALS(cdbl(totalBruto)*100,2), 61, PDF_ALIGN_RIGHT)
					Call GF_writeVerticalText(oPDF, pX + 3, 132, GF_EDIT_DECIMALS(cdbl(totalTara)*100,2), 60, PDF_ALIGN_RIGHT)
					totalMerma = totalBruto - totalTara
					Call GF_writeVerticalText(oPDF, pX + 3, 71 , GF_EDIT_DECIMALS(cdbl(totalMerma)*100,2), 61, PDF_ALIGN_RIGHT)
				end if
				totalMerma = 0
				totalTara  = 0
				totalBruto = 0				
				flagInicio = false				
				pX  = pX + 20
				call GF_setFont(oPDF,"COURIER",10,1)				
				Call GF_writeVerticalText(oPDF, pX, 837 , "PRODUCTO: " & rsRecarga("CDPRODUCTO") & " - " & rsRecarga("DSPRODUCTO"), 200, PDF_ALIGN_LEFT)				
				call GF_setFont(oPDF,"COURIER",8,1)
				pX  = pX + 15
				Call dibujarTitulos()
				px = px + 15
			end if
			if px > PAGE_HEIGHT_SIZE and i <= totalReg then					
				Call nuevaPagina()
				cambioPagina = true
			end if
			totalTara = totalTara + cdbl(rsRecarga("Tara"))
			totalBruto = totalBruto + cdbl(rsRecarga("Bruto"))
			Call GF_writeVerticalText(oPDF, pX + 4, 840 , rsRecarga("Remito"), 32, PDF_ALIGN_CENTER)
			Call GF_writeVerticalText(oPDF, pX + 4, 808, GF_FN2DTE(rsRecarga("Fecha")), 45, PDF_ALIGN_CENTER)
			Call GF_writeVerticalText(oPDF, pX + 4, 763, rsRecarga("Turno"), 35, PDF_ALIGN_RIGHT)
			Call GF_writeVerticalText(oPDF, pX + 4, 728, rsRecarga("IdCamion"), 54, PDF_ALIGN_CENTER)
			Call GF_writeVerticalText(oPDF, pX + 4, 674, GF_EDIT_CTAPTE(GF_nChars(rsRecarga("CP"), 16, "0", CHR_AFT)), 90, PDF_ALIGN_CENTER)
						
			auxCoordinado = rsRecarga("Coordinado")
			if(Len(Trim(auxCoordinado)) > 23)then auxCoordinado = Left(auxCoordinado,22) & ".."
			Call GF_writeVerticalText(oPDF, pX + 4, 582, auxCoordinado, 117, PDF_ALIGN_CENTER)
			
			auxDestinatario = rsRecarga("Destinatario")			
			if(Len(Trim(auxDestinatario)) > 23)then auxDestinatario = Left(auxDestinatario,22) & ".."						
			Call GF_writeVerticalText(oPDF, pX + 4, 467, auxDestinatario, 118, PDF_ALIGN_CENTER)
			
			auxVendedor = rsRecarga("Vendedor")			
			if(Len(Trim(auxVendedor)) > 23)then auxVendedor = Left(auxVendedor,22) & ".."			
			Call GF_writeVerticalText(oPDF, pX + 4, 349, auxVendedor, 118, PDF_ALIGN_CENTER)
			
			Call GF_writeVerticalText(oPDF, pX + 4, 231, GF_EDIT_PATENTE(rsRecarga("Chapa")), 38, PDF_ALIGN_CENTER)
			Call GF_writeVerticalText(oPDF, pX + 4, 193, GF_EDIT_DECIMALS(cdbl(rsRecarga("Bruto"))*100,2), 61, PDF_ALIGN_RIGHT)
			Call GF_writeVerticalText(oPDF, pX + 4, 132, GF_EDIT_DECIMALS(cdbl(rsRecarga("Tara"))*100,2), 60, PDF_ALIGN_RIGHT)
			Call GF_writeVerticalText(oPDF, pX + 4, 72 , GF_EDIT_DECIMALS(cdbl(auxSMerma)*100,2), 61, PDF_ALIGN_RIGHT)			
			px = px + 10
			rsRecarga.MoveNext()
		wend
		Call GF_squareBox(oPDF,pX + 2,12,10 ,825,0,"#E6E6E6", "#000000",0,PDF_SQUARE_NORMAL)		
		Call GF_writeVerticalText(oPDF, pX + 4, 837, "TOTAL", 600, PDF_ALIGN_CENTER)
		Call GF_writeVerticalText(oPDF, pX + 4, 193, GF_EDIT_DECIMALS(cdbl(totalBruto)*100,2), 61, PDF_ALIGN_RIGHT)
		Call GF_writeVerticalText(oPDF, pX + 4, 132, GF_EDIT_DECIMALS(cdbl(totalTara)*100,2), 60, PDF_ALIGN_RIGHT)
		totalMerma = totalBruto - totalTara
		Call GF_writeVerticalText(oPDF, pX + 4, 71 , GF_EDIT_DECIMALS(cdbl(totalMerma)*100,2), 61, PDF_ALIGN_RIGHT)		
		px = px + (SEPARATION * 2) - 2
		call GF_setFont(oPDF,"COURIER",10,1)
		Call GF_writeVerticalText(oPDF, px  , 840, "Fin del Reporte", 840, PDF_ALIGN_CENTER)
	else		
		Call GF_verticalLine(oPDF, px + 10, 35, 785)
		call GF_setFont(oPDF,"COURIER",10,1)
		Call GF_writeVerticalText(oPDF, px + 15 , 840, "No se econtraron resultados", 840, PDF_ALIGN_CENTER)
	end if		
	
End Function
'------------------------------------------------------------------------------------------------------------------------
function nuevaPagina()
	Call GF_newPage(oPDF)
	Call PDFGirarHoja(90)
	px = PAGE_TOP_INIT
	nroPagina = nroPagina + 1
	Call dibujarFormato("REPORTE DE CAMIONES: RECARGAS")	
	currentAuxY = PAGE_TOP_INIT 
end function
'------------------------------------------------------------------------------------------------------------------------
Function dibujarTitulos()	
	call GF_setFont(oPDF,"COURIER",8,1)
	Call GF_squareBox(oPDF,pX,808,15 ,32,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,763,15 ,45,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,728,15 ,35,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,674,15 ,54,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,584,15 ,90,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,467,15 ,117,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,349,15 ,118,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,231,15 ,118,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,193,15 ,38 ,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,132,15 ,61 ,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,72 ,15 ,60 ,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF,pX,11 ,15 ,61 ,0,"#517b4a", "#000000",1,PDF_SQUARE_NORMAL)
	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)	
	Call GF_writeVerticalText(oPDF, pX + 4, 840, "REMITO", 32, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 808, "FECHA", 45, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 763, "TURNO", 35, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 728, "CAMION", 54, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 674, "CTA. PTE.", 90, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 584, "COORDINADO", 117, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 467, "DESTINATARIO", 118, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 349, "VENDEDOR", 118, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 231, "CHAPA", 38, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 193, "BRUTO", 61, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 132, "TARA" , 60, PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF, pX + 4, 72 , "N. S/MERMA", 61, PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"COURIER",8,1)
	Call GF_setFontColor("#000000")
End Function
'------------------------------------------------------------------------------------------------------------------------
Function dibujarFiltros(pto,fechaD,fechaH,pcdProducto, pDsProducto,pcdVendedor,pcdDestinatario,pcdCoordinado)
	Dim auxProducto,auxDestinatario,auxVendedor,auxCoordinado,auxFechaD, yInicio,ySeparation, myFormatFecha
	
	ySeparation = PAGE_TOP_INIT
	call GF_setFont(oPDF,"COURIER",8,0)		
		
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835 , GF_TRADUCIR("Puerto........: ")	& UCASE(pto), 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION	
	
	myFormatFecha = Replace(fechaD,"-", "") 	
	myFormatFecha = GF_FN2DTE(myFormatFecha)
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835 , GF_TRADUCIR("Fecha Desde...: ") & myFormatFecha	, 100, PDF_ALIGN_LEFT)	
	ySeparation = ySeparation + SEPARATION
	
	myFormatFecha = Replace(fechaH,"-", "") 
	myFormatFecha = GF_FN2DTE(myFormatFecha)
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835, GF_TRADUCIR("Fecha Hasta...: ") & myFormatFecha	, 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxProducto = "Todos"
	if(pcdProducto > 0)then auxProducto = Trim(pcdProducto)&" - "&Trim(pDsProducto)	
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835, GF_TRADUCIR("Producto......: ") & auxProducto, 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxDestinatario = "Todos"
	if(pcdDestinatario > 0)then	auxDestinatario = Trim(pcdDestinatario)&" - "&Trim(getDsComprador(pcdDestinatario))	
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835, GF_TRADUCIR("Destinatario..: ") & auxDestinatario , 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxVendedor = "Todos"
	if(pcdVendedor > 0)then	auxVendedor = Trim(pcdVendedor)&" - "& Trim(getDsVendedor(pcdVendedor))	
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835, GF_TRADUCIR("Vendedor......: ") & auxVendedor, 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION
	
	auxCoordinado = "Todos"
	if(pcdCoordinado > 0)then	auxCoordinado = Trim(pcdCoordinado) &" - "& Trim(getDsCliente(pcdCoordinado))
	Call GF_writeVerticalTExt(oPDF, ySeparation, 835, GF_TRADUCIR("Coordinado....: ") & auxCoordinado	, 100, PDF_ALIGN_LEFT)
	ySeparation = ySeparation + SEPARATION		
	
	px = ySeparation
End Function
'******************************************************************************************************
'**************************************** COMIENZO DE PAGINA ******************************************
'******************************************************************************************************
Dim px
filename = "RECARGA_" & g_Puerto & session("MmtoSistema")
SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 542
PAGE_TOP_INIT = 82
nroPagina = 1

Set oPDF = GF_createPDF("PDFTemp")
Call PDFGirarHoja(90)
Call GF_setPDFMODE(PDF_STREAM_MODE)
Call armadoPDF(g_Puerto,g_fechaDesde,g_fechaHasta,g_Producto, dsProducto,g_Vendedor,g_Destinatario,g_Coordinado)
Call GF_closePDF(oPDF)

%>