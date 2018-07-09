<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<%
Const PAGE_LAST_LINE = 550

Function getCamionesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
	Dim strSQL, rs
	
	strSQL = "Select FORMAT(HCD.DTCONTABLE, 'dd/MM/yyyy') DTCONTABLE, " & _
			"		HCD.NUCARTAPORTE, " & _
			"		'' CDVAGON, " & _
			"		HCD.CDCLIENTE, " & _ 
			"		HCD.CDCORREDOR, " & _
			"		HCD.CDVENDEDOR, " & _
			"		HCD.CDPRODUCTO, " & _
			"		HCD.BRUTO, " & _
			"		HCD.TARA, " & _
			"		HCD.MERMA, " & _
			"		case when MV.VLMERMAKILOS is Null then 0 else MV.VLMERMAKILOS end MERMAVOLATIL, " & _
			"		P.DSPRODUCTO, " & _
			"		CL.DSCLIENTE, " & _
			"		C.DSCORREDOR, " & _
			"		V.DSVENDEDOR " & _
			"from VWHCAMIONESDESCARGA_MERMATARABRUTO HCD " & _
			"		left join MERMAVOLATIL MV on HCD.DTCONTABLE=MV.DTCONTABLE and HCD.NUCARTAPORTE=MV.NUCARTAPORTE and HCD.IDCAMION=MV.IDTRANSPORTE " & _
			"		inner join Productos P on P.CDPRODUCTO=HCD.CDPRODUCTO " & _
			"		inner join Clientes CL on CL.CDCLIENTE=HCD.CDCLIENTE " & _
			"		inner join Corredores C on C.CDCORREDOR=HCD.CDCORREDOR " & _
			"		inner join Vendedores V on V.CDVENDEDOR=HCD.CDVENDEDOR " & _
			"where 	HCD.DTCONTABLE >= '" & pDtDesde & "' and HCD.DTCONTABLE <= '" & pDtHasta & "'" & _
			"		and HCD.CDESTADO in ('" & CAMIONES_ESTADO_EGRESADOOK & "', '" & CAMIONES_ESTADO_PESADOTARA & "')"
	if (pCdCliente <> 0) then strSQL = strSQL & " and HCD.CDCLIENTE = " & pCdCliente
	if (pCdCorredor <> 0) then strSQL = strSQL & " and HCD.CDCORREDOR = " & pCdCorredor
	if (pCdVendedor <> 0) then strSQL = strSQL & " and HCD.CDVENDEDOR = " & pCdVendedor	 
	if (pCdProducto <> 0) then strSQL = strSQL & " and HCD.CDPRODUCTO = " & pCdProducto
	strSQL = strSQL & "Order by " & _
			"	HCD.DTCONTABLE,	" & _
			"	HCD.CDCLIENTE, " & _ 
			"	HCD.CDCORREDOR, " & _
			"	HCD.CDVENDEDOR, " & _
			"	HCD.CDPRODUCTO, " & _
			"	HCD.NUCARTAPORTE "
	Call executeQueryDB(pPto, rs, "OPEN", strSQL)
	Set getCamionesDescargados = rs

End Function
'------------------------------------------------------------------------------------------------------------------------
Function getVagonesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)	
	Dim strSQL
	strSQL = getSQLVagones("VWVAGONES_MERMATARABRUTO", pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
	Call executeQueryDB(pPto, rs, "OPEN", strSQL)
	Set getVagonesDescargados = rs
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getHVagonesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)	
	Dim strSQL
	strSQL = getSQLVagones("VWHVAGONES_MERMATARABRUTO", pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
	Call executeQueryDB(pPto, rs, "OPEN", strSQL)
	Set getHVagonesDescargados = rs
End Function
'------------------------------------------------------------------------------------------------------------------------

Function getSQLVagones(pSPVagones, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
	Dim strSQL, rs
	
	strSQL = "Select FORMAT(HCD.DTCONTABLE, 'dd/MM/yyyy') DTCONTABLE, " & _
			"		CONCAT(HCD.NUCARTAPORTESERIE, LEFT(HCD.NUCARTAPORTE, 8)) NUCARTAPORTE, " & _
			"		HCD.CDVAGON CDVAGON, " & _
			"		HCD.CDCLIENTE, " & _ 
			"		HCD.CDCORREDOR, " & _
			"		HCD.CDVENDEDOR, " & _
			"		HCD.CDPRODUCTO, " & _
			"		HCD.BRUTO, " & _
			"		HCD.TARA, " & _
			"		HCD.MERMA, " & _
			"		case when MV.VLMERMAKILOS is Null then 0 else MV.VLMERMAKILOS end MERMAVOLATIL, " & _
			"		P.DSPRODUCTO, " & _
			"		CL.DSCLIENTE, " & _
			"		C.DSCORREDOR, " & _
			"		V.DSVENDEDOR " & _
			"from " & pSPVagones & " HCD " & _
			"		left join MERMAVOLATIL MV on HCD.DTCONTABLE=MV.DTCONTABLE and CONCAT(HCD.NUCARTAPORTESERIE, LEFT(HCD.NUCARTAPORTE, 8))=MV.NUCARTAPORTE and HCD.CDVAGON=MV.IDTRANSPORTE " & _
			"		inner join Productos P on P.CDPRODUCTO=HCD.CDPRODUCTO " & _
			"		inner join Clientes CL on CL.CDCLIENTE=HCD.CDCLIENTE " & _
			"		inner join Corredores C on C.CDCORREDOR=HCD.CDCORREDOR " & _
			"		inner join Vendedores V on V.CDVENDEDOR=HCD.CDVENDEDOR " & _
			"where HCD.DTCONTABLE >= '" & pDtDesde & "' and HCD.DTCONTABLE <= '" & pDtHasta & "'"
	if (pCdCliente <> 0) then strSQL = strSQL & " and HCD.CDCLIENTE = " & pCdCliente
	if (pCdCorredor <> 0) then strSQL = strSQL & " and HCD.CDCORREDOR = " & pCdCorredor
	if (pCdVendedor <> 0) then strSQL = strSQL & " and HCD.CDVENDEDOR = " & pCdVendedor	 
	if (pCdProducto <> 0) then strSQL = strSQL & " and HCD.CDPRODUCTO = " & pCdProducto
	strSQL = strSQL & "Order by " & _
			"	HCD.DTCONTABLE,	" & _
			"	HCD.CDCLIENTE, " & _ 
			"	HCD.CDCORREDOR, " & _
			"	HCD.CDVENDEDOR, " & _
			"	HCD.CDPRODUCTO, " & _
			"	HCD.NUCARTAPORTE, " & _
			"	HCD.CDVAGON "
	getSQLVagones = strSQL
	
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawPageHeader()
	Call GF_writeImage(oPDF, Server.MapPath("..\Images\logo1.jpg"),10, 850, 81, 75, 90)
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)
	Call GF_writeVerticalText(oPDF,10, 110, GF_FN2DTE(session("MmtoSistema")), 100, PDF_ALIGN_RIGHT)
	Call GF_writeVerticalText(oPDF,20, 110, session("Usuario"), 100, PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"ARIAL", 14 , FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF,40, 580, "REPORTE DE DESCARGAS", 300, PDF_ALIGN_CENTER)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawTableTitle(pY)
	Dim iniX, gap
	
	'Call GF_squareBox(oPDF, pY, 30, 22,785,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_verticalLine(oPDF, pY+22, 10, 825)
	Call GF_setFont(oPDF,"ARIAL", 10, FONT_STYLE_BOLD)	
	inix = 830
	Call GF_writeVerticalText(oPDF, pY + 5, iniX		, "F.Desc."	, 60, PDF_ALIGN_CENTER) 	
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 80	, "C. Porte", 100, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 150	, "Vagon"	, 100, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 195, "Cliente"	, 100, PDF_ALIGN_LEFT) 		
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 315, "Corredor"	, 100, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 433, "Vendedor"	, 100, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, pY + 5, iniX - 553, "Producto"	, 100, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, pY	  , iniX - 645, "Kilos"		, 100, PDF_ALIGN_LEFT)		
	Call GF_writeVerticalText(oPDF, pY+10 , iniX - 620, "Bruto"		, 100, PDF_ALIGN_LEFT)		
	Call GF_writeVerticalText(oPDF, pY+10 , iniX - 665, "Tara"		, 100, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, pY	  , iniX - 715, "Merma"		, 100, PDF_ALIGN_LEFT)		
	Call GF_writeVerticalText(oPDF, pY+10 , iniX - 705, "Cal."		, 100, PDF_ALIGN_LEFT)		
	Call GF_writeVerticalText(oPDF, pY+10 , iniX - 740, "Vol."		, 100, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, pY	  , iniX - 780, "Kilos"		, 100, PDF_ALIGN_LEFT)		
	Call GF_writeVerticalText(oPDF, pY+10 , iniX - 780, "Netos"		, 100, PDF_ALIGN_LEFT)
	drawTableTitle = pY + 25 
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawTableBody(pY, pRs, pTipoTransporte)
	Dim iniX, gap
		
	Call GF_setFont(oPDF,"COURIER", 10, FONT_STYLE_NORMAL)	
	inix = 830	
	while ((not pRs.eof) and (pY < PAGE_LAST_LINE))
		neto = CLng(pRs("Bruto")) - CLng(pRs("Tara")) - CLng(pRs("Merma")) - CLng(pRs("MERMAVOLATIL"))
		Call GF_writeVerticalText(oPDF, pY, iniX		, pRs("DTCONTABLE"), 100, PDF_ALIGN_LEFT) 	
		Call GF_writeVerticalText(oPDF, pY, iniX - 65	, GF_EDIT_CBTE(pRs("NUCARTAPORTE")), 100, PDF_ALIGN_LEFT) 	
		Call GF_writeVerticalText(oPDF, pY, iniX - 150	, pRs("CDVAGON"), 100, PDF_ALIGN_LEFT)
		Call GF_writeVerticalText(oPDF, pY, iniX - 195, Left(Trim(pRs("DSCLIENTE")), 18), 100, PDF_ALIGN_LEFT) 		
		Call GF_writeVerticalText(oPDF, pY, iniX - 315, Left(Trim(pRs("DSCORREDOR")), 18), 100, PDF_ALIGN_LEFT)	
		Call GF_writeVerticalText(oPDF, pY, iniX - 433, Left(Trim(pRs("DSVENDEDOR")), 18), 100, PDF_ALIGN_LEFT)	
		Call GF_writeVerticalText(oPDF, pY, iniX - 553, Left(Trim(pRs("DSPRODUCTO")), 9), 100, PDF_ALIGN_LEFT)			
		Call GF_writeVerticalText(oPDF, pY, iniX - 603, GF_EDIT_DECIMALS(pRs("Bruto"), 0)	, 50, PDF_ALIGN_RIGHT)			
		Call GF_writeVerticalText(oPDF, pY, iniX - 647, GF_EDIT_DECIMALS(pRs("Tara"), 0), 50, PDF_ALIGN_RIGHT)			
		Call GF_writeVerticalText(oPDF, pY, iniX - 683, GF_EDIT_DECIMALS(pRs("Merma"), 0), 50, PDF_ALIGN_RIGHT)			
		Call GF_writeVerticalText(oPDF, pY, iniX - 715, GF_EDIT_DECIMALS(pRs("MERMAVOLATIL"), 0), 50, PDF_ALIGN_RIGHT)			
		Call GF_writeVerticalText(oPDF, pY, iniX - 763, GF_EDIT_DECIMALS(neto, 0), 50, PDF_ALIGN_RIGHT)			
		totalDescargas(pTipoTransporte) = CLng(totalDescargas(pTipoTransporte)) + 1
		totalBruto(pTipoTransporte) = CLng(totalBruto(pTipoTransporte)) + CLng(pRs("Bruto"))
		totalTara(pTipoTransporte) = CLng(totalTara(pTipoTransporte)) + CLng(pRs("Tara"))
		totalMermaCal(pTipoTransporte) = CLng(totalMermaCal(pTipoTransporte)) + CLng(pRs("Merma"))
		totalMermaVol(pTipoTransporte) = CLng(totalMermaVol(pTipoTransporte)) + CLng(pRs("MERMAVOLATIL"))
		totalNeto(pTipoTransporte) = Clng(totalNeto(pTipoTransporte)) + CLng(neto)
		pY = pY + 10
		pRs.MoveNext()
	wend
	drawTableBody = pY
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawPageFooter(pPgNbr)
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)
	Call GF_writeVerticalText(oPDF, 580, 130, "Pagina " & pPgNbr, 100, PDF_ALIGN_RIGHT)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawDataFilters(pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)
	Dim aux		
	
	Call GF_setFont(oPDF,"COURIER", 10, FONT_STYLE_NORMAL)			
	Call GF_writeVerticalText(oPDF,  90, 700, "Desde     :" , 75, PDF_ALIGN_LEFT)	
	Call GF_writeVerticalText(oPDF, 100, 700, "Hasta     :" , 75, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, 110, 700, "Producto  :" , 75, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, 120, 700, "Transporte:" , 75, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF,  90, 500, "Cliente :" 	, 50, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, 100, 500, "Corredor:" 	, 50, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, 110, 500, "Vendedor:" 	, 50, PDF_ALIGN_LEFT)
		
	Call GF_setFont(oPDF,"COURIER", 10, FONT_STYLE_BOLD)	
	Call GF_writeVerticalText(oPDF,  90, 625, pDtDesde, 50, PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF, 100, 625, pDtHasta, 50, PDF_ALIGN_LEFT)	
	if (pCdProducto <> 0) then aux = getDSProducto(pCdProducto) else aux = "Todos" end if
	Call GF_writeVerticalText(oPDF, 110, 625, aux, 50, PDF_ALIGN_LEFT)	
	if (pTipoTransporte = TIPO_TRANSPORTE_CAMION) then aux = "CAMIONES" else if (pTipoTransporte = TIPO_TRANSPORTE_VAGON) then aux = "VAGONES" else aux = "Todos" end if
	Call GF_writeVerticalText(oPDF, 120, 625, aux, 50, PDF_ALIGN_LEFT)
	if (pCdCliente <> 0) then aux = getDSCliente(pCdCliente) else aux = "Todos" end if
	Call GF_writeVerticalText(oPDF,  90, 440, aux, 50, PDF_ALIGN_LEFT)	
	if (pCdCorredor <> 0) then aux = getDSCorredor(pCdCorredor) else aux = "Todos" end if
	Call GF_writeVerticalText(oPDF, 100, 440, aux, 50, PDF_ALIGN_LEFT)	
	if (pCdVendedor <> 0) then aux = getDSVendedor(pCdVendedor) else aux = "Todos" end if
	Call GF_writeVerticalText(oPDF, 110, 440, aux, 50, PDF_ALIGN_LEFT)	
	
End Function
'------------------------------------------------------------------------------------------------------------------------
Function addNewPage(pNroPagina)		
	Dim coordY
	coordY = 110
	Call GF_newPage(oPDF)
	Call PDFGirarHoja(90)			
	Call drawPageHeader()
	coordY = drawTableTitle(coordY)
	Call drawPageFooter(pNroPagina)
	addNewPage = coordY
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawCuadroResumen()
	Dim Xo, Yo, i, mytotalBruto, mytotalTara, mytotalMermaCal, mytotalMermaVol, mytotalNeto
		
	mytotalBruto = 0
	mytotalTara = 0
	mytotalMermaCal = 0
	mytotalMermaVol = 0
	mytotalNeto = 0
		
	'Recuadro
	Xo = 160
	Yo = 170
	Call GF_squareBox(oPDF, Yo, Xo, 140, Xo + 395, 0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)	
	Call GF_verticalLine(oPDF, Yo + 33, Xo, Xo + 395)	
	Call GF_verticalLine(oPDF, Yo + 60, Xo, Xo + 395)
	Call GF_verticalLine(oPDF, Yo + 110, Xo, Xo + 395)	
	'Titulos
	Call GF_setFont(oPDF,"COURIER", 16, FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF, Yo + 8, Xo + 565, "RESUMEN"	, 600, PDF_ALIGN_CENTER)			
	Call GF_setFont(oPDF,"COURIER", 12, FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF, Yo + 40, Xo + 565, "Transporte"	, 120, PDF_ALIGN_CENTER)		
	Call GF_writeVerticalText(oPDF, Yo + 40, Xo + 445, "Cant."		,  40, PDF_ALIGN_CENTER)		
	Call GF_writeVerticalText(oPDF, Yo + 35, Xo + 415, "Kilos"		,  90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 45, Xo + 415, "Bruto"		,  90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 35, Xo + 315, "Kilos"		,  90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 45, Xo + 315, "Tara"		,  90, PDF_ALIGN_RIGHT)	
	Call GF_writeVerticalText(oPDF, Yo + 35, Xo + 225, "Merma"		,  60, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 45, Xo + 225, "Cal."		,  60, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 35, Xo + 165, "Merma"		,  60, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 45, Xo + 165, "Vol."		,  60, PDF_ALIGN_RIGHT)	
	Call GF_writeVerticalText(oPDF, Yo + 35, Xo + 105, "Kilos"		,  90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 45, Xo + 105, "Netos"		,  90, PDF_ALIGN_RIGHT)
	'Datos
	Call GF_setFont(oPDF,"COURIER", 12, FONT_STYLE_NORMAL)
	Call GF_writeVerticalText(oPDF, Yo +  70, Xo + 565, "Camiones"	, 120, PDF_ALIGN_CENTER)			
	Call GF_writeVerticalText(oPDF, Yo +  90, Xo + 565, "Vagones"	, 120, PDF_ALIGN_CENTER)		
	For i = LBound(totalDescargas)+1 to UBound(totalDescargas)
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 445, GF_EDIT_DECIMALS(totalDescargas(i), 0), 40, PDF_ALIGN_CENTER)				
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 415, GF_EDIT_DECIMALS(totalBruto(i), 0)	, 90, PDF_ALIGN_RIGHT)		
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 315, GF_EDIT_DECIMALS(totalTara(i), 0)	 	, 90, PDF_ALIGN_RIGHT)		
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 225, GF_EDIT_DECIMALS(totalMermaCal(i), 0) , 60, PDF_ALIGN_RIGHT)		
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 165, GF_EDIT_DECIMALS(totalMermaVol(i), 0) , 60, PDF_ALIGN_RIGHT)	
		Call GF_writeVerticalText(oPDF, Yo + 60 + 15*i, Xo + 105, GF_EDIT_DECIMALS(totalNeto(i), 0)	  	, 90, PDF_ALIGN_RIGHT)		
		mytotalBruto = mytotalBruto + totalBruto(i) 
		mytotalTara = mytotalTara + totalTara(i)
		mytotalMermaCal = mytotalMermaCal + totalMermaCal(i)
		mytotalMermaVol = mytotalMermaVol + totalMermaVol(i)
		mytotalNeto = mytotalNeto + totalNeto(i)
	next		
	Call GF_setFont(oPDF,"COURIER", 12, FONT_STYLE_BOLD)
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 565, "TOTAL"	, 120, PDF_ALIGN_CENTER)			
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 415, GF_EDIT_DECIMALS(mytotalBruto, 0)		, 90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 315, GF_EDIT_DECIMALS(mytotalTara, 0)	 	, 90, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 225, GF_EDIT_DECIMALS(mytotalMermaCal, 0) 	, 60, PDF_ALIGN_RIGHT)		
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 165, GF_EDIT_DECIMALS(mytotalMermaVol, 0) 	, 60, PDF_ALIGN_RIGHT)	
	Call GF_writeVerticalText(oPDF, Yo + 118, Xo + 105, GF_EDIT_DECIMALS(mytotalNeto, 0)	  	, 90, PDF_ALIGN_RIGHT)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawResumen(pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)
	Call setWorkPage(oPDF, 1)
	Call PDFGirarHoja(90)
	Call drawDataFilters(pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)
	Call drawPageHeader()
	Call drawCuadroResumen()
	Call drawPageFooter(1)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawReporte(rs, ByRef pIniPagina, ByRef pY, pTipoTransporte)
	Dim coordY
	
	coordY = pY
	while (not rs.eof)
		if (coordY >= PAGE_LAST_LINE) then 
			pIniPagina = pIniPagina + 1
			coordY = addNewPage(pIniPagina)		
		end if
		coordY = drawTableBody(coordY, rs, pTipoTransporte)		
	wend
	pY = coordY
	
end function
'------------------------------------------------------------------------------------------------------------------------
Function armarPDF(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)
	Dim rs, nroPagina, nroLinea, i
	
	Set oPDF = GF_createPDF("PDFTemp.pdf")
	Call GF_setPDFMODE(PDF_STREAM_MODE)	
	
	for i = LBound(totalDescargas) to UBound(totalDescargas)
		totalDescargas(i) = 0
		totalBruto(i) = 0
		totalTara(i) = 0
		totalMermaCal(i) = 0
		totalMermaVol(i) = 0
		totalNeto(i) = 0
	Next	
	'/******************************************************************************************************************************\ 
	'|* IMPORTANTE: SE ARMA PRIMERO EL DETALLE DESDE LA HOJA 2 (PARA TOTALIZAR) Y AL FINAL SE ARMA LA HOJA 1 CON EL CUADRO RESUMEN *|
	'\******************************************************************************************************************************/ 
	
	nroPagina = 1
	nroLinea = PAGE_LAST_LINE	
	if ((pTipoTransporte = TIPO_TRANSPORTE_CAMION) or ((pTipoTransporte = TIPO_TRANSPORTE_CAMVAG))) then		
		Set rs = getCamionesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
		Call drawReporte(rs, nroPagina, nroLinea, TIPO_TRANSPORTE_CAMION)
	end if
	if ((pTipoTransporte = TIPO_TRANSPORTE_VAGON) or ((pTipoTransporte = TIPO_TRANSPORTE_CAMVAG))) then
		Set rs = getVagonesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
		Call drawReporte(rs, nroPagina, nroLinea, TIPO_TRANSPORTE_VAGON)
		Set rs = getHVagonesDescargados(pPto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor)
		Call drawReporte(rs, nroPagina, nroLinea, TIPO_TRANSPORTE_VAGON)
	end if
	Call GF_verticalLine(oPDF, nroLinea + 5, 10, 825)	
	
	Call drawResumen(pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)
	
	Call GF_closePDF(oPDF)
	
End Function

'********************************************************************************
'******* 					COMIENZO DE LA PAGNA 						  *******
'********************************************************************************
Dim pto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, oPDF, pTipoTransporte
Dim totalDescargas(2), totalBruto(2), totalTara(2), totalMermaCal(2), totalMermaVol(2), totalNeto(2)


Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
pDtDesde = GF_PARAMETROS7("fd", "", 6)
if (pDtDesde = "") then pDtDesde = GF_FN2DTCONTABLE(Left(session("MmtoSistema"), 8))
pDtHasta = GF_PARAMETROS7("fh", "", 6)
if (pDtHasta = "") then pDtHasta = pDtDesde
pCdProducto = GF_PARAMETROS7("prod", 0, 6)
pCdVendedor = GF_PARAMETROS7("ven", 0, 6)
pCdCorredor = GF_PARAMETROS7("cor", 0, 6)
pCdCliente = GF_PARAMETROS7("cl", 0, 6)
pTipoTransporte = GF_PARAMETROS7("tt", 0, 6)
if (pTipoTransporte = 0) then pTipoTransporte = TIPO_TRANSPORTE_CAMVAG

Call armarPDF(pto, pDtDesde, pDtHasta, pCdProducto, pCdCliente, pCdCorredor, pCdVendedor, pTipoTransporte)

%>