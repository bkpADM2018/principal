<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%

Const PIC_TEXTOFIRMAS = 0
Const PIC_FIRMAS = 1
Const PIC_MAXREGISTROS = 40
Const PIC_MAXREGISTROS__AUTHORIZATED = 29
Const PICAJU_MAXREGISTROS__AUTHORIZATED = 8
Const POSITION_Y = 170
Const POSITION_Y_AJU = 95
Const P_Y_FIRMAS = 600
Const P_Y_FIRMAS_AUTHORIZATED = 485
Const P_Y_OBSERVACIONES = 706

dim PIC_idCtzElegida, PIC_idPedido, PIC_cdPedido, PIC_IdDivision, PIC_Obra, PIC_momento, PIC_usuario, PIC_idProveedor
dim PIC_Proveedor, PIC_fecEntrega, PIC_observaciones, PIC_firmante1Ds, PIC_firmante2Ds, PIC_firmante3Ds, PIC_firmante1Tx, PIC_firmante2Tx, PIC_firmante3Tx, PIC_firmante1Cd, PIC_firmante2Cd, PIC_firmante3Cd
dim IT_artID, IT_artDS, IT_cantidad, IT_unidadDS, IT_PartPresup, PIC_estado, PIC_TipoCambio, PIC_Moneda
dim rsDET, connDET, rs, conn, oPDF, totalImporte, firmaSolicitante, firmaResponsable, firmaSuperisor
dim PIC_firmante4Cd, PIC_firmante4Ds, PIC_firmante4Tx,PIC_idContrato, PIC_importePesos, PIC_importeDolares
dim PIC_firmante5Cd, PIC_firmante5Ds, PIC_firmante5Tx,PIC_idObra,PIC_cdAfe, regCargados, g_BgColor
'-----------------------------------------------------------------------------------------------
Function errorAcceso() 
	response.redirect "comprasAccesoDenegado.asp"
End Function
'-----------------------------------------------------------------------------------------------
Function get_DatosFirmas(idCotizacion, byref firmante1Cd, byref firmante1Ds, byref firmante1Tx, byref firmante2cd, byref firmante2Ds, byref firmante2Tx, byref firmante3Cd, byref firmante3Ds, byref firmante3Tx, byref firmante4Cd, byref firmante4Ds, Byref firmante4Tx, ByRef firmante5Cd, ByRef firmante5Ds, ByRef firmante5Tx) 
	Dim strSQL, rsFirmas
	Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLCTZFIRMAS_GET_BY_IDCOTIZACION", idCotizacion)
	while not rsFirmas.eof
			if (firmante1Cd ="") then
				firmante1Cd = rsFirmas("CDUSUARIO")
				firmante1Ds = getUserDescription(rsFirmas("CDUSRROL"))
				firmante1Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante2Cd ="") then
				firmante2Cd = rsFirmas("CDUSUARIO")
				firmante2Ds = getUserDescription(rsFirmas("CDUSRROL"))
				firmante2Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante3Cd ="") then			
				firmante3Cd = rsFirmas("CDUSUARIO")
				firmante3Ds = getUserDescription(rsFirmas("CDUSRROL"))
				firmante3Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))					
			elseif (firmante4Cd ="") then
				firmante4Cd = rsFirmas("CDUSUARIO")
				firmante4Ds = getUserDescription(rsFirmas("CDUSRROL"))
				firmante4Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante5Cd ="") then
				firmante5Cd = rsFirmas("CDUSUARIO")
				firmante5Ds = getUserDescription(rsFirmas("CDUSRROL"))
				firmante5Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			end if				
		rsFirmas.movenext
	wend	
End Function
'-----------------------------------------------------------------------------------------------
Function isPICAuthorizated()
	isPICAuthorizated = false
	if ((PIC_firmante4Cd <> "")or(PIC_firmante5Cd <> "")) then isPICAuthorizated = true
End Function
'-----------------------------------------------------------------------------------------------
Function dibujar_encabezado(pIdContrato)
	Dim titulo
	titulo = "PEDIDO INTERNO DE COMPRA"
	if (CLng(pIdContrato) <> 0) then titulo = "CBTE. ELECTRONICO DE CUMPLIMIENTO"	
	'dibuja recuadro general
	Call GF_squareBox(oPDF, 2, 2, 590, 848, 0, "", "#000000", 2, PDF_SQUARE_ROUND)
	'logo y titulo
	Call GF_writeImage(oPDF, Server.MapPath("images\ADMlogo2.jpg"), 10, 10, 60, 55, 0)
	call GF_setFont(oPDF,"ARIAL",20,0)
	Call GF_writeTextAlign(oPDF,10, 25, GF_TRADUCIR(titulo), 570, PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"ARIAL",14,0)
	Call GF_writeTextAlign(oPDF,300, 50, GF_TRADUCIR("Nro : " & PIC_idCtzElegida), 280, PDF_ALIGN_RIGHT)
end Function
'-----------------------------------------------------------------------------------------------
Function infoPrincipalBox(p_idContrato)	
	Dim p_Y
	p_Y = 130	
	'dibuja celdas
	Call GF_squareBox(oPDF, 10, 70, 570, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	'dibuja partida
	Call GF_squareBox(oPDF, 10, 85, 110, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120, 85, 460, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'dibuja pedido
	Call GF_squareBox(oPDF, 10, 100, 110, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120, 100, 220, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'dibuja division
	Call GF_squareBox(oPDF, 340, 100, 80, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 420, 100, 160, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'dib proveedor	
	Call GF_squareBox(oPDF, 10, 115, 110, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120, 115, 460, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'dibuja  fecha
	'Call GF_squareBox(oPDF, 340, 115, 80, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	'Call GF_squareBox(oPDF, 420, 115, 160, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'comprueba si tiene contrato dibuja la celda
	if (PIC_cdAfe <> "")or(p_idContrato > 0)then        
        Call GF_squareBox(oPDF, 10, p_Y, 110, 15, 0,g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_squareBox(oPDF, 120, p_Y, 230, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
    	Call GF_squareBox(oPDF, 340, p_Y, 80, 15, 0,g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_squareBox(oPDF, 420, p_Y, 160, 15, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
        p_Y = p_Y + 15
    end if
	Call GF_squareBox(oPDF, 10, p_Y, 570, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, p_Y + 15, 570, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_verticalLine(oPDF, 50, p_Y + 15, 15)
	Call GF_verticalLine(oPDF, 345, p_Y + 15, 15)
	Call GF_verticalLine(oPDF, 405, p_Y + 15, 15)	
	Call GF_verticalLine(oPDF, 475, p_Y + 15, 15)
end Function
'-----------------------------------------------------------------------------------------------
Function armarInfoPrincipal(p_partida, p_pedido, p_Proveedor, p_division, p_idContrato)
	dim tituloInt
	Call infoPrincipalBox(p_idContrato)
	'titulos
	Dim py
	py = 132
	call GF_setFont(oPDF,"ARIAL",10,8)
	tituloInt = "Informacion General"
	Call GF_setFontColor("#FFFFFF")
	Call GF_writeTextAlign(oPDF,15, 72, GF_TRADUCIR(tituloInt), 565, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,15, 87, GF_TRADUCIR("Ptda. Presupuestaria"), 100, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,15, 102, GF_TRADUCIR("Pedido"), 100, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,345, 102, GF_TRADUCIR("Divisi�n"), 75, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,15, 117, GF_TRADUCIR("Proveedor"), 100, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,345, 132, GF_TRADUCIR("Contrato"), 75, PDF_ALIGN_LEFT)	    
	Call GF_writeTextAlign(oPDF,15, py, GF_TRADUCIR("AFE"), 100, PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(oPDF,345, py, GF_TRADUCIR("Contrato"), 100, PDF_ALIGN_LEFT)
	py = py + 15	
	Call GF_writeTextAlign(oPDF,15, py, GF_TRADUCIR("Detalle"), 565, PDF_ALIGN_CENTER)
	py = py + 15
	Call GF_writeTextAlign(oPDF,10, py, GF_TRADUCIR("Codigo"), 40, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,55, py, GF_TRADUCIR("Descripci�n"), 295, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,350, py, GF_TRADUCIR("Cantidad"), 50, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,410, py , GF_TRADUCIR("Ptda. Presup."), 60, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,480, py , GF_TRADUCIR("Importe s/IVA"), 90, PDF_ALIGN_CENTER)
	Call GF_setFontColor("#000000")
	'informaci�n
	call GF_setFont(oPDF,"ARIAL",10,0)
	Call GF_writeTextAlign(oPDF,125, 87, p_partida, 550, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,125, 102, p_pedido, 260, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,125, 117, p_Proveedor, 550, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,425, 102, p_division, 100, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,425, 132, getCodigoCTC(p_idContrato), 100, PDF_ALIGN_LEFT)	
    Call GF_writeTextAlign(oPDF,125, 132, PIC_cdAfe , 100, PDF_ALIGN_LEFT)
    'Call GF_writeTextAlign(oPDF,425, 132, getCodigoCTC(p_idContrato), 100, PDF_ALIGN_LEFT)    	
	armarInfoPrincipal = py
end Function
'-----------------------------------------------------------------------------------------------
Function dibujarBoxFirmas(p_y)

	Dim limiteCD, limiteSP, importeCompra, unidadCD, unidadSP, tituloFirma
	
	Call GF_squareBox(oPDF, 10, p_y, 570, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 10, p_y + 15, 570, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 10, p_y + 115, 570, 115, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_setFontColor("#FFFFFF")
	Call GF_writeTextAlign(oPDF,15, p_y + 2, GF_TRADUCIR("Observaciones :"), 190, PDF_ALIGN_LEFT)
	Call GF_setFontColor("#000000")
	if p_y = P_Y_FIRMAS_AUTHORIZATED then
	    Call GF_squareBox(oPDF, 10, p_y + 230, 285, 115, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)			
	    Call GF_squareBox(oPDF, 295, p_y + 230, 285, 115, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)		
	end if
end Function
'-----------------------------------------------------------------------------------------------
Function dibujarFirmas(position_y, firmante1Cd, firmante1Ds, firmante1Tx, firmante2cd, firmante2Ds, firmante2Tx, firmante3Cd, firmante3Ds, firmante3Tx, firmante4Cd, firmante4Ds, firmante4Tx, firmante5Cd, firmante5Ds, firmante5Tx ,p_observaciones)
	Dim firma,strObservaciones1,strObservaciones,py, vNoExcedido,vExcedido,i ,j
	
	if (firmante1Tx <> "") then
		firma = obtenerFirma(firmante1Cd)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & firma), 10, position_y + 131, 190, 75, 0)		
	end if
   	if (firmante2Tx <> "") then
   		firma = obtenerFirma(firmante2Cd)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & firma), 200, position_y + 131, 190, 75, 0)
	end if
   	if (firmante3Tx <> "") then
   		firma = obtenerFirma(firmante3Cd)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & firma), 390, position_y + 131, 190, 75, 0)
	end if
	'Se dibujan las firmas especiales, solo la/s que figuren.
	if (firmante4Tx <> "") then			
		firma = obtenerFirma(firmante4Cd)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & firma), 50, position_y + 246, 200, 75, 0)					
	end if
	if (firmante5Tx <> "") then
		'Se toman los datos de la firma del Director			
		firma = obtenerFirma(firmante5Cd)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & firma), 325, position_y + 246, 200, 75, 0)		
	end if							
	'datos
	call GF_setFont(oPDF,"ARIAL",8,0)
	i = 0	
	j = 0
	if (p_observaciones <> "")	then
		if(InStr(p_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO) > 0)then
			strObservaciones = split(p_observaciones,PIC_TEXTO_DETALLE_PRESUPUESTO)			
			py = GF_writeTextPlus(oPDF, 15, position_y + 18, strObservaciones(0), 570, 8, PDF_ALIGN_LEFT)			
			vNoExcedido = split(strObservaciones(1),ENTER_SYMBOL)			
			Do Until (i > Ubound(vNoExcedido) - 1)
				if(py > P_Y_OBSERVACIONES)Then Exit Do
				py = GF_writeTextPlus(oPDF, 15, py, vNoExcedido(i), 570, 8, PDF_ALIGN_LEFT)
				i = i + 1
			Loop
			call GF_setFontColor("#ff0000")
			call GF_setFont(oPDF,"ARIAL",8,8)
			vExcedido = split(strObservaciones(2),ENTER_SYMBOL)
			Do Until (j > Ubound(vExcedido) - 1)
				if(py > P_Y_OBSERVACIONES)Then Exit Do					
				py = GF_writeTextPlus(oPDF, 15, py, vExcedido(j), 570, 8, PDF_ALIGN_LEFT)
				j = j + 1				
			Loop
			call GF_setFont(oPDF,"ARIAL",8,0)
			call GF_setFontColor("000000")						
		else			
			Call GF_writeTextPlus(oPDF, 15, position_y + 18, editText4DB(p_observaciones), 570, 8, PDF_ALIGN_LEFT)
		end if	
	end if		
	call GF_setFont(oPDF,"ARIAL",10,0)
	if (firmante1Ds <> "")	then	Call GF_writeTextAlign(oPDF, 10, position_y + 207, firmante1Ds, 190, PDF_ALIGN_CENTER)
	if (firmante2Ds <> "")	then	Call GF_writeTextAlign(oPDF,200, position_y + 207, firmante2Ds, 190, PDF_ALIGN_CENTER)
	if (firmante3Ds <> "")	then	Call GF_writeTextAlign(oPDF,390, position_y + 207, firmante3Ds, 190, PDF_ALIGN_CENTER)	
	if (firmante4Ds <> "")	then	Call GF_writeTextAlign(oPDF, 10, position_y + 321, firmante4Ds, 285, PDF_ALIGN_CENTER)
	if (firmante5Ds <> "")	then	Call GF_writeTextAlign(oPDF,285, position_y + 321, firmante5Ds, 285, PDF_ALIGN_CENTER)				
	call GF_setFont(oPDF,"ARIAL",6,0)
	if (firmante1Tx <> "")	then	Call GF_writeTextAlign(oPDF, 10, position_y + 220, firmante1Tx, 190, PDF_ALIGN_CENTER)
	if (firmante2Tx <> "")	then	Call GF_writeTextAlign(oPDF,200, position_y + 220, firmante2Tx, 190, PDF_ALIGN_CENTER)
	if (firmante3Tx <> "")	then	Call GF_writeTextAlign(oPDF,390, position_y + 220, firmante3Tx, 190, PDF_ALIGN_CENTER)
	if (firmante4Tx <> "")	then	Call GF_writeTextAlign(oPDF, 10, position_y + 336, firmante4Tx, 285, PDF_ALIGN_CENTER)
	if (firmante5Tx <> "")	then	Call GF_writeTextAlign(oPDF,285, position_y + 336, firmante5Tx, 285, PDF_ALIGN_CENTER)			
end Function
'-----------------------------------------------------------------------------------------------
Function finPagina(p_modo)
	if (p_modo = PIC_TEXTOFIRMAS) then
		Call dibujarBoxFirmas(P_Y_FIRMAS)
		call GF_setFont(oPDF,"ARIAL",14,0)
		Call GF_writeTextAlign(oPDF,200, 750, GF_TRADUCIR("La firma del pedido"), 190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,200, 770, GF_TRADUCIR("se realiza unicamente"), 190, PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,200, 790, GF_TRADUCIR("en la ultima pagina"), 190, PDF_ALIGN_CENTER)
	else
		if isPICAuthorizated() then
			Call dibujarBoxFirmas(P_Y_FIRMAS_AUTHORIZATED)			
			Call GF_verticalLine(oPDF, 200, 600, 115)
			Call GF_verticalLine(oPDF, 390, 600, 115)
			Call dibujarFirmas(P_Y_FIRMAS_AUTHORIZATED, PIC_firmante1Cd, PIC_firmante1Ds, PIC_firmante1Tx, PIC_firmante2Cd, PIC_firmante2Ds, PIC_firmante2Tx, PIC_firmante3Cd, PIC_firmante3Ds, PIC_firmante3Tx, PIC_firmante4Cd, PIC_firmante4Ds, PIC_firmante4Tx, PIC_firmante5Cd, PIC_firmante5Ds, PIC_firmante5Tx, PIC_observaciones)
		else
			Call dibujarBoxFirmas(P_Y_FIRMAS)
			Call GF_verticalLine(oPDF, 200, 715, 115)
			Call GF_verticalLine(oPDF, 390, 715, 115)
			Call dibujarFirmas(P_Y_FIRMAS, PIC_firmante1Cd, PIC_firmante1Ds, PIC_firmante1Tx, PIC_firmante2Cd, PIC_firmante2Ds, PIC_firmante2Tx, PIC_firmante3Cd, PIC_firmante3Ds, PIC_firmante3Tx, " ", " ", " "," ", " ", " ", PIC_observaciones)
		end if
	end if
	
	
	Call GF_setFontColor("#FF0000")
	if (PIC_estado = CTZ_ANULADA) then
		call GF_setFont(oPDF,"ARIAL",100,8)
		'Call GF_writeTextAlign(oPDF, 20, 280,GF_TRADUCIR("ANULADO"),570, PDF_ALIGN_CENTER)
		call GF_writeText(oPDF, 40, 380, GF_TRADUCIR("ANULADO"), 25)
		'y_aux = 635
	end if
	Call GF_setFontColor("#000000")	
	
	
	
end Function
'-----------------------------------------------------------------------------------------------
Function escribeRegistro(p_y, p_codigo, p_descripcion, p_cantidad, p_abrrunidad, p_ptdapresup, p_importe)
	call GF_setFont(oPDF,"ARIAL",8,0)
	if (p_codigo <> "")     then    Call GF_writeTextAlign(oPDF, 10, p_y, p_codigo,       40, PDF_ALIGN_CENTER)
	if (p_descripcion <> "") then   Call GF_writeTextAlign(oPDF, 55, p_y, p_descripcion, 295, PDF_ALIGN_LEFT)
	if (p_cantidad <> "")   then    Call GF_writeTextAlign(oPDF,350, p_y, p_cantidad & " " & p_abrrunidad, 50, PDF_ALIGN_RIGHT)
	if (p_ptdapresup <> "") then    Call GF_writeTextAlign(oPDF,410, p_y, p_ptdapresup,   60, PDF_ALIGN_CENTER)
	if (p_importe <> "")   then     
	    Call GF_writeTextAlign(oPDF,480, p_y, getSimboloMoneda(PIC_Moneda),      10, PDF_ALIGN_LEFT)
	    Call GF_writeTextAlign(oPDF,495, p_y, p_importe,      80, PDF_ALIGN_RIGHT)
    end if	    
end Function
'-----------------------------------------------------------------------------------------------
Function findetalle(p_TotalImporte)		
	Dim baseY
	
	call GF_setFont(oPDF,"ARIAL",10,0)
	if isPICAuthorizated() then
		baseY=471
	else
		baseY=587
	end if
	
	if (PIC_Moneda = MONEDA_DOLAR) then Call GF_writeTextAlign(oPDF,15, baseY, "Tipo Cambio: " & GF_EDIT_DECIMALS(cdbl(PIC_TipoCambio)*1000,3), 350, PDF_ALIGN_LEFT)		
	Call GF_squareBox(oPDF, 372, baseY-1, 208, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_verticalLine(oPDF, 450, baseY-1, 15)
	call GF_setFont(oPDF,"ARIAL",10,8)		
	Call GF_writeTextAlign(oPDF,372, baseY+1, GF_TRADUCIR("TOTAL"), 75, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,455, baseY+1, getSimboloMoneda(PIC_Moneda), 15, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,475, baseY+1, p_TotalImporte, 100, PDF_ALIGN_RIGHT)	
	call GF_setFont(oPDF,"ARIAL",10,0)
		
end Function
'-----------------------------------------------------------------------------------------------
'Imrpime la leyenda de limites excedidos para las compras directas.
Function printMoneyLimitWarning(ByRef y_posicion, ByRef nroPagina, ByRef regCargados)
    Dim cantPICDirectos30, montoPICDirectos30, cantPICDirectos365, montoPICDirectos365, mmtoDesde,limiteCD
    
    '---Si corresponde se imprimen las lineas especiales sobre las compras directas realizadas al proveedor.	
	if ((PIC_idPedido = 0) and (PIC_idContrato = 0)) then
        'Es compra directa, si la cantidad del �ltimo mes supera el l�mite maximo de una compra directa se muestra la info.
        mmtoDesde = GF_DTEADD(PIC_Momento,-30,"D")
        Call totalizarComprasDirectasProveedor(PIC_idProveedor, PIC_idDivision, mmtoDesde, PIC_Momento, MONEDAL_DOLAR, cantPICDirectos30, montoPICDirectos30)        
        
        mmtoDesde = GF_DTEADD(PIC_Momento,-365,"D")
        Call totalizarComprasDirectasProveedor(PIC_idProveedor, PIC_idDivision, mmtoDesde, PIC_Momento, MONEDAL_DOLAR, cantPICDirectos365, montoPICDirectos365)            
        
        
        Call GF_setFont(oPDF,"ARIAL",10,8)                                    
        if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
        Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("TOTAL Compras Directas realizadas en los �ltimos 30 d�as (" & cantPICDirectos30 - 1 & " Pics): "), 50, PDF_ALIGN_RIGHT)            
        Call GF_writeTextAlign(oPDF,425,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(montoPICDirectos30 - PIC_importeDolares, 2), 150, PDF_ALIGN_RIGHT)                        
        regCargados = regCargados + 1
        y_posicion = y_posicion + 10
        if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
        Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("TOTAL Compras Directas realizadas en los �ltimos 12 meses (" & cantPICDirectos365 - 1 & " Pics):"), 50, PDF_ALIGN_RIGHT)                              
        Call GF_writeTextAlign(oPDF,425,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(montoPICDirectos365 - PIC_importeDolares, 2), 150, PDF_ALIGN_RIGHT)            
        regCargados = regCargados + 1
        y_posicion = y_posicion + 15        
    end if
    '---
End Function
'-----------------------------------------------------------------------------------------------
' Autor: 	
'           Ajaya Nahuel
' Fecha: 	
'           04/12/2014
' Objetivo:	
'			Imprime los distintos saldos que presenta el Pic:
'               -> Saldo de la Partida, Total de la Partida (si el Pic posee una Obra)
'               -> Saldo del Contrato, Total del Contrato (si el Pic posee un Contrato)
'               -> Saldo del Pic (se imprime siempre)
' Devuelve:
'			-
'-----------------------------------------------------------------------------------------------
Function printBalanceOfPIC(ByRef y_posicion, ByRef nroPagina, ByRef regCargados)
    Dim strTotalObra,strSaldoObra,auxSaldo,strTotalCTC,strSaldoCTC,rsSel    
    'Si hay partidas muestro el Total de la Partida y el Saldo de la Partida
    Call GF_setFont(oPDF,"ARIAL",10,8)
    if ((Cdbl(PIC_idObra) <> 0) and (Cdbl(PIC_idObra) <> OBRA_GEID)) then
        Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsSalOBR, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", PIC_idObra & "||0||0" )
        if not rsSalOBR.Eof then
            if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
            Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("TOTAL Partida Presupuestaria"), 50, PDF_ALIGN_RIGHT)
            'Call GF_writeTextAlign(oPDF,420,y_posicion + 5, getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(Cdbl(rsSalOBR("TOTALOBRAPESOS")),2), 75, PDF_ALIGN_RIGHT)
            Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(Cdbl(rsSalOBR("TOTALOBRADOLARES")),2), 175, PDF_ALIGN_RIGHT)
            regCargados = regCargados + 1
            y_posicion = y_posicion + 10
            if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
            Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("SALDO Partida Presupuestaria"), 50, PDF_ALIGN_RIGHT)
            'Call GF_writeTextAlign(oPDF,420,y_posicion + 5, getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(Cdbl(rsSalOBR("TOTALOBRAPESOS")) - (Cdbl(rsSalOBR("IMPORTEPICPESOS")) + Cdbl(rsSalOBR("IMPORTEVALESPESOS"))) ,2), 75, PDF_ALIGN_RIGHT)
            Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(Cdbl(rsSalOBR("TOTALOBRADOLARES")) - (Cdbl(rsSalOBR("IMPORTEPICDOLARES")) + Cdbl(rsSalOBR("IMPORTEVALESDOLARES"))) ,2), 175, PDF_ALIGN_RIGHT)
            regCargados = regCargados + 1
            y_posicion = y_posicion + 15
        end if
    end if
    'Si tiene contrato muestro el Total y Saldo del Contrato
    if (PIC_idContrato > 0) then
        Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsSalCTC, "TBLOBRACONTRATOS_GET_SALDO_BY_IDCONTRATO", PIC_idContrato )
        if not rsSalCTC.Eof then            
            if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
            Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("TOTAL Contrato"), 50, PDF_ALIGN_RIGHT)
            'Call GF_writeTextAlign(oPDF,420,y_posicion + 5, getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(rsSalCTC("TOTALPESOS")),2), 75, PDF_ALIGN_RIGHT)
            Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(rsSalCTC("TOTALDOLARES")),2), 175, PDF_ALIGN_RIGHT)
            regCargados = regCargados + 1
            y_posicion = y_posicion + 10
            if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
            Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("SALDO Contrato"), 50, PDF_ALIGN_RIGHT)
            'Call GF_writeTextAlign(oPDF,420,y_posicion + 5, getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(rsSalCTC("TOTALPESOS")) - Cdbl(rsSalCTC("IMPORTEPESOS")),2), 75, PDF_ALIGN_RIGHT)
            Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(rsSalCTC("TOTALDOLARES")) - Cdbl(rsSalCTC("IMPORTEDOLARES")),2), 175, PDF_ALIGN_RIGHT)
            regCargados = regCargados + 1
            y_posicion = y_posicion + 15
        end if
    end if
    'Siempre muestro el saldo a pagar del PIC        
    saldoPIC = 0
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsSalPIC, "TBLCTZDETALLE_GET_SALDO_A_FACTURAR", "0||" & PIC_idCtzElegida & "||||1")    
    while (not rsSalPIC.Eof)
        saldoPIC = saldoPIC  + CDbl(rsSalPIC("saldo"))
        rsSalPIC.MoveNext()
    wend
    if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)        
    if (PIC_idContrato > 0) then
        Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("SALDO CEC"), 50, PDF_ALIGN_LEFT)            
    else
        Call GF_writeTextAlign(oPDF,100,y_posicion + 5, GF_TRADUCIR("SALDO PIC"), 50, PDF_ALIGN_LEFT)
    end if    
    if (PIC_Moneda = MONEDA_PESO) then
        Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_PESO) &" "& GF_EDIT_DECIMALS(Cdbl(saldoPIC)*100, 2), 175, PDF_ALIGN_RIGHT)
    else
        Call GF_writeTextAlign(oPDF,400,y_posicion + 5, getSimboloMoneda(MONEDA_DOLAR) &" "& GF_EDIT_DECIMALS(Cdbl(saldoPIC)*100, 2), 175, PDF_ALIGN_RIGHT)
    end if                
    regCargados = regCargados + 1
    y_posicion = y_posicion + 15    
End Function
'-----------------------------------------------------------------------------------------------
Function PIC_armadoPDF()
	dim y_posicion, nroPagina	
	y_posicion = POSITION_Y
	regCargados = 0
	nroPagina = 1
	'carga datos
	Call get_datosCotizacion(PIC_idCtzElegida)		
	g_BgColor = "#80A2B7"	
	    
	Call get_DatosFirmas (PIC_idCtzElegida, PIC_firmante1Cd, PIC_firmante1Ds, PIC_firmante1Tx, PIC_firmante2Cd, PIC_firmante2Ds, PIC_firmante2Tx, PIC_firmante3Cd, PIC_firmante3Ds, PIC_firmante3Tx, PIC_firmante4Cd, PIC_firmante4Ds, PIC_firmante4Tx,PIC_firmante5Cd, PIC_firmante5Ds, PIC_firmante5Tx)
	'carga detalles
	if PIC_TIPO = "A" then
		strSQL = "SELECT IDARTICULO, '' AS CANTIDAD, SUM(IMPORTEPESOS) AS IMPORTEPESOS, SUM(IMPORTEDOLARES) AS IMPORTEDOLARES, '' AS IDAREA, '' AS IDDETALLE  FROM TBLCTZAJUSTES WHERE IDCOTIZACION=" & PIC_idCtzElegida & " AND APLICADO='" & TIPO_AFIRMACION & "' GROUP BY IDARTICULO "
	else
		strSQL = "SELECT * from TBLCTZDETALLE where IDCOTIZACION=" & PIC_idCtzElegida & " ORDER BY IDARTICULO"
	end if	
	if((PIC_idContrato > 0)or(PIC_cdAfe <> ""))then y_posicion = y_posicion + 15
	Call executeQueryDB(DBSITE_SQL_INTRA, rsDET, "OPEN", strSQL)
	Call get_PageWrite(nroPagina)	
	While (not rsDET.eof)		    
		While ((not rsDET.eof) and (regCargados < PIC_MAXREGISTROS))
			regCargados = regCargados + 1
			Call get_detallePedido(rsDET, y_posicion)
			y_posicion = y_posicion + 10
			rsDET.MoveNext()
		Wend
		if (not rsDET.eof) then Call addNewPage(y_posicion, nroPagina)
	Wend	
    if (regCargados >= PIC_MAXREGISTROS) then Call addNewPage(y_posicion, nroPagina)
	Call GF_writeTextAlign(oPDF,55,y_posicion, "---------------   Fin del listado de articulos     ---------------", 50, PDF_ALIGN_RIGHT)             			
	regCargados = regCargados + 1
    y_posicion = y_posicion + 10	
    Call printMoneyLimitWarning(y_posicion, nroPagina, regCargados)
    Call printBalanceOfPIC(y_posicion, nroPagina, regCargados)
	if (isPICAuthorizated()) and ((regCargados > PIC_MAXREGISTROS__AUTHORIZATED)) then Call addNewPage(y_posicion, nroPagina)				
	Call get_lastWrite(y_posicion, true)
end Function
'----------------------------------------------------------------------------------------------
Function addNewPage(ByRef y_posicion, ByRef nroPagina)
    Call get_lastWrite(y_posicion, false)
	Call GF_newPage(oPDF)
	y_posicion = POSITION_Y
	if((PIC_idContrato > 0)or(PIC_cdAfe <> ""))then y_posicion = y_posicion + 15
	nroPagina = nroPagina + 1
	regCargados = 0
	Call get_PageWrite(nroPagina)
End Function
'----------------------------------------------------------------------------------------------
Function get_detallePedido(rsDET, y_posicion)
	dim IT_Importe, IT_artID, IT_artDS, IT_unidadDS
	
	IT_artID = rsDET("IDARTICULO")
	call getArticuloFull(IT_artID, IT_artDS, IT_unidadDS)
	IT_cantidad = rsDET("CANTIDAD")
	if (PIC_Moneda = MONEDA_PESO) then
	    IT_Importe = GF_EDIT_DECIMALS(rsDET("IMPORTEPESOS"),2)
    else	    
	    IT_Importe = GF_EDIT_DECIMALS(rsDET("IMPORTEDOLARES"),2)
    end if	    
	IT_PartPresup = rsDET("IDAREA") & " - " & rsDET("IDDETALLE")
	Call escribeRegistro(y_posicion, IT_artID, IT_artDS, IT_cantidad, IT_unidadDS, IT_PartPresup, IT_Importe)
end Function
'-----------------------------------------------------------------------------------------------
Function get_lastWrite(y_posicion, isLastPage)
dim tituInt
	if isLastPage then
		Call GF_setFont(oPDF,"ARIAL",8,0)
		tituInt = "---------------  Fin del Pedido Interno de Compra  ---------------"
		Call GF_writeTextAlign(oPDF,55,y_posicion + 5, GF_TRADUCIR(tituInt), 50, PDF_ALIGN_RIGHT)
		Call findetalle(totalImporte)
		Call finPagina(PIC_FIRMAS)	
	else
		Call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_writeTextAlign(oPDF,55, y_posicion, GF_TRADUCIR("--- Continua en la siguiente pagina ---"), 300, PDF_ALIGN_LEFT)
		Call finPagina(PIC_TEXTOFIRMAS)
	end if
end Function

'-----------------------------------------------------------------------------------------------
Function get_PageWrite(nroPagina)
	Call dibujar_encabezado(PIC_idContrato)
	Call armarInfoPrincipal(PIC_Obra, PIC_cdPedido, PIC_Proveedor, getDivisionDS(PIC_idDivision), PIC_idContrato)
	'lineas laterales del detalle
	Call GF_verticalLine(oPDF, 10, 160, 670)
	Call GF_verticalLine(oPDF, 580, 160, 670)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,530,835, GF_TRADUCIR("P�gina N� " & nroPagina), 50, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,10,835, GF_TRADUCIR("Carg�: " & PIC_usuario & " - " & GF_FN2DTE(PIC_momento) &" - Fecha de impresi�n: " & GF_FN2DTE(session("MmtoSistema"))), 50, PDF_ALIGN_LEFT)    
end Function
'-----------------------------------------------------------------------------------------------
Function get_datosCotizacion(p_idCtzElegida)
	'Leer Datos del PIC correspondiente a la CTZ elegida
	strSQL="SELECT * from TBLCTZCABECERA where IDCOTIZACION=" & p_idCtzElegida
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		if (rs("IDPEDIDO") = "0") then
			Call comprasControlAccesoCM(RES_CD)
		else
			Call comprasControlAccesoCM(RES_CC)
		end if
		PIC_idObra = 0
        if (rs("IDOBRA") = 0) then
			PIC_Obra = "Sin Partida"
		else
		    if (CLng(rs("IDOBRA")) <> OBRA_GEID) then
			    PIC_Obra = getDescripcionObra(rs("IDOBRA"))                
            else
                PIC_Obra = OBRA_GEDS
            end if                
            PIC_idObra = rs("IDOBRA")
		end if
		PIC_idPedido = rs("IDPEDIDO")			
		Call initHeaderDB(PIC_idPedido)						
		PIC_cdPedido = pct_cdPedido
		if (PIC_cdPedido = "") then PIC_cdPedido = "Sin Pedido"
		PIC_idProveedor = rs("IDPROVEEDOR")
		if (rs("IDPROVEEDOR") <> 0) then
			strSQL="select NOMEMP, NRODOC from MET001A where NROEMP=" & rs("IDPROVEEDOR")
			Call executeQueryDb(DBSITE_SQL_MAGIC, rsProv, "OPEN", strSQL)
			PIC_Proveedor = rs("IDPROVEEDOR") & " - "			
			if len(Trim(rsProv("NOMEMP"))) > 35 then	
				PIC_Proveedor = PIC_Proveedor & left(Trim(rsProv("NOMEMP")),35) & "..."
			else
				PIC_Proveedor = PIC_Proveedor & Trim(rsProv("NOMEMP"))
			end if
			PIC_Proveedor = PIC_Proveedor &  " - CUIT(" & GF_STR2CUIT(rsProv("NRODOC")) & ")"
		else
			PIC_Proveedor = "No se encontro Proveedor "
		end if
		if cDbl(rs("FECHAENTREGA")) = 0 then
			PIC_fecEntrega = Left(session("MmtoDato"), 8)
		else	
			PIC_fecEntrega = rs("FECHAENTREGA")	
		end if				
		PIC_observaciones = rs("OBSERVACIONES")
		if (PIC_observaciones = "") then PIC_observaciones = " "
		PIC_importePesos = CDbl(rs("IMPORTEPESOS"))
		PIC_importeDolares = CDbl(rs("IMPORTEDOLARES"))				
		PIC_IdDivision = rs("IDDIVISION")
		PIC_estado = rs("ESTADO")
		PIC_usuario = rs("CDUSUARIO")
		PIC_momento = rs("MOMENTO")
		PIC_TipoCambio = rs("TIPOCAMBIO")
		PIC_Moneda = rs("CDMONEDA")
		if (PIC_Moneda = MONEDA_PESO) then
		    totalImporte = GF_EDIT_DECIMALS(rs("IMPORTEPESOS"),2)
        else		    
		    totalImporte = GF_EDIT_DECIMALS(rs("IMPORTEDOLARES"),2)
		end if
		PIC_idContrato = rs("IDCONTRATO")
        'Busco si tiene AFE relacionado con el PIC
        PIC_cdAfe = "NO TIENE"
        Call executeProcedureDb(DBSITE_SQL_INTRA, rsAFEbyPIC, "TBLCTZCABECERA_GET_AFES_BY_IDCOTIZACION", PIC_idCtzElegida )    
        if not rsAFEbyPIC.Eof then            
            'En caso de que tenga, verifico si puede llegar a tener mas de un AFE
            PIC_cdAfe = ""
            while not rsAFEbyPIC.eof
                PIC_cdAfe = PIC_cdAfe & rsAFEbyPIC("CDAFE") & ","
                rsAFEbyPIC.MoveNext()
            wend
            PIC_cdAfe = left(PIC_cdAfe ,len(PIC_cdAfe )-1)
            if Len(Trim(PIC_cdAfe)) >= 40 then PIC_cdAfe = Left(Trim(PIC_cdAfe),37) & "..."
        end if
	else
		Call errorAcceso()
	end if
End Function
'***********************************************************************************
'****************	             COMIENZO DE LA PAGINA              ****************
'***********************************************************************************

PIC_idCtzElegida = GF_Parametros7("idCotizacionElegida",0,6)
accion = GF_Parametros7("accion","",6)
dim PIC_TIPO
if (accion <> ACCION_EMAIL) then
	if (PIC_idCtzElegida = 0) then Call errorAcceso()
	Set oPDF = GF_createPDF("PDFTemp")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	call PIC_armadoPDF()
	if existeAjusteCotizacion(PIC_idCtzElegida) then
		Call GF_newPage(oPDF)
		call PIC_armadoPDFAjustes()
	end if
	Call GF_closePDF(oPDF)
end if
'*********************************************
'-----------------------------------------------------------------------------------------------
Function PIC_armadoPDFAjustes()
	dim y_posicion, regAjuCargados, nroPagina, isLastPage, cdUsuario, momento
	y_posicion = POSITION_Y_AJU
	regAjuCargados = 0
	nroPagina = 1
	isLastPage = false
	strSQL = "SELECT CTZ.*, '' AS CANTIDAD FROM TBLCTZAJUSTES CTZ WHERE IDCOTIZACION=" & PIC_idCtzElegida & " AND APLICADO='" & TIPO_AFIRMACION & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rsDET, "OPEN", strSQL)
	cdUsuario = rsDET("CDUSUARIO")
	momento = rsDET("MOMENTO")
	Call get_PageWriteAjustes(nroPagina, cdUsuario, momento)
	While (not rsDET.eof)
		regAjuCargados = regAjuCargados + 1
		if (regAjuCargados > PICAJU_MAXREGISTROS__AUTHORIZATED) then
			Call GF_newPage(oPDF)
			nroPagina = nroPagina + 1
			Call get_PageWriteAjustes(nroPagina, cdUsuario, momento)
			y_posicion = POSITION_Y_AJU
			regAjuCargados = 0				
		end if 
		Call get_detallePedidoAjuste(rsDET, y_posicion)
		y_posicion = y_posicion + 20
		rsDET.movenext
	Wend
	isLastPage = true
	Call get_lastAjusteWrite(y_posicion, isLastPage)
end Function
'-----------------------------------------------------------------------------------------------
Function get_DatosFirmasAjustes(idAjuste, byref firmante1Cd, byref firmante1Ds, byref firmante1Tx, byref firmante2Cd, byref firmante2Ds, byref firmante2Tx, byref firmante3Cd, byref firmante3Ds, byref firmante3Tx, byref firmante4Cd, byref firmante4Ds, Byref firmante4Tx, byref firmante5Cd, byref firmante5Ds, Byref firmante5Tx) 
	Dim strSQL, rsFirmas
	
	firmante1Cd =""
	firmante2Cd =""
	firmante3Cd =""
	firmante4Cd =""
	firmante5Cd =""
	Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLCTZAJUSTESFIRMAS_GET_BY_IDAJUSTE", idAjuste)
	while not rsFirmas.eof
			if (firmante1Cd ="") then
				firmante1Cd = rsFirmas("CDUSUARIO")
				firmante1Ds = getUserDescription(firmante1Cd)
				firmante1Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante2Cd ="") then
				firmante2Cd = rsFirmas("CDUSUARIO")
				firmante2Ds = getUserDescription(firmante2Cd)
				firmante2Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante3Cd ="") then			
				firmante3Cd = rsFirmas("CDUSUARIO")
				firmante3Ds = getUserDescription(firmante3Cd)
				firmante3Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))					
			elseif (firmante4Cd ="") then
				firmante4Cd = rsFirmas("CDUSUARIO")
				firmante4Ds = getUserDescription(firmante4Cd)
				firmante4Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			elseif (firmante5Cd ="") then
				firmante5Cd = rsFirmas("CDUSUARIO")
				firmante5Ds = getUserDescription(firmante5Cd)
				firmante5Tx = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			end if				
		rsFirmas.movenext
	wend			
End Function
'-----------------------------------------------------------------------------------------------
Function get_PageWriteAjustes(nroPagina, pCdUsuario, pMomento)
	Call dibujar_encabezadoAjuste(PIC_idContrato)
	Call armarInfoPrincipalAjuste(PIC_Obra, PIC_cdPedido, PIC_Proveedor, getDivisionDS(PIC_idDivision), PIC_Moneda)
	'lineas laterales del detalle
	Call GF_verticalLine(oPDF, 10, 80, 750)
	Call GF_verticalLine(oPDF, 580, 80, 750)
	Call GF_horizontalLine(oPDF, 10, 830, 570)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,500,835, GF_TRADUCIR("P�gina De Ajuste N� " & nroPagina), 50, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,10,835, GF_TRADUCIR("Carg�: " & PIC_usuario & " - " & GF_FN2DTE(pMomento)) &" - Fecha de impresi�n: " & GF_FN2DTE(session("MmtoSistema")), 50, PDF_ALIGN_LEFT)    
end Function
'-----------------------------------------------------------------------------------------------
Function dibujar_encabezadoAjuste(pIdContrato)
	Dim titulo
	titulo = "AJUSTES AL PEDIDO INTERNO DE COMPRA"
	if (CLng(pIdContrato) <> 0) then titulo = "AJUSTE AL CBTE. ELECTRONICA DE CUMPLIMIENTO"
	'dibuja recuadro general
	Call GF_squareBox(oPDF, 2, 2, 590, 848, 0, "", "#000000", 2, PDF_SQUARE_ROUND)
	'logo y titulo
	Call GF_writeImage(oPDF, Server.MapPath("images\ADMlogo2.jpg"), 10, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",20,0)
	Call GF_writeTextAlign(oPDF,10, 25, GF_TRADUCIR(titulo), 570, PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"ARIAL",14,0)
	Call GF_writeTextAlign(oPDF,300, 50, GF_TRADUCIR("Nro : " & PIC_idCtzElegida), 280, PDF_ALIGN_RIGHT)
end Function
'----------------------------------------------------------------------------------------------
Function get_detallePedidoAjuste(rsDET, y_posicion)
	dim IT_Importe, IT_artID, IT_artDS, IT_unidadDS
	IT_artID = rsDET("IDARTICULO")
	call getArticuloFull(IT_artID, IT_artDS, IT_unidadDS)
	if (PIC_Moneda = MONEDA_PESO) then
	    IT_Importe = rsDET("IMPORTEPESOS")
    else	    
	    IT_Importe = rsDET("IMPORTEDOLARES")
    end if	    
	Call escribeRegistroAjuste(y_posicion, rsDET("IDAJUSTE"), IT_artID, IT_artDS, PIC_Moneda, IT_Importe, rsDET("IDAREA"), rsDET("IDDETALLE"), rsDET("OBSERVACIONES"), rsDET("CDUSUARIO"), rsDET("MOMENTO"))
end Function
'-----------------------------------------------------------------------------------------------
Function escribeRegistroAjuste(p_y, pAjuste, p_codigo, p_descripcion, p_Moneda, p_importe, pIdArea, pIdDetalle, pObservaciones, pUsuario, pMomento)

    Dim strSQL, rs
    
	call GF_setFont(oPDF,"ARIAL",10,0)
	'Cuadro gral
	Call GF_squareBox(oPDF, 20 + 40, p_y, 450, 70, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	'idAjuste
	Call GF_squareBox(oPDF, 20 + 40, p_y, 30, 10, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	'Articulo
	Call GF_squareBox(oPDF, 50 + 40, p_y, 300, 10, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	'Partida Presupuestaria
	Call GF_squareBox(oPDF, 286 + 90, p_y, 55, 10, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	'Importe 
	Call GF_squareBox(oPDF, 408 + 20, p_y, 92, 10, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	
	if (pAjuste <> "") then Call GF_writeTextAlign(oPDF,20 + 40, p_y, pAjuste, 30, PDF_ALIGN_CENTER)
	if (p_descripcion <> "") then 
		if (Len(p_descripcion) >50) then p_descripcion = Left(p_descripcion,50) & "..."
		Call GF_writeTextAlign(oPDF,55 + 40, p_y, p_codigo & " - " & p_descripcion, 275, PDF_ALIGN_LEFT)
	end if		
	Call GF_writeTextAlign(oPDF,290 + 90, p_y, pIdArea & " - " & pIdDetalle, 30, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,412 + 20, p_y, getSimboloMoneda(p_Moneda) & " " & GF_EDIT_DECIMALS(p_importe, 2), 85, PDF_ALIGN_RIGHT)	
	'OBSERVACIONES
	p_y = p_y + 10
	'Observaciones
	Call GF_squareBox(oPDF, 20 + 40, p_y, 460, 30, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	if (pObservaciones <> "") then Call GF_writeTextAlign(oPDF,21 + 40, p_y, "OBSERVACIONES", 50, PDF_ALIGN_CENTER)
	p_y = p_y + 10
	call GF_setFont(oPDF,"ARIAL",6,0)
	'if (pObservaciones <> "") then Call GF_writeTextAlign(oPDF,30 + 40, p_y, pObservaciones, 20, PDF_ALIGN_LEFT)
	if (pObservaciones <> "") then Call PF_writeTextPlus(oPDF,30 + 40, p_y, pObservaciones, 450, 6, PDF_ALIGN_LEFT)
	call GF_setFont(oPDF,"ARIAL",8,0)
	'AUTORIZACIONES	
	p_y = p_y + 20
	Call GF_squareBox(oPDF, 20 + 40, p_y, 80, 30, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 100 + 40, p_y, 380, 10, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	call get_DatosFirmasAjustes(pAjuste, PIC_firmante1Cd, PIC_firmante1Ds, PIC_firmante1Tx, PIC_firmante2Cd, PIC_firmante2Ds, PIC_firmante2Tx, PIC_firmante3Cd, PIC_firmante3Ds, PIC_firmante3Tx, PIC_firmante4Cd, PIC_firmante4Ds, PIC_firmante4Tx, PIC_firmante5Cd, PIC_firmante5Ds, PIC_firmante5Tx)
	if (PIC_firmante1Cd <> "") then 
	    Call GF_writeTextAlign(oPDF,101 + 40, p_y,PIC_firmante1Cd & " - " & PIC_firmante1Ds & " - " & PIC_firmante1Tx, 200, PDF_ALIGN_LEFT)
    else
        Call GF_writeTextAlign(oPDF,101 + 40, p_y,"No requiere autorizaci�n por ser un ajuste que reduce el monto autorizado.", 200, PDF_ALIGN_LEFT)
    end if	    
	p_y = p_y + 10
	Call GF_squareBox(oPDF, 100 + 40, p_y, 380, 10, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_writeTextAlign(oPDF,21 + 40, p_y, "AUTORIZACIONES", 80, PDF_ALIGN_CENTER)
	if (PIC_firmante2Cd <> "") then Call GF_writeTextAlign(oPDF,101 + 40, p_y,PIC_firmante2Cd & " - " & PIC_firmante2Ds & " - " & PIC_firmante2Tx, 200, PDF_ALIGN_LEFT)
	p_y = p_y + 10
	Call GF_squareBox(oPDF, 100 + 40, p_y, 380, 10, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	if (PIC_firmante3Cd <> "") then Call GF_writeTextAlign(oPDF,101 + 40, p_y,PIC_firmante3Cd & " - " & PIC_firmante3Ds & " - " & PIC_firmante3Tx, 200, PDF_ALIGN_LEFT)	
	p_y = p_y + 10
	'Call GF_squareBox(oPDF, 20 + 40, p_y, 460, 8, 0, "#ffffff", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",6,0)
	Call GF_writeTextAlign(oPDF,20 + 41, p_y+1, "Cargado por " & pUsuario & " - " & getUserDescription(pUsuario) & " el " & GF_FN2DTE(pMomento), 30, PDF_ALIGN_LEFT)
	call GF_setFont(oPDF,"ARIAL",8,0)
end Function
'-----------------------------------------------------------------------------------------------
Function armarInfoPrincipalAjuste(p_partida, p_pedido, p_Proveedor, p_division, p_fechaEntrega)
	dim tituloInt
	Call GF_squareBox(oPDF, 10, 70, 570, 15, 0, g_BgColor, "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",10,0)	
	tituloInt = "Detalles de los Ajustes al Pedido Interno de Compra"
	Call GF_writeTextAlign(oPDF,15, 72, GF_TRADUCIR(tituloInt), 565, PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"ARIAL",8,0)	
end Function
'-----------------------------------------------------------------------------------------------
Function get_lastAjusteWrite(y_posicion, isLastPage)
dim tituInt
	if isLastPage then
		Call GF_setFont(oPDF,"ARIAL",8,0)
		tituInt = "---------------  Fin de los Ajustes Realizados al Pedido Interno de Compra  ---------------"
		Call GF_writeTextAlign(oPDF,55 + 50,y_posicion + 5, GF_TRADUCIR(tituInt), 50, PDF_ALIGN_RIGHT)
	else
		Call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_writeTextAlign(oPDF,55 + 50, y_posicion, GF_TRADUCIR("--- Continua en la siguiente pagina ---"), 300, PDF_ALIGN_LEFT)
	end if
end Function
'-----------------------------------------------------------------------------------------------
%>