<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
dim myProveedorSeleccionado, rs, conn, strSQL, oPDF, ITOAProDs, ITOANroSo
dim ITComentarios, pcpType

idPedido = GF_PARAMETROS7("idPedido",0,6)

Function armadoPDF()
	Call armadoTitulo(GF_TRADUCIR("Analisis Comparativo de Ofertas"))
	Call armadoCabecera()
	Call armadoProveedores()
	Call armadoComentarios(ITComentarios)	
	Call armadoOfertaAdjudicada(ITOAProDs, ITOANroSo)	
	Call armadoFirmas()	
end Function

Function armadoTitulo(p_titulo)
	call GF_setFont(oPDF,"ARIAL",14,8)
	Call GF_writeVerticalText(oPDF,0,850,p_titulo,850,PDF_ALIGN_CENTER)
end Function

Function armadoCabecera()
	dim ITPuerto, ITObra, ITPedido, ITfechaconc
	'dibuja la tabla
	call GF_squareBox(oPDF, 20, 15, 15, 820, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_squareBox(oPDF, 35, 15, 10, 820, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_horizontalLine(oPDF,20,550,25)
    Call GF_horizontalLine(oPDF,20,735,25)
	call GF_squareBox(oPDF, 50, 715, 15, 120, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_squareBox(oPDF, 50, 15, 15, 700, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'escribe los titulos fijos
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,22,835,GF_TRADUCIR("Puerto"),100,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,22,735,GF_TRADUCIR("Fecha de Concurso"),185,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,22,550,GF_TRADUCIR("Pedido"),535,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,52,830,GF_TRADUCIR("Obra / Trabajo"),110,PDF_ALIGN_CENTER)
	'nombra las variables	
	if ((pct_idObra > 0) and (pct_idObra <> OBRA_GEID)) then	
		ITObra = GF_TRADUCIR("Sin Obra")
	else
		ITObra = getDescripcionObra(pct_idObra)
	end if
	if (pct_idObra > 0) then		
		ITPuerto = getDivisionObra(pct_idObra)	
	else
		ITPuerto = pct_dsDivision
	end if
	if (isnull(pct_cdPedido)) then
		ITPedido = GF_TRADUCIR("Sin Pedido")
	else
		ITPedido = pct_tituloPedido & " (" & pct_cdPedido & ")"
	end if
	ITfechaconc = GF_TRADUCIR("Inicio") & ": " & pct_FechaInicio & " - " & GF_TRADUCIR("Cierre") & ": " & pct_FechaCierre
	'escribe los datos
	call GF_setFont(oPDF,"ARIAL",8,8)
	Call GF_writeVerticalText(oPDF,36,835,ITPuerto,100,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,36,735,ITfechaconc,185,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,36,550,ITPedido,535,PDF_ALIGN_CENTER)
	
	Set obra = obtenerDescripcionCompletaDetalle(pct_idObra,pct_idArea,pct_idDetalle)
	if (not obra.eof) then
		strTextoObra = obra("CDOBRA") & "-" & obra("DSOBRA")
		if (not isNull(obra("IDAREA"))) then
			strTextoObra = strTextoObra & "<br/> (" & pct_idArea
			if (not isNull(obra("IDDETALLE"))) then
				strTextoObra = strTextoObra & "-" & pct_idDetalle & ":" & obra("DSDETALLE")							
			end if
			strTextoObra = strTextoObra & ")"
		end if
		Call GF_writeVerticalText(oPDF,53,710,strTextoObra,700,PDF_ALIGN_LEFT)
    else
        if (pct_idObra = OBRA_GEID) then
            Call GF_writeVerticalText(oPDF,53,710,OBRA_GECD & "-" & OBRA_GEDS,700,PDF_ALIGN_LEFT)             
        end if		
	end if
	
end Function
	
Function armadoProveedores()
	dim ITPrecio, valRsCount, valLongHorizontal, valXCeldaRegistro, valColorRelleno, valColorBorde
	dim ITnroLinea, ITproveedor, ITproveedorDS, ITcaracteristica, ITimporte, ITmoneda, ITcondPago, ITfecentrega
	ITOAProDs = " "
	ITOANroSo = " "
	'llamado a la base de datos de los proveedores
	strSQL="SELECT * from TBLPCPDETALLE where IDPEDIDO=" & idPedido & " order by NROSOBRE"		
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	'cuantos registros trae
	valRsCount = rs.RecordCount
	call GF_squareBox(oPDF, 75, 15, 15, 820, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
	'escribe los titulos fijos
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,77,835,GF_TRADUCIR("NºSobre"),50,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,77,785,GF_TRADUCIR("Proveedor"),275,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,77,500,GF_TRADUCIR("Caracteristicas"),200,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,77,310,GF_TRADUCIR("Precio"),75,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,77,235,GF_TRADUCIR("Cond. de Pago"),150,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,77,80,GF_TRADUCIR("Fecha de E."),60,PDF_ALIGN_CENTER)
	valXCeldaRegistro = 80
	pcpType = true
	'por cada reistrro dibuja la celda e ingresa los datos que trae el rs
	while not rs.eof
		valXCeldaRegistro = valXCeldaRegistro + 10
		ITproveedor = rs("IDPROVEEDOR")
		ITproveedorDS = getDescripcionProveedor(ITproveedor)
		ITcaracteristica = Trim(rs("CARACTERISTICAS"))
		ITimporte = rs("IMPORTE")
		ITimporte = Replace(ITimporte,",",".")/100
		ITmoneda = rs("CDMONEDA")
		ITcondPago = Trim(rs("CONDPAGO"))
		if ITcondPago = "" then ITcondPago = "No Especificada"
		ITfecentrega = rs("FECENTREGA")
		if len(ITfecentrega) < 2 then 
			ITfecentrega = "N/D"
		else
			ITfecentrega = GF_FN2DTE(ITfecentrega)
		end if
		ITnroLinea = rs("NROSOBRE")		
		if (CLng(myProveedorSeleccionado) = CLng(rs("IDPROVEEDOR"))) then
			valColorRelleno   = "#ffeecd"
			valColorBorde = "#f4b800"
			ITOAProDs = ITproveedorDS
			ITOANroSo = ITnroLinea
			pcpType = getPCPAuthorizationType(ITimporte,ITmoneda)
		else
			valColorRelleno   = ""
			valColorBorde = "#000000"
		end if
		Set rsAux = getCotizaciones(idPedido, ITproveedor)
		pct_hayCotizacion = false
		if (not rsAux.eof) then 
			pct_hayCotizacion = true
			pct_pathCotizacion = rsAux("PATHCOTIZACION")
		end if			
		if not pct_hayCotizacion or CStr(pct_pathCotizacion) = "NO_COTIZA" then			
			ITcaracteristica = "No Cotiza"
			ITimporte = 0
			ITcondPago = " "
			ITfecentrega = " "
			ITPrecio = " "
		else					    
	        'Verifico si alguna de las coptizacion del proveedor fue presentada fuera del palzo.
	        Set rsCotizaciones = getCotizaciones(pct_idPedido, ITproveedor)
	        blnFuera = false
	        while ((not rsCotizaciones.eof) and (not blnFuera))
	            if  (GF_DTEDIFF(rsCotizaciones("FECHAPRESENTACION"), GF_DTE2FN(pct_FechaCierre), "D") < 0) then 	    
	                ITcaracteristica = "Cotizacion cargada fuera de termino"
	                blnFuera = true
	            end if 
	            rsCotizaciones.MoveNext()
	        wend
			if(ITmoneda = MONEDA_DOLAR) then 
				ITPrecio = getSimboloMoneda(MONEDA_DOLAR) & " "
			else
				ITPrecio = getSimboloMoneda(MONEDA_PESO) & " "	
			end if
			ITPrecio = ITPrecio & GF_EDIT_DECIMALS((ITimporte*100),2) 
		end if		
		if (ITcondPago = "") then ITcondPago = " "
		call GF_squareBox(oPDF, valXCeldaRegistro, 15, 10, 820, 0, valColorRelleno, valColorBorde, 1, PDF_SQUARE_NORMAL)
		call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 835, ITnroLinea,50,PDF_ALIGN_CENTER)		
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 780, left(ITproveedorDS, 50), 275,PDF_ALIGN_LEFT)						
		call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 500, Lcase(left(ITcaracteristica, 50)), 300,PDF_ALIGN_LEFT)
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 310, ITPrecio,70,PDF_ALIGN_RIGHT)
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 235, Lcase(left(ITcondPago, 38)),150,PDF_ALIGN_CENTER)		
		Call GF_writeVerticalText(oPDF,valXCeldaRegistro + 1, 80, ITfecentrega,60,PDF_ALIGN_CENTER)
		rs.movenext
	wend	
	'dibuja separadores
	valLongHorizontal = (valRsCount * 10) + 15
	Call GF_horizontalLine(oPDF,75,80,valLongHorizontal)
	Call GF_horizontalLine(oPDF,75,235,valLongHorizontal)
	Call GF_horizontalLine(oPDF,75,310,valLongHorizontal)
	Call GF_horizontalLine(oPDF,75,505,valLongHorizontal)
	Call GF_horizontalLine(oPDF,75,785,valLongHorizontal)
end Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Function armadoResponsableTecnico(p_Responsable, p_FirmaResponsable, p_ResponsableCd)
	'dibuja celdas
	call GF_squareBox(oPDF, 225, 15, 15, 205, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
    call GF_squareBox(oPDF, 240, 15, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'datos fijos
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,227,220,GF_TRADUCIR("Resp. Técnico"),205,PDF_ALIGN_CENTER)
	'ingresa texto utilizando la funcion para escribir en varias lineas
	if p_Responsable <> "" then
		call GF_setFont(oPDF,"ARIAL",10,0)
		Call GF_writeVerticalText(oPDF,317,220,p_Responsable,205,PDF_ALIGN_CENTER)
	end if
	'dibuja fima
	if (p_FirmaResponsable <> "") then
		call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,330,220,p_FirmaResponsable,300,PDF_ALIGN_CENTER)
		Call GF_writeImage(oPDF, Server.MapPath("images\firmas\" & obtenerFirma(p_ResponsableCd)), 241, 219, 200, 75, 90)
	end if
end Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Function armadoComentarios(p_Comentarios)
	dim vTexto
	'dibuja celdas
	call GF_squareBox(oPDF, 190, 220, 15, 615, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
    call GF_squareBox(oPDF, 205, 220, 135, 615, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	'datos fijos
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,192,830,GF_TRADUCIR("Comentarios / Sugerencias Técnicas:"),610,PDF_ALIGN_LEFT)
	'ingresa texto utilizando la funcion para escribir en varias lineas
	call GF_setFont(oPDF,"ARIAL",8,0)								
	Call GF_writeVerticalTextPlus(oPDF,207,830,p_Comentarios,610, 8,PDF_ALIGN_LEFT)	
end Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Function armadoOfertaAdjudicada(p_ProDs, p_NroSo)
	'armado celdas
	call GF_squareBox(oPDF, 190, 15, 15, 205, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
    call GF_squareBox(oPDF, 205, 15, 10, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
    call GF_squareBox(oPDF, 215, 15, 10, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_horizontalLine(oPDF,205,65,20)
	'ingreso datos
	call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,192,215,GF_TRADUCIR("Oferta sugerida para adjudicar:"),205,PDF_ALIGN_LEFT)
	call GF_setFont(oPDF,"ARIAL",8,8)
	Call GF_writeVerticalText(oPDF,206,220,GF_TRADUCIR("Proveedor"),140,PDF_ALIGN_CENTER)
	Call GF_writeVerticalText(oPDF,206,65,GF_TRADUCIR("NºSobre"),50,PDF_ALIGN_CENTER)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeVerticalText(oPDF,216,215,p_ProDs,140,PDF_ALIGN_LEFT)
	Call GF_writeVerticalText(oPDF,216,65,p_NroSo,50,PDF_ALIGN_CENTER)
end Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Function armadoFirmas()
    Dim rsFirmas, strSQL, myFirmaDs, myFirmaTx, myFirmaCd
    dim member1Cd, member2Cd, member3Cd, member4Cd, responsableCd, firmaResponsable, firma1, firma2,firma3, firma4,memberDireccionCd,firmaDireccion,memberDireccion
    
    Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLPCPFIRMAS_GET_BY_IDPEDIDO", pct_idPedido)
	if (not rsFirmas.eof) then
	    'Empiezo tomando una firma, si no es la ultima la agrego al primer lugar disponible.
	    'Cuando llegue la ultima, siempre se carga en el lugar de aprobacion general de la planilla fuera del bucle.
	    myFirmaCd = rsFirmas("CDUSUARIO")
	    myFirmaDs = getUserDescription(rsFirmas("CDUSUARIO"))
	    myFirmaTx = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
	    rsFirmas.MoveNext()
	    while (not rsFirmas.eof)	    
	        'La coloco en el primer lugar disponible
	        if (responsableCd = "") then
	            responsableCd = myFirmaCd
	            firmaResponsable = myFirmaTx
	            ITResponsable = myFirmaDs
	        else
	            if (member1Cd = "") then
	                member1Cd = myFirmaCd			
		            member1 = myFirmaDs
		            firma1 = myFirmaTx
	            else
	                if (member2Cd = "") then
	                    member2Cd = myFirmaCd			
		                member2 = myFirmaDs
		                firma2 = myFirmaTx
	                else
	                    if (member3Cd = "") then
	                        member3Cd = myFirmaCd			
		                    member3 = myFirmaDs
		                    firma3 = myFirmaTx
                        else
	                        if (member4Cd = "") then
	                            member4Cd = myFirmaCd			
		                        member4 = myFirmaDs
		                        firma4 = myFirmaTx
	                        end if	        	  		                    
	                    end if	        	        
                    end if
                end if
            end if	                    
	        'Tomo la proxima firma.	    
	        myFirmaCd = rsFirmas("CDUSUARIO")
	        myFirmaDs = getUserDescription(rsFirmas("CDUSUARIO"))	        
	        myFirmaTx = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
            rsFirmas.MoveNext()    	    
	    wend
	    'Siempre que salgo del bucle tengo listos los datos de la ultima firma para completar.
	    memberDireccionCd = myFirmaCd
		firmaDireccion = myFirmaTx
		memberDireccion = myFirmaDs
	end if		
    'Muestro las firmas en la planilla.
    Call armadoResponsableTecnico(ITResponsable, firmaResponsable, responsableCd)
	Call armadoMiembrosComite(member1, member1Cd, firma1, member2, member2Cd, firma2, member3, member3Cd, firma3, member4, member4Cd, firma4, memberDireccion, memberDireccionCd, firmaDireccion)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Function armadoMiembrosComite(p_member1, p_member1Cd, p_firma1, p_member2, p_member2Cd, p_firma2, p_member3, p_member3Cd, p_firma3, p_member4, p_member4Cd, p_firma4, p_memberDireccion, p_memberDireccionCd, p_firmaDireccion)
	'armado celdas
	Call GF_squareBox(oPDF, 350, 15, 15, 820, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_squareBox(oPDF, 365, 630, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 365, 425, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 365, 220, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 365, 15, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)	
		
	'ingreso datos fijos
	Call GF_setFont(oPDF,"ARIAL",10,8)
	Call GF_writeVerticalText(oPDF,352,830,GF_TRADUCIR("Miembros del comité de adjudicación:"),610,PDF_ALIGN_LEFT)
	'ingreso datos de miembros		
	Call GF_setFont(oPDF,"ARIAL",12,0)	
	if (p_firma1 <> "") then	    
		Call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,457,840,p_firma1,300,PDF_ALIGN_CENTER)
    	Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(p_member1Cd)), 382, 838, 200, 75, 90)		
    	Call GF_setFont(oPDF,"ARIAL",12,0)
	end if			
	if (p_member1 <> "") then Call GF_writeVerticalText(oPDF,442,840,getUserDescription(p_member1),205,PDF_ALIGN_CENTER)	    	
	if (p_firma2 <> "") then	    
		Call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,457,635,p_firma2,300,PDF_ALIGN_CENTER)
    	Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(p_member2Cd)), 382, 633, 200, 75, 90)		
    	Call GF_setFont(oPDF,"ARIAL",12,0)
	end if			
	if (p_member2 <> "") then Call GF_writeVerticalText(oPDF,442,635,getUserDescription(p_member2),205,PDF_ALIGN_CENTER)	    	
	if (p_firma3 <> "") then
		call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,457,430,p_firma3,300,PDF_ALIGN_CENTER)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(p_member3Cd)), 382, 428, 200, 75, 90)
		Call GF_setFont(oPDF,"ARIAL",12,0)
	end if	
	if (p_member3 <> "") then Call GF_writeVerticalText(oPDF,442,430,p_member3,205,PDF_ALIGN_CENTER)		
	if (p_firma4 <> "") then
		call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,457,225,p_firma4,300,PDF_ALIGN_CENTER)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(p_member4Cd)), 382, 223, 200, 75, 90)
		Call GF_setFont(oPDF,"ARIAL",12,0)
	end if
	if (p_member4 <> "") then Call GF_writeVerticalText(oPDF,442,225,p_member4,205,PDF_ALIGN_CENTER)
	call GF_squareBox(oPDF, 465, 15, 15, 205, 0, "#dcf7dc", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_squareBox(oPDF, 480, 15, 100, 205, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",10,8)
	
    Call GF_writeVerticalText(oPDF,467,215,GF_TRADUCIR("Aprobación Final de la Pedido"),200,PDF_ALIGN_LEFT)    
    
	if (p_memberDireccion <> "") then
		call GF_setFont(oPDF,"ARIAL",12,0)
		Call GF_writeVerticalText(oPDF,557,220,p_memberDireccion,205,PDF_ALIGN_CENTER)
	end if					
	if (p_firmaDireccion <> "") then
		call GF_setFont(oPDF,"ARIAL",5,8)
		Call GF_writeVerticalText(oPDF,572,220,p_firmaDireccion,300,PDF_ALIGN_CENTER)
		Call GF_writeImage(oPDF, server.MapPath("images\firmas\" & obtenerFirma(p_memberDireccionCd)), 482, 218, 200, 75, 90)		
	end if
	
end Function
'**************************************************************************************************************************
'************************************ COMIENZO DE PAGINA  **************************************************************
'**************************************************************************************************************************

Call initHeader(idPedido)
myProveedorSeleccionado = pct_idProveedorElegido
strSQL="SELECT * from TBLPCPCABECERA where IDPEDIDO=" & idPedido
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.eof) then ITComentarios = rs("COMENTARIOS")

Set oPDF = GF_createPDF("PDFTemp")
Call PDFGirarHoja(90)
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF()
Call GF_closePDF(oPDF)

%>