<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosAFE.asp"-->
<!-- #include file="Includes/procedimientosCompras.asp"-->
<!-- #include file="Includes/procedimientosPCT.asp"-->
<!-- #include file="Includes/procedimientosObras.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosMath.asp"-->
<%

Function dibujarPagina()
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\ADMlogo2.jpg"), 10, 10, 60, 60, 0)
End Function
'----------------------------------------------------------------------------------
'MODIFICACION:   dibujarTexto()
'CNA-Ajaya Nahuel
' Se agrego la funcion que permite tomar el texto de la base de datos y convertir sus espacios  (<br>) en saltos de lineas,
'ademas se agrego la nota al pie que indica el usuario y fecha q lo cargo
'----------------------------------------------------------------------------------
Function dibujarTexto()
	dim texto,y
	Call GP_ConfigurarMomentos()
	Call GF_setFont(oPDF,"ARIAL",10,0)	
	Call GF_writeTextAlign(oPDF,5,40, "Buenos Aires " & GF_DateGet("D", session("MmtoSistema"))&" de "& GF_INT2MES(GF_DateGet("M", session("MmtoSistema")))&" del "&GF_DateGet("A", session("MmtoSistema")), 520,PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,10,835, GF_TRADUCIR("Cargó: " & rsNot("CDUSRCARGA") & " - " & GF_FN2DTE(rsNot("MOMENTO"))), 50, PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"ARIAL",16,0)
	Call GF_writeTextAlign(oPDF,5,80, GF_TRADUCIR("NOTA DE ACEPTACION") , 580 , PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,205,97,178)	
	Call GF_setFont(oPDF,"ARIAL",10,0)	
	texto = replace((rsNot("DSMENSAJE")&""), vbCrLf, "<br />")
	Call GF_writeTextPlus(oPDF, 15, 140, texto,520, 15,PDF_ALIGN_JUSTIFY)	
End Function
'----------------------------------------------------------------------------------
Function CrearPdf(p_idPedido,p_idCotizacion,p_idNDA,p_mode)	
	dim rs
	if(p_idPedido > 0) then
		set rs = cargaMensajeNDA(p_idPedido,p_idCotizacion)	'con esta funcion cargo solo el codigo de pedido'
		pathPDF = Server.MapPath("temp\NDA-REF-" & rs("cdpedido") & ".pdf" )
	else	
		'se hace referencia al IdCotizacion'
		pathPDF = Server.MapPath("temp\NDA-REF-" & p_idCotizacion & ".pdf" )		
	end if	
	Set oPDF = GF_createPDF(pathPDF)
	Call GF_setPDFMODE(p_mode)			
	Set rsNot = getRsNotaAceptacion(p_idNDA)
	Call dibujarPagina()
	Call dibujarTexto()	
	
	Call GF_closePDF(oPDF)	
	
	CrearPdf=pathPDF
End Function

'******************************************************************************************************
'**********************************		INICIO DE LA PAGINA 	*******************************
'******************************************************************************************************
Dim oPDF
Dim dsProveedor,fechaPedido,cdPedido,tituloPedido,admin
Dim idPedido
Dim tipoCompra,cdadmin,rsNot

idPedido = GF_Parametros7("idPedido",0,6)
idCotizacion = GF_Parametros7("idcotizacion",0,6)
idNDA = GF_Parametros7("IdNDA",0,6)

if (LCase(request.servervariables("script_name")) = "/intranet/comprasnotadeaceptacionprint.asp") then
	Call CrearPdf(idPedido,idCotizacion,idNDA,PDF_STREAM_MODE)
end if

%>
