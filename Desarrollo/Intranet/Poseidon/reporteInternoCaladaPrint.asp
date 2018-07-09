<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<%
Const PY_INICIO_ORIGINAL = 2
Const PY_INICIO_DUPLICADO = 430
Const PY_OBSERVACIONES_ORIGINAL = 342
Const PY_OBSERVACIONES_DUPLICADO = 770
Const PARAM_LEYENDA_LUGAR_PTO = "LEYENDALUGAR"
'*****************************************************************************************************************************
'*****************************************************************************************************************************
Function armadoPDF(ByRef p_rsCalada, pY_observaciones, p_Puerto)
	GP_CONFIGURARMOMENTOS
    Call dibujarEncabezado(p_Puerto)
    if (not p_rsCalada.Eof) then
        Call dibujarCabecera(p_rsCalada)
        Call dibujarTitulosDetalle()
        while (not p_rsCalada.Eof)
            Call dibujarDetalleRubros(p_rsCalada)
            p_rsCalada.MoveNext()
        wend
        Call dibujarObservaciones(pY_observaciones)
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarEncabezado(p_Puerto)
	Dim momento,leyendaPto
    momento = GF_FN2DTE(session("MmtoSistema"))    
    Call GF_squareBox(oPDF, 2, pY, 590, 400, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
    pY = pY + 8	
	Call GF_writeImage(oPDF, Server.MapPath("..\images\ADMlogo2.jpg"), 20, pY, 81, 75, 0)
	'Titulo reporte
    call GF_setFont(oPDF,"ARIAL",16,8)
    pY = pY + 15
	Call GF_writeTextAlign(oPDF,2, pY, GF_TRADUCIR("INTERNO DE CALADA"), 590, PDF_ALIGN_CENTER)
    'Fecha, hora, puerto
    Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , pY, "Fecha: " & Left(momento,10), 555 , PDF_ALIGN_RIGHT)
    pY = pY + 12
    Call GF_writeTextAlign(oPDF, 10 , pY, "Hora: " & Right(momento,8), 555 , PDF_ALIGN_RIGHT)
    pY = pY + 12
    leyendaPto = getValueParametro(PARAM_LEYENDA_LUGAR_PTO,p_Puerto)
    Call GF_writeTextAlign(oPDF, 10 , pY, leyendaPto, 555 , PDF_ALIGN_RIGHT)
    pY = pY + 30
end Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCabecera(p_RsCalada)
    
    call GF_setFont(oPDF,"ARIAL",12,8)
	Call GF_writeTextAlign(oPDF,200, pY, GF_TRADUCIR("PRODUCTO: ") & Ucase(Trim(p_RsCalada("DSPRODUCTO"))), 380, PDF_ALIGN_LEFT)
    pY = pY + 12
    Call GF_writeTextAlign(oPDF,200, pY, GF_TRADUCIR("PATENTE: ") & GF_EDIT_PATENTE(p_RsCalada("CDCHAPACAMION")), 140, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,340, pY, GF_TRADUCIR("TURNO: ") & p_RsCalada("SQTURNO"), 280, PDF_ALIGN_LEFT)
    pY = pY + 30
    
    call GF_setFont(oPDF,"ARIAL",9,8)
	Call GF_writeTextAlign(oPDF,20, pY, GF_TRADUCIR("TARJETA: ") & p_RsCalada("IDCAMION"), 270, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,285, pY, GF_TRADUCIR("RECIBIDOR: ") & Ucase(p_RsCalada("DSLASTNAME")) &", "& Ucase(p_RsCalada("DSNAME")), 300, PDF_ALIGN_LEFT)
    pY = pY + 12
    Call GF_writeTextAlign(oPDF,20, pY, GF_TRADUCIR("CARTA DE PORTE: ") & GF_EDIT_CTAPTE(p_RsCalada("NUCARTAPORTE")), 150, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,185, pY, GF_TRADUCIR("CTG: ") & p_RsCalada("CTG"), 70, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,285, pY, GF_TRADUCIR("FECHA Y HORA: ") & GF_FN2DTE(p_RsCalada("DTCALADA")), 300, PDF_ALIGN_LEFT)
    pY = pY + 12
    Call GF_writeTextAlign(oPDF,20, pY, GF_TRADUCIR("PROCEDENCIA: ") & Ucase(Trim(p_RsCalada("DSPROCEDENCIA"))), 560, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,285, pY, GF_TRADUCIR("ACEPTACION: ") & Ucase(p_RsCalada("DSACEPTACION")) , 300, PDF_ALIGN_LEFT)
    
    pY = pY + 35
End function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTitulosDetalle()
    call GF_setFont(oPDF,"COURIER",10,0) 
    Call GF_writeTextAlign(oPDF,20, pY, GF_TRADUCIR("CALIDAD"), 170, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,190, pY, GF_TRADUCIR("VL"), 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,250, pY, GF_TRADUCIR("%"), 60, PDF_ALIGN_CENTER)
    py = py + 10
    call GF_horizontalLine(oPDF, 20, py, 295)
    py = py + 5
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarDetalleRubros(p_RsCalada)
    Call GF_writeTextAlign(oPDF,20, pY, UCASE(p_RsCalada("DSRUBRO")), 170, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,190, pY, GF_EDIT_DECIMALS(cDbl(p_RsCalada("VLBONREBAJA"))*100,2), 60, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(oPDF,250, pY, GF_EDIT_DECIMALS(cDbl(p_RsCalada("VLMERMA"))*100,2), 60, PDF_ALIGN_RIGHT)
    py = py + 10
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarObservaciones(ByRef pY_observaciones)
    Call GF_writeTextAlign(oPDF,20, pY_observaciones, GF_TRADUCIR("OBSERVACIONES:"), 150, PDF_ALIGN_LEFT)
    call GF_setFont(oPDF,"COURIER",10,0) 
    Call GF_writeTextAlign(oPDF,105, pY_observaciones, string(78,"."), 450, PDF_ALIGN_LEFT)
    pY_observaciones = pY_observaciones + 15
    Call GF_writeTextAlign(oPDF,20, pY_observaciones, string(92,"."), 450, PDF_ALIGN_LEFT)
    pY_observaciones = pY_observaciones + 15
    Call GF_writeTextAlign(oPDF,20, pY_observaciones, string(92,"."), 450, PDF_ALIGN_LEFT)
    pY_observaciones = pY_observaciones + 15
End Function
'****************************************************************************************************************************
'********************************	             COMIENZO DE LA PAGINA              *****************************************
'****************************************************************************************************************************

Dim puerto,fecha,idCamion,oPDF,pY, rsCalada
Dim nroCopias, cantCopiasAux


idCamion = GF_Parametros7("idcamion","",6)
fecha = GF_Parametros7("dtcontable","",6) 'Formato AAAAMMDD
puerto = GF_Parametros7("pto","",6)
nroCopias = GF_Parametros7("nroCopias",0,6)
cantCopiasAux = 1

Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)

Call GF_BD_Puertos(puerto, rsCalada, "OPEN","EXEC HCAMIONES_GET_INTERNO_CALADA_BY_PARAMETERS '"& idCamion &"',"& fecha )
	if (not rsCalada.Eof) then
		py = PY_INICIO_ORIGINAL
		call armadoPDF(rsCalada, PY_OBSERVACIONES_ORIGINAL,puerto)
		rsCalada.MoveFirst()
		py = PY_INICIO_DUPLICADO
		call armadoPDF(rsCalada, PY_OBSERVACIONES_DUPLICADO,puerto)
		if nroCopias > 1 then
			while cantCopiasAux < nroCopias
				Call GF_newPage (oPDF)
				rsCalada.MoveFirst()
				'Call GF_BD_Puertos(puerto, rsCalada, "OPEN","EXEC HCAMIONES_GET_INTERNO_CALADA_BY_PARAMETERS '"& idCamion &"',"& fecha )
				if (not rsCalada.Eof) then
					py = PY_INICIO_ORIGINAL
					call armadoPDF(rsCalada, PY_OBSERVACIONES_ORIGINAL,puerto)
					rsCalada.MoveFirst()
					py = PY_INICIO_DUPLICADO
					call armadoPDF(rsCalada, PY_OBSERVACIONES_DUPLICADO,puerto)
					cantCopiasAux = cantCopiasAux + 1
				end if
			wend
		end if
	else
		py = PY_INICIO_ORIGINAL 
		Call dibujarEncabezado(puerto)
		Call GF_writeTextAlign(oPDF,20, py , GF_TRADUCIR("No se encontraron resultados"), 560, PDF_ALIGN_CENTER)
		py = PY_INICIO_DUPLICADO
		Call dibujarEncabezado(puerto)
		Call GF_writeTextAlign(oPDF,20, py , GF_TRADUCIR("No se encontraron resultados"), 560, PDF_ALIGN_CENTER)
	end if
Call GF_closePDF(oPDF)




%>