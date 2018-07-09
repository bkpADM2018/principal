<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="reporteCaladorCommon.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<%Response.Buffer = False

Const INDEX_CDRUBRO         = 0
Const INDEX_ABRUBRO         = 1 
Const INDEX_VLRUBRO         = 2
Const INDEX_DSRUBRO         = 3
Const INDEX_NETO_SIN_MERMA  = 13

const ANGLE_RECTO           = 90

const SIZE_SEPARATION       = 13
const SIZE_TITULOS          = 30
const SIZE_SUBTITULOS       = 12
const SIZE_TEXTO            = 10

const WIDTH_TABLE_CABECERA  = -819
const WIDTH_TABLE_DETALLE   = -390

const HEIGHT_BOX_SUBTITULO  = 14 
const HEIGHT_CABECERA       = 56

const INICIO_SUBTABLAS      = 633
const INICIO_EJE_H          = 835
const INICIO_EJE_V          = 13
const MAX_H_PAGINA          = 845
Const MAX_V_PAGINA			= 570
'-------------------------------------------------------------------------------------------------------------------------
Function armadoPDF()
    DIM eje_V, eje_H, sizeControl
    Call dibujarPagina()
    Call GF_writeImage(oPDF, Server.MapPath("..\images\kogge64.gif"),INICIO_EJE_V,INICIO_EJE_H+5,60,60,ANGLE_RECTO)
    Call dibujarDatosSession()
    eje_V = dibujarTituloReporte("Reporte de Calador")
    eje_V = dibujarFiltros(eje_V)
    IF (g_chkCamiones = ESTADO_ACTIVO) THEN
        If (dicCam.Count > 0) Then
            for each strItem in dicCam.Items
                str = Split(strItem, SECTOR_TOKEN)      'LA FUNCION SPLIT DIVIDE LOS DATOS DE CABECERA Y DETALLE
                netoSinMerma = sinMerma(str(0))         'SE GUARDA UN DATO DE LA CABECERA PARA UTILIZARLA EN DETALLE     
                if (g_chkResumen <> ESTADO_ACTIVO) then 'VERIFICANDO SI SOLO SE IMPRIME EL RESUMEN DE CAMIONES
                    eje_V = dibujarCabeceras(eje_V,str(0))
                    eje_V = dibujarDetalle(eje_V,str(1),netoSinMerma)
                else
                    call totalizarRubro(str(1),netoSinMerma)    'SOLO RESUMEN
                end if
            next
            sizeControl = ((dicTest.count + 1) * HEIGHT_BOX_SUBTITULO ) + eje_V
            if sizeControl > MAX_V_PAGINA then eje_V = nuevaPagina()
            eje_V = dibujarTotales(eje_V)               'DUBJA LOS TOTALES DEL PROMEDIO
            dicTest.RemoveAll
            dicContRubro.RemoveAll
        Else
            Call GF_squareBox(oPDF, eje_V, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#DCDCDC", "#DCDCDC", 1, PDF_SQUARE_NORMAL)
            Call GF_setFontColor("#000000")
            Call GF_setFont(oPDF,"ARIAL",SIZE_SUBTITULOS,FONT_STYLE_BOLD)
            call GF_writeVerticalText(oPDF,eje_V,MAX_H_PAGINA,GF_TRADUCIR("No se encontraron camiones"), MAX_H_PAGINA , PDF_ALIGN_CENTER )
            eje_V = eje_V  + SIZE_SEPARATION
        End if
        If (g_chkVagones = ESTADO_ACTIVO) then eje_V = nuevaPagina()
    END IF
    IF (g_chkVagones = ESTADO_ACTIVO) THEN
        If (dicVag.Count > 0) Then
            for each strItem in dicVag.Items
                str = Split(strItem, SECTOR_TOKEN)          'LA FUNCION SPLIT DIVIDE LOS DATOS DE CABECERA Y DETALLE
                netoSinMerma = sinMerma(str(0))             'SE GUARDA UN DATO DE LA CABECERA PARA UTILIZARLA EN DETALLE
                if (g_chkResumen <> ESTADO_ACTIVO) then     'VERIFICANDO SI SOLO SE IMPRIME EL RESUMEN DE VAGONES
                    eje_V = dibujarCabeceras(eje_V,str(0))
                    eje_V = dibujarDetalle(eje_V,str(1),netoSinMerma)
                else
                    call totalizarRubro(str(1),netoSinMerma)    'SOLO RESUMEN
                end if
            next
            sizeControl = ((dicTest.count + 1) * HEIGHT_BOX_SUBTITULO ) + eje_V
            if sizeControl > MAX_V_PAGINA then eje_V = nuevaPagina()
	        eje_V = dibujarTotales(eje_V)                   'DUBJA LOS TOTALES DEL PROMEDIO
            dicTest.RemoveAll
            dicContRubro.RemoveAll
        Else
            Call GF_squareBox(oPDF, eje_V, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#DCDCDC", "#DCDCDC", 1, PDF_SQUARE_NORMAL)
            Call GF_setFontColor("#000000")
            Call GF_setFont(oPDF,"ARIAL",SIZE_SUBTITULOS,FONT_STYLE_BOLD)
            call GF_writeVerticalText(oPDF,eje_V,MAX_H_PAGINA,GF_TRADUCIR("No se encontraron vagones"), MAX_H_PAGINA, PDF_ALIGN_CENTER )
            eje_V = eje_V  + SIZE_SEPARATION
        End if
    END IF
    'FIN DEL REPORTE
    call GF_setFont(oPDF,"ARIAL",SIZE_TEXTO,FONT_STYLE_NORMAL)
    Call GF_writeVerticalText(oPDF, MAX_V_PAGINA + 8 , MAX_H_PAGINA,"Fin del Reporte", MAX_H_PAGINA,PDF_ALIGN_CENTER)   
End Function
'-------------------------------------------------------------------------------------------------------------------------
'USUARIO Y MMTO DEL SISTEMA
Function dibujarDatosSession () 
    GP_CONFIGURARMOMENTOS
    Call GF_setFont(oPDF,"COURIER",SIZE_TEXTO,FONT_STYLE_NORMAL)
    Call GF_writeVerticalText(oPDF,INICIO_EJE_V,MAX_H_PAGINA,GF_FN2DTE(session("MmtoSistema")),MAX_H_PAGINA-5,PDF_ALIGN_RIGHT)
    posicion_V = INICIO_EJE_V + SIZE_SEPARATION
    Call GF_writeVerticalText(oPDF,posicion_V,MAX_H_PAGINA,         session("Usuario")      , MAX_H_PAGINA-5 ,PDF_ALIGN_RIGHT)
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Dibuja TITULO DEL REPORTE
Function dibujarTituloReporte(pTitulo)
DIM posicion_V
    posicion_V = INICIO_EJE_V + SIZE_SEPARATION    
    Call GF_setFontColor("#000000")    
    Call GF_setFont(oPDF,"ARIAL",SIZE_TITULOS,FONT_STYLE_BOLD)
    call GF_writeVerticalText(oPDF,posicion_V ,MAX_H_PAGINA ,pTitulo ,MAX_H_PAGINA  ,PDF_ALIGN_CENTER)
    posicion_V = posicion_V + (SIZE_SEPARATION*2) + SIZE_TITULOS
    Call GF_squareBox(oPDF, posicion_V, 3, 1, MAX_H_PAGINA, 0, "#FFFFFF", "#000000", 2, PDF_SQUARE_ROUND)'BORDE INFERIOR
    dibujarTituloReporte = posicion_V
end function
'------------------------------------------------------------------------------------------------------------------------
Function dibujarFiltros(p_EjeVertical)
    Dim auxCoordinador, auxCordinado, auxProducto,auxFechaDesde,auxCalador,auxFechaHasta,auxChkResumen
	p_EjeVertical = p_EjeVertical + SIZE_SEPARATION
	call GF_setFont(oPDF,"ARIAL",SIZE_TEXTO,FONT_STYLE_NORMAL)		

    auxFechaDesde = "Todas" 
    if (g_FechaDesde <> "") then auxFechaDesde = g_FechaDesde
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H     , GF_TRADUCIR("Fecha Desde:"), 200 , PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-150 ,        auxFechaDesde       , 200 , PDF_ALIGN_LEFT)
    	
    auxFechaHasta = "Todas" 
	if (g_FechaHasta <> "") then auxFechaHasta = g_FechaHasta
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-300, GF_TRADUCIR("Fecha Hasta:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-450,        auxFechaHasta	     , 200, PDF_ALIGN_LEFT)

    auxChkResumen = "NO"
	if(g_chkResumen = 1) then auxChkResumen = "SI" 'SI g_chkResumen ES IGUAL A 1(uno) SOLO MOSTRARA EL PROMEDIO(resumen) 
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-600, GF_TRADUCIR("Solo Resumen:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-750,        auxChkResumen	      , 200, PDF_ALIGN_LEFT)
	p_EjeVertical = p_EjeVertical + SIZE_SEPARATION
    '**************************************************
    auxUsuario = "Todos"
	if (g_cdUsuario <> "") then	auxUsuario = Trim(g_cdUsuario)&"-"&Trim(getUserDescription(g_cdUsuario))
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H    , GF_TRADUCIR("Usuario:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-150,         auxUsuario	     , 200, PDF_ALIGN_LEFT)

    auxCoordinado = "Todos"
	if (g_cdCoordinado > 0) then	auxCoordinado = Trim(g_cdCoordinado)&"-"&Trim(g_dsCoordinado) 
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-300, GF_TRADUCIR("Coordinado:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-450,        auxCoordinado	    , 200, PDF_ALIGN_LEFT)

    auxChkCamiones = "NO"
	if(g_chkCamiones = 1) then auxChkCamiones = "SI"
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-600, GF_TRADUCIR("Ver Camiones:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-750,        auxChkCamiones	      , 200, PDF_ALIGN_LEFT)
	p_EjeVertical = p_EjeVertical + SIZE_SEPARATION
    '**************************************************
    auxCorredor = "Todos"
	if (g_cdCorredor > 0) then	auxCorredor = Trim(g_cdCorredor)&"-"&Trim(g_dsCorredor)
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H    , GF_TRADUCIR("Corredor:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-150,        auxCorredor	      , 200, PDF_ALIGN_LEFT)

    auxVendedor = "Todos"
	if (g_cdVendedor > 0) then	auxVendedor = Trim(g_cdVendedor)&"-"&Trim(g_dsVendedor)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-300, GF_TRADUCIR("Vendedor:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-450,        auxVendedor 	  , 200, PDF_ALIGN_LEFT)

    auxChkVagones = "NO"
	if(g_chkVagones = 1) then auxChkVagones = "SI"
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-600, GF_TRADUCIR("Ver Vagones:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-750,        auxChkVagones	      , 200, PDF_ALIGN_LEFT)
	p_EjeVertical = p_EjeVertical + SIZE_SEPARATION
    '**************************************************
    auxAceptacion = "Todos"
	if (g_cdAceptacion) then auxAceptacion = Trim(getDsAceptacion(g_cdAceptacion))
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H    , GF_TRADUCIR("Aceptacion:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-150,        auxAceptacion	    , 200, PDF_ALIGN_LEFT)

	auxProducto = "Todos"
	if(g_cdProducto <> 0)then auxProducto = Trim(g_cdProducto) & "-" & Trim(getDsProducto(g_cdProducto))	
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-300, GF_TRADUCIR("Producto:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-450,        auxProducto       , 200, PDF_ALIGN_LEFT)
    p_EjeVertical = p_EjeVertical + SIZE_SEPARATION
    '**************************************************
    auxRubro = "Todos"
	if(g_cdRubro > 0)then auxRubro = g_cdRubro & "-" & getDsRubro(g_cdRubro)
    	Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H    , GF_TRADUCIR("Rubro:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-150,        auxRubro	   , 200, PDF_ALIGN_LEFT)

	auxChkPromediar = "NO"
	if(g_chkPromediar = 1) then auxChkPromediar = "SI"
	    Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-300, GF_TRADUCIR("Promediar:") , 200, PDF_ALIGN_LEFT)
        Call GF_writeVerticalTExt(oPDF, p_EjeVertical, INICIO_EJE_H-450,        auxChkPromediar	   , 200, PDF_ALIGN_LEFT)
    p_EjeVertical = p_EjeVertical + (SIZE_SEPARATION*2)
    '**************************************************
    dibujarFiltros = p_EjeVertical
End Function
'------------------------------------------------------------------------------------------------------------------------
'-----------------------------------FUNCIONES RELACIONADAS A CABECERA----------------------------------------------------
Function dibujarCabeceras(p_ejeVertical,p_StrCab) 'CUERPO DE LA CABECERA
    Dim auxEjeV
    auxEjeV = p_EjeVertical
    call dibujarRecuadroCabecera(auxEjeV)
    Call llenarTitulosCabecera(auxEjeV,arrTitulosVagones)
    auxEjeV = llenarDatosCabecera(auxEjeV,p_StrCab)
    if ((auxEjeV+HEIGHT_CABECERA) >= MAX_V_PAGINA) then auxEjeV = nuevaPagina() 'SE COMPRUEBA SI LA SIGUIENTE CABECERA ENTRA EN EL RECUADRO DE PAGINA  
    dibujarCabeceras = auxEjeV
End Function
'------------------------------------------------------------------------------------------------------------------------
Function dibujarRecuadroCabecera(p_ejeVertical) 'ESQUELETO DE LA CABECERA
    DIM auxEjeV
    auxEjeV = p_EjeVertical
    Call GF_squareBox(oPDF, auxEjeV, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#517B4A", "#517B4A", 1, PDF_SQUARE_NORMAL)
    auxEjeV = auxEjeV + HEIGHT_BOX_SUBTITULO
    Call GF_squareBox(oPDF, auxEjeV, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#DCDCDC", "#DCDCDC", 1, PDF_SQUARE_NORMAL)
    auxEjeV = auxEjeV + HEIGHT_BOX_SUBTITULO
    Call GF_squareBox(oPDF, auxEjeV, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#517B4A", "#517B4A", 1, PDF_SQUARE_NORMAL)
    auxEjeV = auxEjeV + HEIGHT_BOX_SUBTITULO
    Call GF_squareBox(oPDF, auxEjeV, INICIO_EJE_H, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_CABECERA, 0, "#DCDCDC", "#DCDCDC", 1, PDF_SQUARE_NORMAL)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function llenarTitulosCabecera(p_ejeVertical,pArr)
    DIM auxEjeV
    auxEjeV = p_ejeVertical
    Call GF_setFontColor("#FFFFFFF")
    Call GF_setFont(oPDF,"ARIAL",SIZE_SUBTITULOS,FONT_STYLE_BOLD)    
    for i = 0 to 8      '1ra PARTE TITUTLOS
        call GF_writeVerticalText(oPDF,auxEjeV,arrPositionTitulos(i),pArr(i), arrWidthTitulos(i) , PDF_ALIGN_CENTER)
    next
        auxEjeV = auxEjeV + (HEIGHT_BOX_SUBTITULO*2)
    for i = 9 to 17     '2da PARTE TITULOS
        call GF_writeVerticalText(oPDF,auxEjeV,arrPositionTitulos(i),pArr(i), arrWidthTitulos(i) , PDF_ALIGN_CENTER)
    next
End Function
'-------------------------------------------------------------------------------------------------------------
Function llenarDatosCabecera(p_ejeVertical,pCbra) 
    Dim myField, h, rtrn , auxEjeV
    auxEjeV = p_EjeVertical + HEIGHT_BOX_SUBTITULO
    myField = Split(pCbra, FIELD_TOKEN)
    Call GF_setFontColor("#006400")
    Call GF_setFont(oPDF,"ARIAL",SIZE_TEXTO,PDF_SQUARE_NORMAL)
    For h = 0 To 8 'DETALLE 1ra PARTE
	    rtrn = Split(myField(h), "=")	
	    if (h = INDEX_NETO_SIN_MERMA) then pMerma = rtrn(1)
        if LEN (rtrn(1)) => 20 then 'CONTROL DEL LARGO DE LA PALABRA
   	        call GF_writeVerticalText(oPDF,auxEjeV+2,arrPositionTitulos(h),Mid(rtrn(1),1,18)&"...", arrWidthTitulos(i) , PDF_ALIGN_CENTER )
        else
            call GF_writeVerticalText(oPDF,auxEjeV+2,arrPositionTitulos(h),rtrn(1), arrWidthTitulos(i) , PDF_ALIGN_CENTER )
        end if
    Next
        auxEjeV = auxEjeV + HEIGHT_BOX_SUBTITULO*2
    For h = 9 To 17 'DETALLE 2da PARTE
	    rtrn = Split(myField(h), "=")		
	    if (h = INDEX_NETO_SIN_MERMA) then pMerma = rtrn(1)
        if LEN (rtrn(1)) => 20 then 'CONTROL DEL LARGO DE LA PALABRA
   	        call GF_writeVerticalText(oPDF,auxEjeV+2,arrPositionTitulos(h),Mid(rtrn(1),1,18)&"...", arrWidthTitulos(i) , PDF_ALIGN_CENTER )
        else
            call GF_writeVerticalText(oPDF,auxEjeV+2,arrPositionTitulos(h),rtrn(1), arrWidthTitulos(i) , PDF_ALIGN_CENTER )
        end if
    Next 
    auxEjeV = auxEjeV + HEIGHT_BOX_SUBTITULO*2
    llenarDatosCabecera = auxEjeV
End Function
'-------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------FUNCIONES RELACIONADAS A DETALLE------------------------------------------------
Function dibujarDetalle(p_ejeVertical,p_strDet,p_netoSMerma) 'CUERPO DEL DETALLE
    Dim auxEjeV
    auxEjeV = p_EjeVertical
    myRegistro = Split(p_strDet, DETAIL_TOKEN)
    Call dibujarRecuadroDetalle(auxEjeV,UBound(myRegistro))
    auxEjeV =  llenarTituloDetalle(auxEjeV)
    For h = 0 To UBound(myRegistro)
        Call llenarDetalle(auxEjeV,myRegistro(h))
        Call guardarPromedioRubro(p_netoSMerma)     'REQUERIDA PARA LUEGO USARLO EN TOTALIZAR
        auxEjeV = auxEjeV + SIZE_SEPARATION
    next
    if ((auxEjeV + SIZE_SEPARATION*(UBound(myRegistro)+1)) >= MAX_V_PAGINA) then auxEjeV = nuevaPagina()
    dibujarDetalle = auxEjeV + SIZE_SEPARATION
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function dibujarRecuadroDetalle(p_ejeVertical,p_cantReg)    'ESQUELETO DEL DETALLE
    Dim auxEjeV,f,auxposicion
    auxEjeV = p_ejeVertical
    auxposicion = INICIO_SUBTABLAS
    Call GF_squareBox(oPDF, auxEjeV, auxposicion, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_DETALLE, 0, "#DBDBDB", "#DBDBDB", 1, PDF_SQUARE_NORMAL)
    for f = 0 to p_cantReg
        auxEjeV = auxEjeV + SIZE_SEPARATION
        Call GF_squareBox(oPDF, auxEjeV, auxposicion, SIZE_SEPARATION, WIDTH_TABLE_DETALLE, 0, "#F2F5A9", "#F2F5A9", 1, PDF_SQUARE_NORMAL)
    next
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function llenarTituloDetalle(p_ejeVertical)
    Dim auxposicion
    auxposicion = INICIO_SUBTABLAS
    call GF_setFontColor("#517B4A")
    Call GF_setFont(oPDF,"ARIAL",SIZE_SUBTITULOS,FONT_STYLE_BOLD)
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosRubros(0), 130 , PDF_ALIGN_CENTER )
    auxposicion = auxposicion - 130
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosRubros(1), 130 , PDF_ALIGN_CENTER )
    auxposicion = auxposicion - 130
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosRubros(2), 130 , PDF_ALIGN_CENTER )
    llenarTituloDetalle = p_ejeVertical + HEIGHT_BOX_SUBTITULO
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function llenarDetalle(p_ejeVertical,p_arrDet)
    Dim auxposicion
    auxposicion = INICIO_SUBTABLAS
    call GF_setFontColor("#000000")
    Call GF_setFont(oPDF,"ARIAL",SIZE_TEXTO,FONT_STYLE_NORMAL)
    myField = Split(p_arrDet, FIELD_TOKEN)
    For z = 0 To UBound(myField)
        rtrn = Split(myField(z), "=")
        Call loadPropertyRubro(z, rtrn(1))
        if (z < INDEX_DSRUBRO)then
            call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,rtrn(1), 130 , PDF_ALIGN_CENTER )
        end if
        auxposicion = auxposicion - 130
    next
End Function
'------------------------------------------------------------------------------------------------------------------------
'--------------------------------------FUNCIONES RELACIONADAS CON TABLA TOTALES Y PROMEDIOS------------------------------
Function dibujarTotales(p_ejeVertical)      'CUERPO DE TOTALES DE PROMEDIOS
    Dim auxEjeV
    auxEjeV = p_ejeVertical
    Call dibujarRecuadroTotales(auxEjeV,dicTest.count-1)
    auxEjeV = llenarTitulosTotales(auxEjeV)
	for each strKey in dicTest.Keys
        Call llenarDatosTotales(auxEjeV,strKey)
        auxEjeV = auxEjeV + SIZE_SEPARATION
	Next
    dibujarTotales = auxEjeV
End Function
'------------------------------------------------------------------------------------------------------------------------
Function dibujarRecuadroTotales(p_ejeVertical,p_cantRubro)  'ESQUELETO DE TOTALES DE PROMEDIO
    Dim h, auxEjeV,auxposicion
    auxposicion = INICIO_SUBTABLAS
    auxEjeV = p_ejeVertical
    Call GF_squareBox(oPDF, auxEjeV, auxposicion, HEIGHT_BOX_SUBTITULO, WIDTH_TABLE_DETALLE, 0, "#517B4A", "#517B4A", 1, PDF_SQUARE_NORMAL)
    for h = 0 to p_cantRubro
        auxEjeV = auxEjeV + SIZE_SEPARATION
        Call GF_squareBox(oPDF, auxEjeV, auxposicion, SIZE_SEPARATION, WIDTH_TABLE_DETALLE, 0, "#DBDBDB", "#DBDBDB", 1, PDF_SQUARE_NORMAL)
    next
End Function
'------------------------------------------------------------------------------------------------------------------------
Function llenarTitulosTotales(p_ejeVertical)
    Dim auxposicion
    auxposicion = INICIO_SUBTABLAS
    Call GF_setFontColor("#FFFFFF")
    Call GF_setFont(oPDF,"ARIAL",SIZE_SUBTITULOS,FONT_STYLE_BOLD)
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosTotales(0), 130,  PDF_ALIGN_CENTER )
    auxposicion = auxposicion - 130
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosTotales(1), 130,  PDF_ALIGN_CENTER )
    auxposicion = auxposicion - 130
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,arrTitulosTotales(2), 130,  PDF_ALIGN_CENTER )
    llenarTitulosTotales = p_ejeVertical + HEIGHT_BOX_SUBTITULO
End Function
'------------------------------------------------------------------------------------------------------------------------
Function llenarDatosTotales(p_ejeVertical,p_strKey)
    Dim aux , auxposicion
    auxposicion = INICIO_SUBTABLAS
    Call GF_setFontColor("#000000")
    Call GF_setFont(oPDF,"ARIAL",SIZE_TEXT,FONT_STYLE_NORMAL)
    aux = Split(p_strKey,"|")
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,aux(1), 130,PDF_ALIGN_LEFT )
    auxposicion = auxposicion - 130
    call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,aux(2), 130,PDF_ALIGN_LEFT )
    auxposicion = auxposicion - 130
    if Cdbl(dicContRubro.Item(p_strKey))  = 0 then
        call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,"error", 130,PDF_ALIGN_RIGHT )
    else
        call GF_writeVerticalText(oPDF,p_ejeVertical,auxposicion,round(Cdbl(dicTest.Item(p_strKey))/Cdbl(dicContRubro.Item(p_strKey)) ,2) , 130, PDF_ALIGN_RIGHT )
    end if
End Function
'------------------------------------------------------------------------------------------------------------------------
'SE CALCULA EL NETO SIN MERMA POR EL VALOR DEL RUBRO PARA LUEGO SACAR EL PROMEDIO
Function guardarPromedioRubro(p_netoSMerma) 
    Dim auxKey
    auxKey = Trim(g_CdRubro&"|"& Trim(g_AbRubro) &"|"& Trim(g_DsRubro))
    if (not dicTest.Exists(auxKey)) Then
        Call dicTest.Add(auxKey,p_netoSMerma*g_VlRubro)
        Call dicContRubro.Add(auxKey,p_netoSMerma)
    else					
        if (Cdbl(p_netoSMerma) > 0) then 
            dicTest.Item(auxKey) = Cdbl(dicTest.Item(auxKey)) + (p_netoSMerma * g_VlRubro)
            dicContRubro.Item(auxKey) = dicContRubro.Item(auxKey) + Cdbl(p_netoSMerma)
        end if	
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------
'FUNCION QUE SOLO SE UTILIZA CUNADO SE MUESTRA SOLO RESUMEN PROMEDIADO
Function totalizarRubro(p_strDetalle,p_netoSMerma)
    DIM myRegistro,auxKey,myField
    myRegistro = Split(p_strDetalle, DETAIL_TOKEN)
    FOR h = 0 To UBound(myRegistro)
        myField = Split(myRegistro(h), FIELD_TOKEN)
        For z = 0 To UBound(myField)
		    rtrn = Split(myField(z), "=")
		    Call loadPropertyRubro(z, rtrn(1))
        next
        auxKey = Trim(g_CdRubro&"|"& Trim(g_AbRubro) &"|"& Trim(g_DsRubro))
        if (not dicTest.Exists(auxKey)) Then
	        Call dicTest.Add(auxKey,p_netoSMerma*g_VlRubro)
	        Call dicContRubro.Add(auxKey,p_netoSMerma)
        else					
	        if (Cdbl(p_netoSMerma) > 0) then 
		        dicTest.Item(auxKey) = Cdbl(dicTest.Item(auxKey)) + (p_netoSMerma * g_VlRubro)
		        dicContRubro.Item(auxKey) = dicContRubro.Item(auxKey) + Cdbl(p_netoSMerma)
	        end if	
        end if
    NEXT
End Function
'------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------OTRAS FUNCIONES------------------------------------------------------
'EXTRAE EL NETO SIN MERMA DE UN STRING
Function sinMerma(pArr)
    Dim myField, h, rtrn,auxNetoSMerma 	
	auxNetoSMerma = 0
	myField = Split(pArr, FIELD_TOKEN)	
	For h = 0 To UBound(myField)
		rtrn = Split(myField(h), "=")		
		if (h = INDEX_NETO_SIN_MERMA) then auxNetoSMerma = rtrn(1)
	Next
	sinMerma = auxNetoSMerma
End Function
'-------------------------------------------------------------------------------------------------------------
Function loadPropertyRubro(pIndex,pVl)
	Select Case pIndex
		Case INDEX_CDRUBRO
			g_CdRubro = pVl
		Case INDEX_ABRUBRO
			g_AbRubro = pVl
		Case INDEX_VLRUBRO
			g_VlRubro = pVl
		Case INDEX_DSRUBRO
			g_DsRubro = pVl
	End Select		
End Function
'-------------------------------------------------------------------------------------------------------------
function nuevaPagina()
    dim eje_Vertical
	Call GF_newPage(oPDF)
	Call PDFGirarHoja(ANGLE_RECTO)
    nroPagina = nroPagina + 1
    call dibujarPagina()
    nuevaPagina = INICIO_EJE_V
end function
'-------------------------------------------------------------------------------------------------------------
'DIBUJARA RECUADRO DE MARGEN + NUMERACION DE PAGINA        
Function dibujarPagina()
Dim posicion_V
Call GF_squareBox(oPDF, 2, 3, MAX_V_PAGINA, MAX_H_PAGINA , 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)'RCUADROS
Call GF_setFont(oPDF,"ARIAL",SIZE_TEXTO,FONT_STYLE_NORMAL)
Call GF_setFontColor("#000000")
call GF_writeVerticalText(oPDF, MAX_V_PAGINA + 10 , MAX_H_PAGINA, "Pagina " & nroPagina, MAX_H_PAGINA-5 , PDF_ALIGN_RIGHT)'NUMERACION DE PAGINA
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'   ETAPA 2 PDF - GENERACION DEL PDF
'*****************************************************************************************

Dim str,index,txtLine,arrData, flagHayResultado,fadm,auxCdRubro,auxDsRubro,auxAbRubro,dicContRubro,totNetoSinMerma

g_FechaDesde = g_FechaDesdeA & "-" & g_FechaDesdeM & "-" & g_FechaDesdeD
g_FechaHasta = g_FechaHastaA & "-" & g_FechaHastaM & "-" & g_FechaHastaD

fname = "REPORTE_CALADOR_" & g_Pto
index = 0
Set dicTest = Server.CreateObject("Scripting.Dictionary")
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set dicCam = Server.CreateObject("Scripting.Dictionary")
Set dicVag = Server.CreateObject("Scripting.Dictionary")
Set dicContRubro = Server.CreateObject("Scripting.Dictionary")
while index <= maxSegment
	pStrPath = Server.MapPath("Temp/REPORTE_CALADOR_" & session("Usuario") & "_" & index & ".txt")
    if (fs.FileExists(pStrPath)) then
		Set fadm = fs.OpenTextFile(pStrPath, 1)
		while (not fadm.AtEndOfStream)
			txtLine = fadm.ReadLine()
			if (Trim(txtLine) = REPORTE_CAMIONES) then
				isCamion = true
				isVagon  = false
			else if (Trim(txtLine) = REPORTE_VAGONES) then
					isCamion = false
					isVagon  = true
				else
					if isCamion then Call dicCam.add(cont,txtLine)
					if isVagon then  Call dicVag.add(cont,txtLine)
				end if
			end if
			cont = cont + 1
		wend
		Set fadm = nothing
		fs.DeleteFile(pStrPath)
	end if
	index = index + 1
wend

nroPagina = 1
Set oPDF = GF_createPDF(fname)
Call PDFGirarHoja(ANGLE_RECTO)
Call GF_setPDFMODE(PDF_STREAM_MODE)
Call armadoPDF()
Call GF_closePDF(oPDF)

%>