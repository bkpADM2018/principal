<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->


<%
Const P_Y_BODY = 225
Const P_Y_FINALLY = 700
Const SEPARATION = 10
'*****************************************************************************************************************************
'*****************************************************************************************************************************
Function armadoPDF(p_Minuta, p_Evento, p_Fecha, p_TipoCbte)
    call dibujarEncabezado(p_Minuta)
    Set rsCab = getCabeceraMinuta(p_Minuta)
    if (not rsCab.Eof) then
        call dibujarCabecera(rsCab,p_Minuta)
        pY = P_Y_BODY
        set rsDetMin = getDetalleCuenta(p_Minuta,rsCab("iddivision"))
        if (not rsDetMin.Eof) then
             Call procesarDetalleCuenta(rsDetMin)
             pY = pY + SEPARATION
        end if
        set rsDetPic = getDetalleCotizaciones(p_Minuta,rsCab("moneda"))
        if (not rsDetPic.Eof) then
             Call procesarDetalleCotizaciones(rsDetPic)
            pY = pY + SEPARATION
        end if
        Call dibujarTotalMinuta(rsCab, p_Minuta)
        Call procesarMinutaTipoDeCambio(p_Minuta, p_Fecha, rsCab("fechaComprobante"), rsCab("tcCbte"),p_Evento)
        Call dibujarFirmasMinuta(rsCab("ususarioCarga"),rsCab("fechacarga"),rsCab("horacarga"))
        Call dibujarTotalNroPagina()
    end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTotalNroPagina()
    'Se graba la cantidad total de cupos.
    for i = 1 to nroPagina
        Call setWorkPage(oPDF, i)
        call GF_setFont(oPDF,"ARIAL",8,0)
	    Call GF_writeTextAlign(oPDF, 10 , 840, " de " & i , 575 , PDF_ALIGN_RIGHT)
    next
End Function 
'-----------------------------------------------------------------------------------------------------------------------------
'Procesa y verifica la informacion de Tipo de cambio de la minuta en caso que tenga
Function procesarMinutaTipoDeCambio(p_Minuta, p_Fecha, p_FechaCbte, p_TCCbte,p_Evento)
    Dim strSQL, rsTC, pX
    strSQL = " SELECT A.MINUSA, "&_
             "        A.FECHSA, "&_
             "        CASE WHEN B.CDUSUARIO IS NULL THEN '' ELSE B.CDUSUARIO END AS CDUSUARIO, "&_
             "        CASE WHEN B.FECHAFIRMA IS NULL THEN '' ELSE B.FECHAFIRMA END AS FECHAFIRMA, "&_
             "        CASE WHEN B.HKEY IS NULL THEN '' ELSE B.HKEY END AS HKEY "&_
             " FROM MERFL.MER301F1 A "&_
             " LEFT JOIN MERFL.TBLMINUTASFIRMAS B ON A.FECHSA = B.FECHA AND A.MINUSA = B.MINUTA AND B.TIPODOCUMENTO = '"& p_Evento &"'"&_
             " WHERE A.MINUSA="& p_Minuta &" AND A.FECHSA="& p_Fecha &" AND A.EVENSA='"& AUTH_TYPE_PICC &"' AND A.FORMSA='"& PREFIX_FAC &"'"&_
             " ORDER BY B.SECUENCIA"
    Call executeQuery(rsTC, "OPEN", strSQL)
    if (not rsTC.Eof) then
        Call GF_squareBox(oPDF, 15, py, 560, 60, 0, "", "#0B3B0B", 1, PDF_SQUARE_ROUND)
        pY = pY + 5
        Call GF_writeTextAlign(oPDF,20, pY,  "La factura posee un tipo de cambio que no se corresponde con el de su día de emisión" , 500, PDF_ALIGN_LEFT)
        pY = pY + SEPARATION
        Call GF_writeTextAlign(oPDF,20, pY,  "Tipo de cambio factura: $ " & GF_EDIT_DECIMALS(Cdbl(p_TCCbte)*100,2), 260, PDF_ALIGN_CENTER)
        auxTCSistema = getTipoCambioCV(MONEDA_DOLAR, p_FechaCbte, T_CAMBIO_VENDEDOR)    
        Call GF_writeTextAlign(oPDF,280, pY,  "Tipo de cambio del día: $ " & GF_EDIT_DECIMALS(Cdbl(auxTCSistema)*100,2), 260, PDF_ALIGN_CENTER)
        pY = pY + SEPARATION
        Call GF_writeTextAlign(oPDF,20, pY,  "Autorizan:", 260, PDF_ALIGN_LEFT)
        pY = pY + SEPARATION
        while (not rsTC.Eof)
            if (rsTC("CDUSUARIO") <> "") then Call GF_writeTextAlign(oPDF,20, pY, rsTC("CDUSUARIO") & " - " & getUserDescription(rsTC("CDUSUARIO")) & " - " & armarTextoPlanoFirma(rsTC("HKEY"), rsTC("FECHAFIRMA")), 500, PDF_ALIGN_LEFT)
            pY = pY + SEPARATION
            rsTC.MoveNext()
        wend
        call GF_setFont(oPDF,"COURIER",8,0) 
        pY = pY + SEPARATION*2
    end if
    if (Cint(pY) > P_Y_FINALLY) then Call nuevaHojaMinuta()
End Function
'-----------------------------------------------------------------------------------------------------------------------------
'esta funcion solo dibuja el lugar de firmas para que se firmen a mano
Function dibujarFirmasMinuta(p_User, p_FechaCarga, p_HoraCrga)
    call GF_squareBox(oPDF, 15, py, 560, 100, 0, "", "#0B3B0B", 1, PDF_SQUARE_ROUND)
    pY = pY + 5
    Call GF_writeTextAlign(oPDF,20, pY,  "Comprobante procesado por: " & getUserDescription(p_User) , 100, PDF_ALIGN_LEFT)
    pY = pY + 8
    Call GF_writeTextAlign(oPDF,20, pY, GF_TRADUCIR("Fecha y hora de carga: ") & GF_FN2DTE(p_FechaCarga) &" "&Left(p_HoraCrga,2)&":"&Mid(p_HoraCrga,3,2)&":"&Right(p_HoraCrga,2), 565, PDF_ALIGN_LEFT)
    pY = pY + 14
    call GF_horizontalLine(oPDF, 15, py, 560)
    Call GF_writeTextAlign(oPDF,20, pY,  "Autorización Sector" , 185, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,200, pY,  "Autorización 2" , 185, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,385, pY,  "Controller/Compliance de corresponder" , 185, PDF_ALIGN_CENTER)
    Call GF_verticalLine(oPDF, 201, pY, 73)
    Call GF_verticalLine(oPDF, 387, pY, 73)
    pY = pY + SEPARATION
    call GF_horizontalLine(oPDF, 15, py, 560)
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTotalMinuta(p_Rs, p_Minuta)
    Dim rsTot
    strSQL = "select * from tesfl.tes111f1 "&_
             " where APNING = "& p_Minuta &_
 	         "  and APTCBT = '"& p_Rs("tipoCbte") &"'"&_
 	         "  and APFPAG > "& p_Rs("fechacarga") 
    Call executeQuery(rsTot, "OPEN", strSQL)
    if (not rsTot.Eof) then
        Call GF_writeTextAlign(oPDF,15, pY,  GF_TRADUCIR("FECHA") , 100, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,115, pY,  GF_TRADUCIR("IMPORTE") , 100, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,215, pY,  GF_TRADUCIR("EDO") , 60, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,275, pY,  GF_TRADUCIR("ORDEN P/C") , 80, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,355, pY,  GF_TRADUCIR("BCO") , 80, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,435, pY,  GF_TRADUCIR("SUC") , 80, PDF_ALIGN_CENTER)
        pY = pY + SEPARATION
        Call GF_horizontalLine(oPDF, 15, pY, 560)
        pY = pY + 5
        Call GF_writeTextAlign(oPDF,15, pY,  GF_FN2DTE(rsTot("APFPAG")) , 100, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,115, pY,  GF_EDIT_DECIMALS(cDbl(rsTot("APICBT"))*100,2) , 100, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,215, pY,  rsTot("APESTA") , 60, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,275, pY,  rsTot("APORPC") , 80, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,355, pY,  rsTot("APCBCO") , 80, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,435, pY,  rsTot("APSBCO") , 80, PDF_ALIGN_CENTER)
    end if
    pY = pY + SEPARATION*2
    if (Cint(pY) > P_Y_FINALLY) then Call nuevaHojaMinuta()
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function nuevaHojaMinuta()
    Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1
	Call dibujarEncabezado()
    'el recordset de cabecera de minuta no se mueve al proximo registro debido a que posee un unico registro y se utiliza para nuevas paginas
    call dibujarCabecera(rsCab,minuta)
	pY = P_Y_BODY
End function
'-----------------------------------------------------------------------------------------------------------------------------
Function procesarDetalleCotizaciones(p_RsDet)
    Call dibujarTituloDetalleCotizaciones()
    while not p_RsDet.eof
        Call dibujarCuerpoDetalleCotizaciones(p_RsDet)
        if (Cint(pY) > P_Y_FINALLY) then
            Call nuevaHojaMinuta()
            Call dibujarTituloDetalleCotizaciones()
        end if
        p_RsDet.MoveNext()
    wend
End Function 
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTituloDetalleCotizaciones()
    Call GF_writeTextAlign(oPDF,15, pY,  GF_TRADUCIR("COTIZACIÓN") , 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,75, pY, GF_TRADUCIR("CANTIDAD") , 45, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,120, pY, GF_TRADUCIR("ARTÍCULO") , 150, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,270, pY, GF_TRADUCIR("P.PRESUP") , 80, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,350, pY, GF_TRADUCIR("PRECIO") , 80, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,430, pY, GF_TRADUCIR("IMPORTE") , 80, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,510, pY, GF_TRADUCIR("TASA") , 60, PDF_ALIGN_CENTER)
    pY = pY + 10
    Call GF_horizontalLine(oPDF, 15, pY, 560)
    pY = pY + 5
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCuerpoDetalleCotizaciones(p_Rs)
    Call GF_writeTextAlign(oPDF,15, pY,  p_Rs("IDCOTIZACION") , 60, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,75, pY,  GF_EDIT_DECIMALS(cDbl(p_Rs("CANTIDAD"))*1000,0) , 40, PDF_ALIGN_RIGHT)
    auxDsArticulo = Trim(p_Rs("DSARTICULO"))
    if (Len(auxDsArticulo) > 32) then auxDsArticulo = Left(auxDsArticulo,30) & ".."
    Call GF_writeTextAlign(oPDF,120, pY, auxDsArticulo, 150, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,270, pY, p_Rs("AREA") &" "& p_Rs("DETALLE") , 80, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,350, pY, GF_EDIT_DECIMALS(cDbl(p_Rs("PRECIO"))*100,2) , 80, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(oPDF,430, pY, GF_EDIT_DECIMALS(cDbl(p_Rs("IMPORTE"))*100,2) , 80, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(oPDF,510, pY, GF_EDIT_DECIMALS(cDbl(p_Rs("TASA"))*100,2) , 40, PDF_ALIGN_RIGHT)
    pY = pY + 8
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function getDetalleCotizaciones(p_Minuta,p_Moneda)
    Dim strSQL,rs
    
    strSQL = "select acd.cotix7 as idcotizacion,"&_
             "       acd.cantx7 as cantidad, "&_
             "       acd.artix7 as idarticulo,"&_
             "       art.dsarticulo,"&_
             "       acd.areax7 as area,"&_
             "       acd.detax7 as detalle,"&_
             "       acd.tivax7 as tasa, "
    if (p_Moneda = MONEDA_PESO) then
        strSQL = strSQL & " precx7 as precio,impox7 as importe "
    else
        strSQL = strSQL & " predx7 as precio,impdx7 as importe "
    end if
    strSQL = strSQL & " from provfl.acd7rep acd "&_
    "       inner join toepferdb.tblarticulos art on art.idarticulo = acd.artix7 "&_
    "       where ningx7 = "& p_Minuta
    Call executeQuery(rs, "OPEN", strSQL)
    Set getDetalleCotizaciones = rs
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function procesarDetalleCuenta(p_RsDet)
    Call dibujarTituloDetalleCuenta()
    while not p_RsDet.eof
        Call dibujarCuerpoDetalleCuenta(p_RsDet)
        if (Cint(pY) > P_Y_FINALLY) then
            Call nuevaHojaMinuta()
            Call dibujarTituloDetalleCuenta()
        end if
        p_RsDet.MoveNext()
    wend
End function
'-----------------------------------------------------------------------------------------------------------------------------
Function getDetalleCuenta(p_Minuta,p_idDivision)
    dim strSQL,rs,myNombreMiembro, auxDivision
    'Por medio de la division obtengo el miembro de la tabla C50011F1
    auxDivision = getDivisionAbreviada(p_idDivision)
    select case CStr(auxDivision)
	    case CODIGO_EXPORTACION
		    myNombreMiembro = "M01"
	    case CODIGO_ARROYO
    		myNombreMiembro = "M09"	
    	case CODIGO_PIEDRABUENA
		    myNombreMiembro = "M10"
	    case CODIGO_TRANSITO 
    		myNombreMiembro = "M07"
    end select 
    'Creo el alias a la tabla C50011F1 para su respecteivo miembro segun la division
    if not executeQuery(rs,"EXISTS","C50011F1" & myNombreMiembro) then
        Call executeQuery(rs,"EXEC","CREATE ALIAS MERFL.C50011F1" & myNombreMiembro & " FOR CGFL.C50011F1("& myNombreMiembro &")")
    end if
    'Obtengo los datos de la cuenta de costo de la minuta
    strSQL = " select DWBTCD as cuenta,"&_
             "        DWCTOS as cc,"&_
             "        DWD9ST as dc, "&_
             "        DWQ6NB as pesos, "&_
             "        DWQ7NB as dolares,"&_
             "        case when b.PNOMCT is null then '' else b.PNOMCT end as descripcion "&_
             " from provfl.acdwrep A "&_
             "  left join MERFL.C50011F1"&myNombreMiembro&" B on b.PCDCTA = a.DWBTCD and b.PCCTOS = A.DWCTOS " &_
             " where a.dwqfnb = "&p_Minuta &_
             " order by DWBTCD,DWCTOS "
    Call executeQuery(rs, "OPEN", strSQL)
    'Borro el alias creado
    Call executeQuery(rs, "EXEC", "DROP ALIAS MERFL.C50011F1" & myNombreMiembro)
    Set getDetalleCuenta = rs
End function 
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCuerpoDetalleCuenta(p_Rs)
    Dim auxDescripcion
    Call GF_writeTextAlign(oPDF,15, pY, Trim(p_Rs("cuenta")) , 90, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,105, pY, Trim(p_Rs("cc")) , 40, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,145, pY, p_Rs("dc") , 40, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,185, pY, GF_EDIT_DECIMALS(cDbl(p_Rs("pesos"))*100,2) , 90, PDF_ALIGN_RIGHT)
    Call GF_writeTextAlign(oPDF,285, pY, GF_EDIT_DECIMALS(cDbl(p_Rs("dolares"))*100,2) , 90, PDF_ALIGN_RIGHT)
    auxDescripcion = Trim(p_Rs("descripcion"))
    if (Len(auxDescripcion) > 40) then auxDescripcion = Left(auxDescripcion,38)& ".."
    Call GF_writeTextAlign(oPDF,385, pY, auxDescripcion , 200, PDF_ALIGN_LEFT)
    pY = pY + 8
End Function 
'-----------------------------------------------------------------------------------------------------------------------------
Function getCabeceraMinuta(p_Minuta)
    Dim rs
    strSQL = " select A.*,B.dsdivision,B.iddivision,c.*,D.D1HCTX as dsSector,E.dsEmpresa as dsProveedor,F.dsEmpresa as dsProvRetenciones,G.CBIOX9 as tcCbte "&_
             "from ("&_
             "select DSC8ST as cddivision, "&_
	         "     dsdust as tipoCbte, "&_
	         "     dsqgnb as factura, "&_
	         "     '20' || SUBSTR(dscqdt,2,6) fechapago, "&_
	         "     DSOPR1 as idproveedor, "&_
	         "     DSQSNB as idProveedorretenciones, "&_
	         "     DSD1ST as concepto, "&_
	         "     DSDWST as moneda, "&_
	         "     DSTIGA as retenciones, "&_
	         "     DSDYST as ingBrutos, "&_
	         "     DSGYTX as ususarioCarga, "&_
             "     '20' || SUBSTR(DSCSDT,2,6) fechacarga, "&_
	         "     DSATTM as horacarga, "&_
             "     DSGXTX as observaciones, "&_
             "     DSQMNB + DSQNNB as bruto, "&_
             "     DSQJNB as neto, "&_
             "     DSQHNB as iva,"&_
             "     DSDIPO, "&_
             "     dsbgcd, "&_
             "     '20' || SUBSTR(DSCRDT,2,6) fechaComprobante, "&_
             "     dsqfnb as minuta "&_
             "from provfl.acdsrep where dsqfnb = "&p_Minuta &" ) A "&_
             " left join TOEPFERDB.TBLDIVISIONES B on b.CDDIVISIONABR = a.cddivision "&_
             " left join provfl.acdtrel0 C on C.DTD1ST = a.concepto "&_
             " left join provfl.acd1rep D on D.D1BGCD = A.DSBGCD "&_
             " left join toepferdb.vwempresas E on E.idempresa = A.idproveedor"&_
             " left join toepferdb.vwempresas F on F.idempresa = A.idProveedorretenciones "&_
             " left join provfl.acd9rep G on G.ningx9 = A.minuta "
    Call executeQuery(rs, "OPEN", strSQL)
    Set getCabeceraMinuta = rs
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarEncabezado(p_Minuta)
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,2, 25, GF_TRADUCIR("MINUTA DE CARGA"), 590, PDF_ALIGN_CENTER)
    call GF_setFont(oPDF,"ARIAL",14,8)
    Call GF_writeTextAlign(oPDF,400, 45, GF_TRADUCIR("N° ") & p_Minuta, 180, PDF_ALIGN_RIGHT)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Página  " & nroPagina		 , 558 , PDF_ALIGN_RIGHT)
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+SEPARATION,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
end Function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarCabecera(p_rsCab,p_Minuta)
    Dim auxObservaciones,auxTCMinuta
    'Obtenemos el tipo de cambio de la fecha que se cargo la minuta
    auxTCMinuta = getTipoCambioCV(MONEDA_DOLAR, p_rsCab("fechacarga"), T_CAMBIO_VENDEDOR)
    call GF_setFont(oPDF,"COURIER",8,0)
    Call GF_writeTextAlign(oPDF,15, 72, GF_TRADUCIR("Comprobante del proveedor:"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 82, GF_TRADUCIR("Tipo: ") & p_rsCab("tipoCbte") , 80, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,100, 82, GF_TRADUCIR("Nro.: ") & GF_EDIT_CBTE(p_rsCab("factura")) , 130, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,230, 82, GF_TRADUCIR("Fecha: ") & GF_FN2DTE(p_rsCab("fechaComprobante")) , 115, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,345, 82, GF_TRADUCIR("T.C: ") & GF_EDIT_DECIMALS(Cdbl(p_rsCab("tcCbte"))*100,2) , 100, PDF_ALIGN_LEFT)
    
    Call GF_writeTextAlign(oPDF,15, 102, GF_TRADUCIR("División...........: ") & Trim(p_rsCab("dsdivision")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 112, GF_TRADUCIR("Sector.............: ") & Trim(p_rsCab("dsSector")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 122, GF_TRADUCIR("Fecha de Pago......: ") & GF_FN2DTE(p_rsCab("fechapago")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 132, GF_TRADUCIR("Proveedor..........: ") & p_rsCab("idproveedor")&"-"&p_rsCab("dsProveedor"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 142, GF_TRADUCIR("Retenciones a......: ") & p_rsCab("idProveedorretenciones")&"-"&p_rsCab("dsProvRetenciones"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 152, GF_TRADUCIR("Concepto...........: ") & p_rsCab("concepto") &" "& Trim(p_rsCab("DTG0TX")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 162, GF_TRADUCIR("Moneda pago........: ") & getNombreMoneda(p_rsCab("moneda")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,265, 162, GF_TRADUCIR("T.C Minuta: ") & GF_EDIT_DECIMALS(auxTCMinuta*100,2) , 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 172, GF_TRADUCIR("Retenciones: IGA...: ") & p_rsCab("DTGAAP"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,150, 172, GF_TRADUCIR("Ing.B.: ") & p_rsCab("DTBRAP"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,265, 172, GF_TRADUCIR("I.V.A.: ") & p_rsCab("DTGCST"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,450, 172, GF_TRADUCIR("Dest.: ") & p_rsCab("DSDIPO"), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,15, 182, GF_TRADUCIR("Ingresado por......: ") &  getUserDescription(p_rsCab("ususarioCarga")), 285, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,265, 182, GF_TRADUCIR("Fecha carga: ") & GF_FN2DTE(p_rsCab("fechacarga")), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,450, 182, GF_TRADUCIR("Hora: ") & Left(p_rsCab("horacarga"),2)&":"&Mid(p_rsCab("horacarga"),3,2)&":"&Right(p_rsCab("horacarga"),2), 565, PDF_ALIGN_LEFT)
    auxObservaciones = Trim(p_rsCab("observaciones"))
    if (Len(auxObservaciones) > 95) then auxObservaciones = Left(auxObservaciones,93) &".."
    Call GF_writeTextAlign(oPDF,15, 192, GF_TRADUCIR("Observaciones......: ") & auxObservaciones, 565, PDF_ALIGN_LEFT)
    
    Call GF_writeTextAlign(oPDF,15, 202, GF_TRADUCIR("Imp. Bruto.........: ") & GF_EDIT_DECIMALS(cDbl(p_rsCab("bruto"))*100,2), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,265, 202, GF_TRADUCIR("IVA: ") & GF_EDIT_DECIMALS(cDbl(p_rsCab("iva"))*100,2), 565, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF,450, 202, GF_TRADUCIR("Imp. Neto: ") & GF_EDIT_DECIMALS(cDbl(p_rsCab("neto"))*100,2), 565, PDF_ALIGN_LEFT)

    Call drawCodeBar(p_Minuta,500,80,35)

    Call GF_horizontalLine(oPDF,2,216,590)

End function
'-----------------------------------------------------------------------------------------------------------------------------
Function dibujarTituloDetalleCuenta()
    Call GF_writeTextAlign(oPDF,15, pY, GF_TRADUCIR("CUENTA/C.") , 90, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,105, pY, GF_TRADUCIR("COSTOS") , 40, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,145, pY, GF_TRADUCIR("DC") , 40, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,185, pY, GF_TRADUCIR("PESOS") , 100, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,285, pY, GF_TRADUCIR("DOLARES") , 100, PDF_ALIGN_CENTER)
    Call GF_writeTextAlign(oPDF,385, pY, GF_TRADUCIR("DESCRIPCIÓN") , 200, PDF_ALIGN_CENTER)
    pY = pY + 10
    Call GF_horizontalLine(oPDF, 15, pY, 560)
    pY = pY + 5
End Function

'****************************************************************************************************************************
'********************************	             COMIENZO DE LA PAGINA              ********************************
'***********************************************************************************************************************************
Dim minuta,evento,fecha,tipoCbte,oPDF,rsCab,rsDetMin,rsDetPic,nroPagina,pY



minuta = GF_Parametros7("minuta","",6)
evento = GF_Parametros7("evento","",6)    
fecha = GF_Parametros7("fecha","",6)
tipoCbte = GF_Parametros7("tipoCbte","",6)
nroPagina = 1
Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF(minuta,evento,fecha,tipoCbte)
Call GF_closePDF(oPDF)

'*********************************************


%>
