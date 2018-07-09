<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientosRoles.asp"-->
<%
Const SEPARACION_Y = 12
Const PAGE_HEIGHT_SIZE = 800
'----------------------------------------------------------------------------------------------------------------------	
Function corteControlFechaMermaVolatil(p_Rs,p_Fecha)
    corteControlFechaMermaVolatil = false
    if (not p_Rs.Eof) then
        if (Cdbl(p_Fecha) = Cdbl(p_Rs("FECHADESDE"))) then corteControlFechaMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function corteControlProdcutoMermaVolatil(p_Rs,p_Producto,p_Fecha)
    corteControlProdcutoMermaVolatil = false
    if (not p_Rs.Eof) then
        if ((Cdbl(p_Producto) = Cdbl(p_Rs("CDPRODUCTO")))and(Cdbl(p_Fecha) = Cdbl(p_Rs("FECHADESDE")))) then corteControlProdcutoMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function corteControlClienteMermaVolatil(p_Rs,p_Cliente,p_Fecha,p_Producto)
    corteControlClienteMermaVolatil = false
    if (not p_Rs.Eof) then
        if ((Cdbl(p_Cliente) = Cdbl(p_Rs("IDORIGEN")))and(Cdbl(p_Fecha) = Cdbl(p_Rs("FECHADESDE")))and(Cdbl(p_Producto) = Cdbl(p_Rs("CDPRODUCTO")))) then corteControlClienteMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Esta funcion recibe como parametro el puerto y el recordset cargado con todos los ajustes que tiene el periodo
Function writeObservationsMermaVolatil(pto, ByRef pRs, ByRef py)
    Dim auxFecha,auxCdProducto,auxCdCliente,auxDsCliente,totalKilos,totalKilosFecha,totalKilosProducto,auxDsProducto
    totalKilosFecha = 0
    py = py + 12
    while not pRs.Eof
        'La fecha desde y hasta en la tabla de ajuste para merma volatil son iguales, de esta manera se toma cualquiera de las dos
        auxFecha = pRs("FECHADESDE")
        while (corteControlFechaMermaVolatil(pRs, auxFecha))
            auxCdProducto = pRs("CDPRODUCTO")
            auxDsProducto = Trim(pRs("DSPRODUCTO"))
            Call GF_writeTextAlign(oPDF, 60, py,GF_TRADUCIR("Producto: " & auxDsProducto), 200, PDF_ALIGN_LEFT)
            totalKilosProducto = 0
            while (corteControlProdcutoMermaVolatil(pRs, auxCdProducto,auxFecha))
                auxCdCliente = Cdbl(pRs("IDORIGEN"))
                totalKilos = 0
                while (corteControlClienteMermaVolatil(pRs, auxCdCliente, auxFecha, auxCdProducto))
                    totalKilos = Cdbl(totalKilos) + ABS(Cdbl(pRs("KILOSAJUSTE")))
                    pRs.MoveNext()
                wend
                py = py + 12
                if (CDbl(py) > PAGE_HEIGHT_SIZE) then py = nuevaHoja()
                Call GF_writeTextAlign(oPDF, 90, py, getDsCliente(auxCdCliente) , 200, PDF_ALIGN_LEFT)
                Call GF_writeTextAlign(oPDF, 250, py, GF_EDIT_DECIMALS(Cdbl(totalKilos)*100,2) & " KG", 50, PDF_ALIGN_RIGHT)
                totalKilosProducto = Cdbl(totalKilosProducto) + Cdbl(totalKilos)
            wend
            py = py + 12
            Call GF_setFont(oPDF,"ARIAL",8,8)
            Call GF_writeTextAlign(oPDF, 60, py,GF_TRADUCIR("Total "& auxDsProducto), 200, PDF_ALIGN_LEFT)
            Call GF_writeTextAlign(oPDF, 250, py, GF_EDIT_DECIMALS(Cdbl(totalKilosProducto)*100,2) & " KG", 50, PDF_ALIGN_RIGHT)
            totalKilosFecha = Cdbl(totalKilosFecha) + Cdbl(totalKilosProducto)
            Call GF_setFont(oPDF,"ARIAL",8,0)
            py = py + 20
            if (CDbl(py) > PAGE_HEIGHT_SIZE) then py = nuevaHoja()
        wend
    wend
    Call GF_setFont(oPDF,"ARIAL",8,8)
    Call GF_writeTextAlign(oPDF, 60, py,GF_TRADUCIR("Total Periodo"), 200, PDF_ALIGN_LEFT)
    Call GF_writeTextAlign(oPDF, 250, py, GF_EDIT_DECIMALS(Cdbl(totalKilosFecha)*100,2) & " KG", 50, PDF_ALIGN_RIGHT)
    Call GF_setFont(oPDF,"ARIAL",8,0)
    py = py + 20
    Call GF_writeTextAlign(oPDF, 260, py,GF_TRADUCIR("Fin del reporte"), 70, PDF_ALIGN_CENTER)
    py = py + 5
    Call GF_horizontalLine(oPDF, 60, py, 200)
    Call GF_horizontalLine(oPDF, 330, py, 200)

End function
'----------------------------------------------------------------------------------------------------------------------	
function nuevaHoja()
	Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1
    Call drawReportFormat()
    nuevaHoja = 80
end function
'----------------------------------------------------------------------------------------------------------------------	
'escribe las observaciones de un Draft Survey
Function writeObservationsDraft(pto, pIdOrigen, ByRef py)
	Dim msg,rs, diff, kg3ros, tipoDiffDraft, tipoDiff3ros, auxKgBzaToepfer
	strSQL = "SELECT A.*, B.CDBUQUE, C.DSBUQUE " &_
			 "FROM (SELECT * " &_
			 "		FROM TBLEMBARQUESDRAFTSURVEY WHERE IDDRAFT = " & pIdOrigen &") A " &_
			 "	 INNER JOIN " & _
			 "   (  (Select getDate() DTCONTABLE, CDAVISO, NUOPERACION, CDBUQUE, CDEMPRESAEMBARQUE, VLCALADOINICIAL, VLCALADOFINAL, ICFUMIGACION, CDEMPRESAFUMIG, ICAGUA, DSOBSERVACIONES, DTAVISO, CDCOORDINADORA, ICESTADOLIQ from EMBARQUES) " &_
	         "          union " &_
	         "      (Select DTCONTABLE, CDAVISO, NUOPERACION, CDBUQUE, CDEMPRESAEMBARQUE, VLCALADOINICIAL, VLCALADOFINAL, ICFUMIGACION, CDEMPRESAFUMIG, ICAGUA, DSOBSERVACIONES, DTAVISO, CDCOORDINADORA, ICESTADOLIQ from HEMBARQUES)) B  ON A.CDAVISO = B.CDAVISO " &_
			 "	 INNER JOIN BUQUES C ON C.CDBUQUE = B.CDBUQUE "			 

    call GF_BD_Puertos (pto, rs, "OPEN",strSQL)
	if not rs.EoF then
		diff = CDbl(rs("TOTALDRAFT")) - CDbl(rs("TOTALBALANZA"))
		tipoDiffDraft = "Faltante"
	    tipoDiff3ros = "Sobrante"
	    if (diff > 0) then
	        tipoDiffDraft = "Sobrante"
	        tipoDiff3ros = "Faltante"
	    end if
	    
		Call GF_writeTextAlign(oPDF, 115, py,GF_TRADUCIR("Código y Nombre del Buque.......: " & rs("CDBUQUE") &" - "& rs("DSBUQUE")), 200, PDF_ALIGN_LEFT)
		py = py + SEPARACION_Y
		Call GF_writeTextAlign(oPDF, 115, py,GF_TRADUCIR("Código de Aviso de Embarque...: " & rs("CDAVISO")), 200, PDF_ALIGN_LEFT)
		py = py + SEPARACION_Y
		Call GF_setFont(oPDF,"ARIAL",8,8)			
		Call GF_writeTextAlign(oPDF, 115, py,GF_TRADUCIR("DATOS DE BALANZA"), 200, PDF_ALIGN_LEFT)		
		auxKgBzaToepfer = 0
		if (not isNull(rs("KGBZATOEPFER"))) then auxKgBzaToepfer = Cdbl(rs("KGBZATOEPFER"))		
		py = py + SEPARACION_Y
		Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR("Cargas de ADM Agro:"), 100, PDF_ALIGN_RIGHT)
		Call GF_setFont(oPDF,"ARIAL",8,0)			
	    Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(auxKgBzaToepfer,0) &" Kg.", 100, PDF_ALIGN_RIGHT)
		Call GF_setFont(oPDF,"ARIAL",8,8)
		kg3ros = CDbl(rs("TOTALBALANZA")) - auxKgBzaToepfer
		py = py + SEPARACION_Y
		Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR("Cargas de 3ros:"), 100, PDF_ALIGN_RIGHT)
		Call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(kg3ros,0) &" Kg.", 100, PDF_ALIGN_RIGHT)			  
		Call GF_setFont(oPDF,"ARIAL",8,8)
		py = py + SEPARACION_Y
		Call GF_horizontalLine(oPDF, 240, py,100)
		py = py + 4
		Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR("Carga Total Balanza:"), 100, PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(CDbl(rs("TOTALBALANZA")),0) &" Kg.", 100, PDF_ALIGN_RIGHT)
		py = py + (SEPARACION_Y * 2)
		Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR("Resultado de Draft Survey:"), 100, PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(CDbl(rs("TOTALDRAFT")),0) &" Kg.", 100, PDF_ALIGN_RIGHT)
		py = py + (SEPARACION_Y * 2)
		Call GF_writeTextAlign(oPDF, 115, py,GF_TRADUCIR("RESULTADO"), 100, PDF_ALIGN_LEFT)
		py = py + SEPARACION_Y
		Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR(tipoDiffDraft &" x D.Survey:"), 100, PDF_ALIGN_RIGHT)
		Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(diff,0) &" Kg.", 100, PDF_ALIGN_RIGHT)		
		if (kg3ros <> 0) then
			  py = py + SEPARACION_Y
              diff3ros = -diff * (CDbl(rs("TOTALBALANZA")) - auxKgBzaToepfer)/CDbl(rs("TOTALBALANZA"))
              Call GF_writeTextAlign(oPDF, 135, py,GF_TRADUCIR(tipoDiff3ros & " x Cargas de 3ros:"), 100, PDF_ALIGN_RIGHT)
			  Call GF_writeTextAlign(oPDF, 235, py,GF_EDIT_DECIMALS(diff3ros,0) &" Kg.", 100, PDF_ALIGN_RIGHT)
        end if
        Call GF_setFont(oPDF,"ARIAL",8,0)
	end if
End Function
'------------------------------------------------------------------------------------------------------
'Function writeSignature: escribe las firmas que se registraron hasta el momento del Ajuste
Function writeSignature(pto, idAjust ) 
	Dim strSQL, rs,cdGerente, dsGerente, firmaGerente, cdController, dsController, firmaController, cdDirector, dsDirector, firmaDirector
	strSQL = "Select * from TBLAJUSTESFIRMAS where IDAJUSTE=" & idAjuste & " order by SECUENCIA"
	call GF_BD_Puertos (pto, rs, "OPEN",strSQL)
	if not rs.Eof then
        cdGerente = rs("CDUSUARIO")
		dsGerente = getUserDescription(cdGerente)
		firmaGerente = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if
    if not rs.Eof then
        cdController = rs("CDUSUARIO")
        dsController = getUserDescription(cdController)
		firmaController = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if
	if not rs.Eof then
        cdDirector = rs("CDUSUARIO")
		dsDirector = getUserDescription(cdDirector)
		firmaDirector = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
		rs.movenext
    end if
	
	if (firmaGerente <> "") then
		firma = obtenerFirma(cdGerente)		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 10, 725, 185, 75, 0)
	end if
   	if (firmaController <> "") then
   		firma = obtenerFirma(cdController)   		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 202,725, 185, 75, 0)
	end if
   	if (firmaDirector <> "") then
   		firma = obtenerFirma(cdDirector)   		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 394, 725, 185, 75, 0)
	end if
	call GF_setFont(oPDF,"ARIAL",10,0)
	if (dsGerente <> "") then Call GF_writeTextAlign(oPDF,10, 795, dsGerente, 190, PDF_ALIGN_CENTER)
	if (dsController <> "") then Call GF_writeTextAlign(oPDF,195, 795, dsController, 190, PDF_ALIGN_CENTER)
	if (dsDirector <> "") then Call GF_writeTextAlign(oPDF,380, 795, dsDirector, 190, PDF_ALIGN_CENTER)			
	call GF_setFont(oPDF,"ARIAL",6,0)
	if (firmaGerente <> "")	then Call GF_writeTextAlign(oPDF,10, 810, firmaGerente, 190, PDF_ALIGN_CENTER)
	if (firmaController <> "") then	Call GF_writeTextAlign(oPDF,195, 810, firmaController, 190, PDF_ALIGN_CENTER)
	if (firmaDirector <> "") then Call GF_writeTextAlign(oPDF,380, 810, firmaDirector, 190, PDF_ALIGN_CENTER)
End Function
'------------------------------------------------------------------------------------------------------		 
'Dibuja la firma de un ajuste determinado
Function writeSignatureMermaVolatil(pPto, pIdAjuste)
    Dim strSQL, rs,cdGerente, dsGerente, firmaGerente, cdController, dsController, firmaController, cdDirector, dsDirector, firmaDirector, cdSupPto, dsSupPto, firmaMercaderia,rsEst
    
    strSQL =  "SELECT * FROM TBLAJUSTESFIRMAS WHERE IDAJUSTE = "& pIdAjuste &" ORDER BY SECUENCIA"
	Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	if not rs.Eof then
        cdGtePto = rs("CDUSUARIO")
		dsGtePto = getUserDescription(cdGtePto)
		firmaGtePto = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if
    if not rs.Eof then
        cdSupPto = rs("CDUSUARIO")
		dsSupPto = getUserDescription(cdSupPto)
		firmaMercaderia = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if
    if not rs.Eof then
        cdController = rs("CDUSUARIO")
		dsController = getUserDescription(cdController)
        firmaController = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if
	if not rs.Eof then
        cdDirector = rs("CDUSUARIO")
		dsDirector = getUserDescription(cdDirector)
        firmaDirector = armarTextoPlanoFirma(rs("HKEY"), rs("MMTO"))
        rs.MoveNext()
    end if

    if (firmaGtePto <> "") then
	    firma = obtenerFirma(cdGtePto)		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 40, 620, 245, 75, 0)
	end if
    if (firmaMercaderia <> "") then
	    firma = obtenerFirma(cdSupPto)		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 326, 620, 245, 75, 0)
	end if
   	if (firmaController <> "") then
   		firma = obtenerFirma(cdController)   		
	    Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 40, 735, 245, 75, 0)
	end if
    if (firmaDirector <> "") then
   	    firma = obtenerFirma(cdDirector)   		
		Call GF_writeImage(oPDF, pathImg & "\Firmas\" & firma, 326, 735, 245, 75, 0)
	end if

	call GF_setFont(oPDF,"ARIAL",10,0)
    if (dsGtePto <> "") then Call GF_writeTextAlign(oPDF,10, 688, dsGtePto, 286, PDF_ALIGN_CENTER)
	if (dsSupPto <> "") then Call GF_writeTextAlign(oPDF,296, 688, dsSupPto, 286, PDF_ALIGN_CENTER)
    if (dsController <> "") then Call GF_writeTextAlign(oPDF,10, 803, dsController, 286, PDF_ALIGN_CENTER)
	if (dsDirector <> "") then Call GF_writeTextAlign(oPDF,296, 803, dsDirector, 286, PDF_ALIGN_CENTER)			
	call GF_setFont(oPDF,"ARIAL",6,0)
    if (firmaGtePto <> "")	then Call GF_writeTextAlign(oPDF,10, 700, firmaGtePto, 286, PDF_ALIGN_CENTER)
	if (firmaMercaderia <> "")	then Call GF_writeTextAlign(oPDF,296, 700, firmaMercaderia, 286, PDF_ALIGN_CENTER)
	if (firmaController <> "") then	Call GF_writeTextAlign(oPDF,10, 815, firmaController, 286, PDF_ALIGN_CENTER)
	if (firmaDirector <> "") then Call GF_writeTextAlign(oPDF,296, 815, firmaDirector, 286, PDF_ALIGN_CENTER)
    
End Function
'------------------------------------------------------------------------------------------------------		 
'Function drawBoxSignature: dibuja la estructura de las firmas 
Function drawBoxSignature()
    pY = 706
	Call GF_squareBox(oPDF, 10, pY, 190, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, 200, pY, 190, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, 390, pY, 190, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	
    pY = pY + 3
	Call GF_setFont(oPDF,"ARIAL",8,8)
	
    Call GF_writeTextAlign(oPDF,10, pY,  GF_TRADUCIR("Aprobación Gerente de Planta"),   	      190, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,195, pY, GF_TRADUCIR("Aprobación Controller"),    190, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,380, pY, GF_TRADUCIR("Aprobación Director"),    190, PDF_ALIGN_CENTER)
	
    pY = pY + 12
    Call GF_squareBox(oPDF, 10,  pY, 190, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 200, pY, 190, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 390, pY, 190, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	
    Call GF_setFontColor("#000000")
	Call GF_setFont(oPDF,"ARIAL",8,0)	
End Function
'------------------------------------------------------------------------------------------------------		 
'Function drawBoxSignature: dibuja la estructura de las firmas 
Function drawBoxSignatureMermaVolatil()
    pY = 595
	Call GF_squareBox(oPDF, 10,  pY, 286, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, 296, pY, 286, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 3
	Call GF_setFont(oPDF,"ARIAL",8,8)
	
    Call GF_writeTextAlign(oPDF,10, pY,  GF_TRADUCIR("Aprobación Gerente de Puerto"),   	286, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,296, pY, GF_TRADUCIR("Aprobación Supervisor de Puertos"),  	286, PDF_ALIGN_CENTER)
    pY = pY + 12
    Call GF_squareBox(oPDF, 10,  pY, 286, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 296, pY, 286, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 100
	
    Call GF_squareBox(oPDF, 10,  pY, 286, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    Call GF_squareBox(oPDF, 296, pY, 286, 15, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
    pY = pY + 3
	Call GF_setFont(oPDF,"ARIAL",8,8)
	
    Call GF_writeTextAlign(oPDF,10, pY,  GF_TRADUCIR("Aprobación Controller"),   	      286, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,296, pY, GF_TRADUCIR("Aprobación Director"),    286, PDF_ALIGN_CENTER)
    pY = pY + 12
    Call GF_squareBox(oPDF, 10,  pY, 286, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 296, pY, 286, 100, 0, "", "#000000", 1, PDF_SQUARE_NORMAL)

    Call GF_setFontColor("#000000")
	Call GF_setFont(oPDF,"ARIAL",8,0)	
End Function
'------------------------------------------------------------------------------------------------------
Function drawBoxObservaciones(ByRef py)
    Call GF_setFont(oPDF,"ARIAL",8,8)
    Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("DETALLE:")			, 180, PDF_ALIGN_LEFT)
    Call GF_setFontColor("#000000")	
    Call GF_setFont(oPDF,"ARIAL",8,0)
End Function
'------------------------------------------------------------------------------------------------------
Function armadoPDF(pto, pIdAjuste,pFechaDesde,pFechaHasta,pConcepto,pMode,pRol)
	Dim strSQL, rs, pathPDF	,flagFirmaMermaVolatil,auxFechaDesde,auxFechaHasta
	pathPDF = Server.MapPath("temp\AJUSTE_PTO_"& pIdAjuste &"-"& pto &".pdf" )
	Set oPDF = GF_createPDF(pathPDF)
	Call GF_setPDFMODE(pMode)
    strSQL = "SELECT CASE WHEN A.FECHADESDE IS NULL THEN 0 ELSE A.FECHADESDE END AS FECHADESDE, "&_
             "       CASE WHEN A.FECHAHASTA IS NULL THEN 0 ELSE A.FECHAHASTA END AS FECHAHASTA, "&_
             "       A.CDAJUSTE, "&_
             "       A.CDPRODUCTO, "&_
             "       CASE WHEN A.KILOSAJUSTE IS NULL THEN 0 ELSE A.KILOSAJUSTE END AS KILOSAJUSTE, "&_
             "       A.IDORIGEN, "&_
             "       B.DSPRODUCTO "&_
             "FROM   DBO.TBLAJUSTES A "&_
             "   INNER JOIN DBO.PRODUCTOS B ON B.CDPRODUCTO = A.CDPRODUCTO "&_
             "WHERE  1 = 1 "
    if (Cdbl(pIdAjuste) <> 0) then   strSQL = strSQL & " AND A.IDAJUSTE = " & pIdAjuste
    if (pFechaDesde <> "") then      strSQL = strSQL & " AND FECHADESDE >= "&pFechaDesde
    if (pFechaHasta <> "") then      strSQL = strSQL & " AND FECHADESDE <= "&pFechaHasta
    if (pConcepto <> "") then        strSQL = strSQL & " AND CDAJUSTE = '"& pConcepto &"'"
	'Si no tiene un idajuste entonces tiene rol asignado , filtro aquellos ajustes que tengan el estado adecuado para verlo
    
    if (Cdbl(pIdAjuste) = 0) then
        strSQL = strSQL & " AND A.ESTADO IN (SELECT DISTINCT ESTADOACTUAL "&_
	                      "										     FROM TBLESTADOSTRANSICION "&_
                          "                                          WHERE EVENTO = 1 AND IDSISTEMA = "& SEC_SYS_POSEIDON &" AND DATOAUXILIAR = "& pRol &" AND TIPOOBJETO = A.CDAJUSTE) "    
    end if
    strSQL = strSQL & "ORDER  BY FECHADESDE, FECHAHASTA, CDPRODUCTO, IDORIGEN "
    call GF_BD_Puertos (pto, rs, "OPEN",strSQL)
    'Response.Write strSQL 
    'Response.End
	if not rs.Eof then
        Call drawReportFormat()
        auxFechaDesde = Cdbl(rs("FECHADESDE"))
        auxFechaHasta = Cdbl(rs("FECHAHASTA"))
        if ((pFechaDesde <> "")and(pFechaHasta <> "")) then
            auxFechaDesde = pFechaDesde
            auxFechaHasta = pFechaHasta
        end if
        py = drawHead(pto, pIdAjuste, rs("CDAJUSTE"), auxFechaHasta, auxFechaDesde, Cdbl(rs("CDPRODUCTO")), rs("DSPRODUCTO"), Cdbl(rs("KILOSAJUSTE")), rs("IDORIGEN"))
		'Dibujo las Observaciones, como el tipo ajuste de Merma volatil difiere en su detalle del resto de los tipos de ajuste se los dibuja de diferente modo
        Call drawBoxObservaciones(py)
        flagFirmaMermaVolatil = false
        select case rs("CDAJUSTE")
		    case AJUSTE_DRAFT_SURVEY
    			Call writeObservationsDraft(pto, rs("IDORIGEN"), py)
		    case AJUSTE_MANIPULEO
    			Call GF_writeTextAlign(oPDF, 115, py,GF_TRADUCIR("Ajuste estimativo de Merma por Manipuleo"), 200, PDF_ALIGN_LEFT)
            case AJUSTE_MERMA_VOLATIL
                'Dentro de esta funcion dibuja todo el detalle de las observaciones de merma volatil, se debe totalizar por producto y cliente
    			Call writeObservationsMermaVolatil(pto, rs, py)
                flagFirmaMermaVolatil = true
        end select
        'Si el reporte viene con un idAjuste muestro las firmas de los ajustes, caso contrario no
        if (Cdbl(pIdAjuste) <> 0) then
            if (flagFirmaMermaVolatil) then
                Call drawBoxSignatureMermaVolatil()
		        Call writeSignatureMermaVolatil(pto,pIdAjuste)
            else
                Call drawBoxSignature()
		        Call writeSignature(pto,pIdAjuste)
            end if
        end if
	end if
	Call GF_closePDF(oPDF)
	armadoPDF=pathPDF	
end Function
'----------------------------------------------------------------------------------------------------------------------------------------------
Function drawReportFormat()
    Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,15,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeImage(oPDF, pathImg & "\ADMlogo2.jpg", 10, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL", 18,0)
	Call GF_writeTextAlign(oPDF,10, 20, GF_TRADUCIR("Autorización de Ajuste de Stock de Mercadería"), 570, PDF_ALIGN_CENTER)	
	call GF_setFont(oPDF,"ARIAL",10,0)
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
Function drawHead(pto, pIdAjuste, pCdAjuste, pFechaHasta, pFechaDesde, pCdProducto, pDsProducto, pKilosAjuste, pOrigen)
    Dim fecha, py
	
	'Dibujo los box de la cabecera
    py = 90
	Call GF_squareBox(oPDF, 10 , py , 100, 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 110, py , 390, 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 500, py , 30 , 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 530, py , 50 , 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)	
    py = py + 18
	Call GF_squareBox(oPDF, 10 , py, 100, 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 110, py, 470, 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    py = py + 18
	Call GF_squareBox(oPDF, 10 , py, 100, 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 110, py, 470, 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
    py = py + 18
    if (pCdAjuste <> AJUSTE_MERMA_VOLATIL) then
        Call GF_squareBox(oPDF, 10 , py, 100, 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	    Call GF_squareBox(oPDF, 110, py, 470, 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
        py = py + 18
        Call GF_squareBox(oPDF, 10 , py, 100, 18, 0, "#F4F4F4", "#000000", 1, PDF_SQUARE_NORMAL)
	    Call GF_squareBox(oPDF, 110, py, 470, 18, 0, "#FFFFFF", "#000000", 1, PDF_SQUARE_NORMAL)
        py = py + 18
    end if

	'Dibujo los titulos de la cabecera
    Call GF_setFont(oPDF,"ARIAL",8,8)
	
    py = 95
	Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("CONCEPTO:")				, 180, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,510, py, GF_TRADUCIR("N°:")						,  30, PDF_ALIGN_LEFT)
    py = py + 18
	Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("PUERTO:")				, 180, PDF_ALIGN_LEFT)
	py = py + 18
    Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("PERÍODO:")			, 180, PDF_ALIGN_LEFT)
    py = py + 18
    if (pCdAjuste <> AJUSTE_MERMA_VOLATIL) then
        Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("PRODUCTO:")				, 180, PDF_ALIGN_LEFT)	
        py = py + 18
        Call GF_writeTextAlign(oPDF, 18, py, GF_TRADUCIR("KILOS:")					, 180, PDF_ALIGN_LEFT)
        py = py + 18
    end if
    Call GF_setFontColor("#000000")
    'Dibujo los datos de la caebcera	
	Call GF_setFont(oPDF,"ARIAL",8,0)
    fecha = GF_FN2DTE(pFechaDesde)
	if (pFechaDesde <> pFechaHasta) then fecha = "Desde " & GF_FN2DTE(pFechaDesde) & " hasta " & GF_FN2DTE(pFechaHasta)	
	py = 94
    Call GF_writeTextAlign(oPDF, 115, py, getDsCodigoAjustePuerto(pCdAjuste) &" ("&pCdAjuste&")", 200, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 535, py, pIdAjuste	,  50, PDF_ALIGN_LEFT)
    py = py + 18
	Call GF_writeTextAlign(oPDF, 115, py, UCase(pto) ,  50, PDF_ALIGN_LEFT)
    py = py + 18
	Call GF_writeTextAlign(oPDF, 115, py, fecha	, 200, PDF_ALIGN_LEFT)
    py = py + 18
	if (pCdAjuste <> AJUSTE_MERMA_VOLATIL) then
        Call GF_writeTextAlign(oPDF, 115, py, pCdProducto & " - " &	pDsProducto	, 200, PDF_ALIGN_LEFT)
        py = py + 18
	    if (pKilosAjs < 0) then Call GF_setFontColor("#FF0000")	
	    Call GF_writeTextAlign(oPDF, 115, py, GF_EDIT_DECIMALS(pKilosAjs*100,2) & " Kg."	, 200, PDF_ALIGN_LEFT)
        py = py + 18
    end if
    drawHead = py
End function
'******************************************************************************************************
'**************************************** COMIENZO DE PAGINA ******************************************
'******************************************************************************************************
Dim px, idAjuste,g_Puerto, nroPagina, oPDF, pathImg,fechaDesde,fechaHasta,concepto,cdCliente,cdProducto,idRol

idAjuste = GF_PARAMETROS7("idAjuste",0,6)
concepto = GF_PARAMETROS7("concepto","",6)
g_Puerto = GF_PARAMETROS7("pto","",6)
idRol = 0
if (Cdbl(idAjuste) = 0) then
    fechaDesde = GF_PARAMETROS7("fechaDesde","",6)
    fechaHasta = GF_PARAMETROS7("fechaHasta","",6)
    'Debo obtener el rol del que ingresa a ver el pdf para mostrarle los Ajustes que debe firmar
    Set rsRol = getRolesUsuario(Session("Usuario"), SEC_SYS_POSEIDON)
    if (not rsRol.Eof) then idRol = rsRol("IDROL")
end if
filename = "AJUSTE_PTO_" & g_Puerto & "_" & session("MmtoSistema")
nroPagina = 1
g_strPuerto = g_Puerto

pathImg = server.MapPath("..\images")
if (LCase(request.servervariables("script_name")) = "/actisaintra/poseidon/ajusteautorizacionprint.asp") then
	Call armadoPDF(g_Puerto,idAjuste,fechaDesde,fechaHasta,concepto,PDF_STREAM_MODE,idRol)
end if
%>