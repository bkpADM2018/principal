<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<%
Const PAGE_HEIGHT_SIZE = 800
Const SEPARATION = 12
Const PAGE_TOP_INIT = 62
'------------------------------------------------------------------------------------------	
Function drawFormato(pTitulo)
	Call GF_setFont(oPDF,"COURIER",12,0)
	Call GF_writeTextAlign(oPDF,20, 25, GF_TRADUCIR(pTitulo), 590, PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER",8,0) ' Seteo el formato de la letra, va a ser el mismo para TODO el reporte
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)	
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,15,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
End Function
'------------------------------------------------------------------------------------------
Function drawFiter(pFechaDesde, pFechaHasta, pPto)
	dim xSeparation
	currentAuxY = PAGE_TOP_INIT
	xSeparation = 0
	
	Call GF_writeTextAlign(oPDF, 20 , currentAuxY , GF_TRADUCIR("Puerto: ") & ucase(pPto)	, 180, PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(oPDF, 200 , currentAuxY , GF_TRADUCIR("Fecha Desde: ") & pFechaDesde	, 180, PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(oPDF, 380 , currentAuxY, GF_TRADUCIR("Fecha Hasta: ") & pFechaHasta	, 180, PDF_ALIGN_LEFT)
		
	currentAuxY = currentAuxY + 20
	
	Call GF_horizontalLineDash(oPDF, 20, currentAuxY, 560)	
end Function
'------------------------------------------------------------------------------------------
Function getSQLDescargaEmbarque(pFechaDesde,pFechaHasta,pCdProducto,pVerDetalle,pPto)
	Dim strSQL,rs,diaHoy,strGroupDetalle
	
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)	
		
	If (pVerDetalle)		then strGroupDetalle = "  ,TFINAL.fecha "
	
	If (pVerDetalle) then
		strSQL = "select TFINAL.neto_descarga_cam as neto_descarga_cam, TFINAL.neto_descarga_vag as neto_descarga_vag, TFINAL.neto_descarga_vag + TFINAL.neto_descarga_cam as neto_descarga, TFINAL.neto_embarque, TFINAL.producto, PRO.dsproducto, (YEAR(TFINAL.fecha)*10000 + Month(TFINAL.fecha)*100 + DAY(TFINAL.fecha)) as fecha "
	else
		strSQL = "select sum(TFINAL.neto_descarga_cam) as neto_descarga_cam, sum(TFINAL.neto_descarga_vag) as neto_descarga_vag,sum(TFINAL.neto_descarga_vag)+ sum(TFINAL.neto_descarga_cam) as neto_descarga,Sum (TFINAL.neto_embarque) as neto_embarque,TFINAL.PRODUCTO,PRO.dsproducto "
	end if
	strSQL = strSQL & " from (( "&_
			"	select CASE WHEN Cam.fecha IS NULL THEN  CASE WHEN Vag.fecha IS NULL THEN Emb.fecha ELSE Vag.fecha END "&_
            "   ELSE Cam.fecha END AS FECHA, "&_
            "   CASE WHEN Cam.producto IS NULL THEN  CASE WHEN Vag.producto IS NULL THEN Emb.producto ELSE Vag.producto END "&_
            "   ELSE Cam.producto END AS PRODUCTO, "&_
            "   case when Cam.neto_descarga_cam is null then 0 else Cam.neto_descarga_cam end as neto_descarga_cam, "&_
            "   case when Emb.neto_embarque is null then 0 else Emb.neto_embarque end as neto_embarque, "&_
            "   case when Vag.neto_descarga_vag is null then 0 else Vag.neto_descarga_vag end as neto_descarga_vag "&_            
			"        from ( "&_
			"             SELECT sum (vlNETO) as neto_descarga_cam,fecha,producto "&_
			"             FROM   (SELECT dtcontable as fecha ,cdproducto as producto, "&_
			"                            ((SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
			"                              FROM   dbo.hpesadascamion HC "&_
			"                              WHERE  HC.cdpesada = 1 AND HC.dtcontable = a.dtcontable AND HC.idcamion = a.idcamion "&_
			"                                  AND HC.sqpesada = (SELECT Max(T1.sqpesada) "&_
			"                                                     FROM   dbo.hpesadascamion T1 "&_
			"                                                     WHERE T1.dtcontable = HC.dtcontable AND HC.idcamion = T1.idcamion AND T1.cdpesada = 1)) - "&_
			"                              (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
			"                               FROM dbo.hpesadascamion HC "&_ 
			"                               WHERE HC.cdpesada = 2 AND HC.dtcontable = a.dtcontable AND HC.idcamion = a.idcamion "&_
			"                                   AND HC.sqpesada = (SELECT Max(T1.sqpesada) "&_
			"                                                      FROM   dbo.hpesadascamion T1 "&_
			"                                                      WHERE  HC.dtcontable = T1.dtcontable AND HC.idcamion = T1.idcamion AND T1.cdpesada = 2)) -  "&_
			"                              (SELECT CASE WHEN HMC.vlmermakilos IS NULL THEN 0 ELSE HMC.vlmermakilos END AS VLMERMAKILOS "&_
			"                               FROM hmermascamiones HMC "&_ 
			"                               WHERE HMC.dtcontable = a.dtcontable AND HMC.idcamion = a.idcamion "&_
			"                                    AND HMC.sqpesada = (SELECT Max(sqpesada) "&_
			"                                                        FROM   hmermascamiones "&_
			"                                                        WHERE  HMC.dtcontable = dtcontable AND HMC.idcamion = idcamion)) ) AS vlNETO "&_
			"                     FROM   (SELECT B.dtcontable,B.idcamion,B.cdproducto "&_
			"                             FROM (select dtcontable,idcamion,cdproducto from hcamiones WHERE dtcontable >= '"& pFechaDesde &"' AND dtcontable <= '"& pFechaHasta &"' AND cdestado IN( 6, 8 ) "
			if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
			strSQL = strSQL & "							) B "&_
			"                             INNER JOIN hcamionesdescarga c ON c.idcamion = B.idcamion AND c.dtcontable = b.dtcontable) A "&_
			"                )  as T1 group by fecha,producto ) as  Cam "&_
			" FULL OUTER JOIN ( "&_
            "            SELECT Sum (vlneto) AS neto_descarga_vag, fecha,producto "&_
            "             FROM   (SELECT AA.CDOPERATIVO, AA.NUCARTAPORTE,AA.CDVAGON, AA.dtcontable AS fecha, AA.cdproducto AS producto, "&_
            "                            ((SELECT CASE WHEN pv.vlpesada IS NULL THEN 0 ELSE pv.vlpesada END AS vlpesada "&_
            "                               FROM   dbo.Hpesadasvagon pv "&_
            "                               WHERE pv.dtcontable = AA.dtcontable AND pv.CDOPERATIVO = AA.CDOPERATIVO AND pv.NUCARTAPORTE = AA.NUCARTAPORTE AND pv.CDVAGON = AA.CDVAGON AND pv.CDPESADA = 1 "&_
            "                                    AND pv.sqpesada = ( SELECT MAX(SQPESADA)  "&_
            "                                                        FROM dbo.hPESADASVAGON  "&_
            "                                                        WHERE pv.dtcontable = dtcontable AND pv.NUCARTAPORTE = NUCARTAPORTE AND pv.CDVAGON = CDVAGON AND CDPESADA = 1 )) - "&_
            "                             (SELECT CASE WHEN pv.vlpesada IS NULL THEN 0 ELSE pv.vlpesada END AS vlpesada "&_
            "                               FROM   dbo.hpesadasvagon pv "&_
            "                               WHERE pv.dtcontable = AA.dtcontable AND pv.CDOPERATIVO = AA.CDOPERATIVO AND pv.NUCARTAPORTE = AA.NUCARTAPORTE AND pv.CDVAGON = AA.CDVAGON AND pv.CDPESADA = 2 "&_
            "                                    AND pv.sqpesada = ( SELECT MAX(SQPESADA)  "&_
            "                                                        FROM dbo.HPESADASVAGON  "&_
            "                                                        WHERE pv.dtcontable = dtcontable AND pv.NUCARTAPORTE = NUCARTAPORTE AND pv.CDVAGON = CDVAGON AND CDPESADA = 2 ))  - "&_
            "                                (SELECT CASE WHEN mv.VLMERMAKILOS IS NULL THEN 0 ELSE mv.VLMERMAKILOS END  "&_
            "                                 FROM HMERMASVAGONES mv "&_
            "                                 WHERE mv.dtcontable = AA.dtcontable AND mv.CDOPERATIVO = AA.CDOPERATIVO AND mv.NUCARTAPORTE = AA.NUCARTAPORTE  AND mv.CDVAGON = AA.CDVAGON   "&_
            "                                    AND mv.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASVAGONES "&_
            "                                                        WHERE mv.dtcontable = dtcontable AND NUCARTAPORTE = AA.NUCARTAPORTE AND CDVAGON = AA.CDVAGON )) ) AS vlNETO "&_
            "                     FROM   (SELECT B.CDOPERATIVO, B.NUCARTAPORTE,B.CDVAGON,B.CDPRODUCTO,B.DTCONTABLE "&_
            "                             FROM   (SELECT dtcontable, CDOPERATIVO, NUCARTAPORTE,CDVAGON,CDPRODUCTO "&_
            "                                     FROM   dbo.Hvagones WHERE dtcontable >= '"& pFechaDesde &"' AND dtcontable <= '"& pFechaHasta &"' AND cdestado IN( 6, 8 ) "
            if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
            strSQL = strSQL & "							) B "&_
            "                             INNER JOIN HOPERATIVOS O ON O.DTCONTABLE = B.DTCONTABLE AND O.NUCARTAPORTE = B.NUCARTAPORTE and O.CDOPERATIVO = B.CDOPERATIVO ) AA) AS T1 "&_
            "              GROUP  BY  fecha, producto "&_
            "    ) AS Vag ON Vag.fecha = cam.fecha AND Vag.producto = cam.producto "&_
			"                 FULL OUTER JOIN (SELECT dtcarga as fecha,cdproducto as producto ,sum(vlkilos) as neto_embarque "&_
			"                                  FROM   dbo.hcargasembarque WHERE dtcarga >= '"& pFechaDesde &"' AND dtcarga <= '"& pFechaHasta &"'"
			if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
			strSQL = strSQL & "					group by dtcarga,cdproducto) Emb "&_
			"                     ON Emb.fecha = cam.fecha AND Emb.producto = cam.producto  "&_
			"          ) UNION ( "&_
			"   select CASE WHEN Cam.fecha IS NULL THEN CASE WHEN Vag.fecha IS NULL THEN Emb.fecha ELSE Vag.fecha END "&_
            "			ELSE Cam.fecha END AS FECHA, "&_
            "			CASE WHEN Cam.producto IS NULL THEN CASE WHEN Vag.producto IS NULL THEN Emb.producto ELSE Vag.producto END "&_
            "			ELSE Cam.producto END AS PRODUCTO,  "&_
            "			case when Cam.neto_descarga_cam is null then 0 else Cam.neto_descarga_cam end as neto_descarga_cam, "&_
            "			case when Emb.neto_embarque is null then 0 else Emb.neto_embarque end as neto_embarque, "&_
            "			case when Vag.neto_descarga_vag is null then 0 else Vag.neto_descarga_vag end as neto_descarga_vag "&_
			"          from ( "&_
			"             SELECT sum (vlNETO) as neto_descarga_cam,fecha,producto "&_
			"             FROM   (SELECT dtcontable as fecha ,cdproducto as producto, "&_
			"                       ((SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
			"                         FROM   dbo.pesadascamion HC "&_
			"                         WHERE  HC.cdpesada = 1 AND a.idcamion = hc.idcamion "&_
			"                                AND HC.sqpesada = (SELECT Max(T1.sqpesada) "&_
			"                                                   FROM   dbo.pesadascamion T1 "&_
			"                                                   WHERE  T1.cdpesada = 1 AND T1.idcamion = Hc.idcamion)) - "&_
			"                         (SELECT CASE WHEN HC.vlpesada IS NULL THEN 0 ELSE HC.vlpesada END AS vlpesada "&_
			"                          FROM   dbo.pesadascamion HC "&_
			"                          WHERE  HC.cdpesada = 2 AND a.idcamion = hc.idcamion "&_
			"                            AND sqpesada = (SELECT Max(T1.sqpesada) "&_
			"                                            FROM   dbo.pesadascamion T1 "&_
			"                                            WHERE  T1.cdpesada = 2 AND T1.idcamion = HC.idcamion)) - "&_
			"                         (SELECT CASE WHEN MC.vlmermakilos IS NULL THEN 0 ELSE MC.vlmermakilos END AS VLMERMAKILOS "&_
			"                          FROM   mermascamiones MC "&_
			"                          WHERE  MC.idcamion = a.idcamion "&_
			"                            AND MC.sqpesada = (SELECT Max(sqpesada) "&_
			"                                               FROM   mermascamiones "&_
			"                                               WHERE  MC.idcamion = idcamion))) AS vlNETO "&_
			"                     FROM   (SELECT B.dtcontable,B.idcamion,B.cdproducto "&_
			"                             FROM (select '"&diaHoy&"' AS dtcontable,idcamion,cdproducto from camiones WHERE  cdestado IN( 6, 8 ) "
			if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
			strSQL = strSQL & "							) B "&_
			"                             INNER JOIN camionesdescarga c ON c.idcamion = B.idcamion ) A "&_
			"                )  as T1 group by fecha,producto ) as  Cam "&_
			" FULL OUTER JOIN (  "&_
            "     SELECT Sum (vlneto) AS neto_descarga_vag, fecha, producto "&_
            "     FROM   (SELECT AA.CDOPERATIVO, AA.NUCARTAPORTE,AA.CDVAGON, AA.dtcontable AS fecha, AA.cdproducto AS producto, "&_
            "                    ((SELECT CASE WHEN pv.vlpesada IS NULL THEN 0 ELSE pv.vlpesada END AS vlpesada "&_
            "                       FROM   dbo.pesadasvagon pv "&_
            "                      WHERE pv.CDOPERATIVO = AA.CDOPERATIVO AND pv.NUCARTAPORTE = AA.NUCARTAPORTE AND pv.CDVAGON = AA.CDVAGON AND pv.CDPESADA = 1 "&_
            "                            AND pv.sqpesada = ( SELECT MAX(SQPESADA) "&_ 
            "                                                FROM dbo.PESADASVAGON  "&_
            "                                                WHERE pv.NUCARTAPORTE = NUCARTAPORTE AND pv.CDVAGON = CDVAGON AND CDPESADA = 1 )) - "&_
            "                     (SELECT CASE WHEN pv.vlpesada IS NULL THEN 0 ELSE pv.vlpesada END AS vlpesada "&_
            "                       FROM   dbo.pesadasvagon pv "&_
            "                       WHERE pv.CDOPERATIVO = AA.CDOPERATIVO AND pv.NUCARTAPORTE = AA.NUCARTAPORTE AND pv.CDVAGON = AA.CDVAGON AND pv.CDPESADA = 2 "&_
            "                            AND pv.sqpesada = ( SELECT MAX(SQPESADA) "&_ 
            "                                                FROM dbo.PESADASVAGON  "&_
            "                                                WHERE pv.NUCARTAPORTE = NUCARTAPORTE AND pv.CDVAGON = CDVAGON AND CDPESADA = 2 ))  - "&_
            "                        (SELECT CASE WHEN mv.VLMERMAKILOS IS NULL THEN 0 ELSE mv.VLMERMAKILOS END  "&_
            "                         FROM MERMASVAGONES mv "&_
            "                         WHERE mv.CDOPERATIVO = AA.CDOPERATIVO AND mv.NUCARTAPORTE = AA.NUCARTAPORTE  AND mv.CDVAGON = AA.CDVAGON   "&_
            "                            AND mv.SQPESADA= (SELECT MAX(SQPESADA) FROM MERMASVAGONES "&_
            "                                                WHERE NUCARTAPORTE = AA.NUCARTAPORTE AND CDVAGON = AA.CDVAGON )) ) AS vlNETO "&_
            "             FROM   (SELECT B.CDOPERATIVO, B.NUCARTAPORTE,B.CDVAGON,B.CDPRODUCTO,B.DTCONTABLE "&_
            "                     FROM   (SELECT '"&diaHoy&"' AS dtcontable, CDOPERATIVO, NUCARTAPORTE,CDVAGON,CDPRODUCTO "&_
            "                             FROM   dbo.vagones WHERE  cdestado IN( 6, 8 ) "
    		if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
			strSQL = strSQL & "							) B "&_
            "                     INNER JOIN OPERATIVOS O ON O.NUCARTAPORTE = B.NUCARTAPORTE and O.CDOPERATIVO = B.CDOPERATIVO ) AA) AS T1 "&_
            "      GROUP  BY  fecha, producto "&_
            " ) AS Vag ON Vag.fecha = cam.fecha AND Vag.producto = cam.producto "&_
			"                 FULL OUTER JOIN (SELECT dtcarga as fecha,cdproducto as producto ,sum(vlkilos) as neto_embarque "&_
			"                                  FROM   dbo.cargasembarque WHERE dtcarga >= '"& pFechaDesde &"' AND dtcarga <= '"& pFechaHasta &"'"
			if (pCdProducto <> 0)	then strSQL = strSQL & " and cdproducto in("& pCdProducto &")"
			strSQL = strSQL & "					group by dtcarga,cdproducto) Emb "&_
			"                     ON Emb.fecha = cam.fecha AND Emb.producto = cam.producto  "&_
			"          )) as TFINAL "&_
			"INNER JOIN productos PRO ON PRO.cdproducto = TFINAL.PRODUCTO "&_
			"WHERE TFINAL.fecha >= '"& pFechaDesde &"' and TFINAL.fecha <= '"& pFechaHasta &"' "
			If (not pVerDetalle) then strSQL = strSQL & " GROUP BY TFINAL.PRODUCTO,PRO.dsproducto" & strGroupDetalle
			strSQL = strSQL & " ORDER BY TFINAL.PRODUCTO,PRO.dsproducto" & strGroupDetalle
						
	Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	Set getSQLDescargaEmbarque = rs
End Function
'---------------------------------------------------------------------------------------------------------------
Function drawBody(pFechaDesde,pFechaHasta,pCdProducto,pVerDetalle,pPto)
    Dim totToneladaDescargaCam,totToneladaDescargaVag,totToneladaEmbarque
    Set rs = getSQLDescargaEmbarque(pFechaDesde,pFechaHasta,pCdProducto,pVerDetalle,pPto)
	currentAuxY = currentAuxY + SEPARATION
	if (rs.Eof) then
		Call GF_writeTextAlign(oPDF, 20, currentAuxY , GF_TRADUCIR("No se encontraron resultados"), 550	, PDF_ALIGN_CENTER)
    else
		while (not rs.Eof)
			if currentAuxY > PAGE_HEIGHT_SIZE then Call nuevaPagina()
			Call drawCabeceraItem(rs("DSPRODUCTO"),pVerDetalle)
			totToneladaDescargaVag = 0
			totToneladaDescargaCam = 0
			totToneladaEmbarque	   = 0
			cdproducto_old = rs("PRODUCTO")
			while (corteControlByProducto(rs,cdproducto_old))
				if currentAuxY > PAGE_HEIGHT_SIZE then Call nuevaPagina()
				Call drawItem(rs,pVerDetalle)
				if (pVerDetalle) then
					totToneladaDescargaCam = totToneladaDescargaCam + Cdbl(rs("NETO_DESCARGA_CAM"))
					totToneladaDescargaVag = totToneladaDescargaVag + Cdbl(rs("NETO_DESCARGA_VAG"))
					totToneladaEmbarque = totToneladaEmbarque + Cdbl(rs("NETO_EMBARQUE"))
				end if
				rs.MoveNext()
			wend
			if (pVerDetalle) then Call drawTotalesToneladas(totToneladaEmbarque,totToneladaDescargaCam,totToneladaDescargaVag)
			if not rs.Eof then Call GF_horizontalLineDash(oPDF, 20, currentAuxY, 560)
			currentAuxY = currentAuxY + SEPARATION
		wend
    end if
End Function
'---------------------------------------------------------------------------------------------------------------
Function corteControlByProducto(pRs,pCdProdDes)
    Dim rtrn
    rtrn = false
    if not pRs.Eof then		
		if (Cdbl(pRs("PRODUCTO")) = Cdbl(pCdProdDes)) then rtrn = true
	end if
    corteControlByProducto = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------
Function drawTotalesToneladas(totToneladaEmbarque,totToneladaDescargaCam,totToneladaDescargaVag)
	Call GF_setFont(oPDF,"COURIER",8,8)
	currentAuxY = currentAuxY + 6
	Call GF_writeTextAlign(oPDF, 30,  currentAuxY  , GF_TRADUCIR("Total"), 70	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 100, currentAuxY , GF_EDIT_DECIMALS(CDbl(totToneladaDescargaCam),3) &" Tn." , 100	, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, 200, currentAuxY , GF_EDIT_DECIMALS(CDbl(totToneladaDescargaVag),3) &" Tn." , 100	, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, 300, currentAuxY , GF_EDIT_DECIMALS(CDbl(totToneladaDescargaCam) + CDbl(totToneladaDescargaVag),3) &" Tn." , 100	, PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF, 400, currentAuxY , GF_EDIT_DECIMALS(CDbl(totToneladaEmbarque),3) &" Tn." , 150	, PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"COURIER",8,0)
	currentAuxY = currentAuxY + SEPARATION
End Function
'-----------------------------------------------------------------------------------
Function drawCabeceraItem(pDsProd,pVerDetalle)
	Call GF_writeTextAlign(oPDF, 30, currentAuxY , GF_TRADUCIR("Producto: ") & pDsProd , 220	, PDF_ALIGN_LEFT)
	currentAuxY = currentAuxY + SEPARATION
	if (pVerDetalle) then Call GF_writeTextAlign(oPDF, 30, currentAuxY	, GF_TRADUCIR("Fecha"), 70	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 100, currentAuxY , GF_TRADUCIR("Recibido") , 300	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 400, currentAuxY , GF_TRADUCIR("Embarcado"), 170	, PDF_ALIGN_CENTER)
	currentAuxY = currentAuxY + SEPARATION
	Call GF_writeTextAlign(oPDF, 100, currentAuxY , GF_TRADUCIR("Camiones"), 100	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 200, currentAuxY , GF_TRADUCIR("Vagones"),  100	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 300, currentAuxY , GF_TRADUCIR("Total"),    100	, PDF_ALIGN_CENTER)
	currentAuxY = currentAuxY + SEPARATION
End Function
'-----------------------------------------------------------------------------------
Function drawItem(pRs,pVerDetalle)
	
	if (pVerDetalle) then
		Call GF_writeTextAlign(oPDF, 30, currentAuxY	, GF_FN2DTE(pRs("FECHA")) , 70	, PDF_ALIGN_CENTER)
	else
		Call GF_writeTextAlign(oPDF, 30, currentAuxY	, GF_TRADUCIR("Total"), 70	, PDF_ALIGN_CENTER)
	end if	
	'DESCARGA DE CAMIONES
	Call GF_writeTextAlign(oPDF, 100, currentAuxY	, GF_EDIT_DECIMALS(Cdbl(pRs("NETO_DESCARGA_CAM")),3) &" Tn."	, 100	, PDF_ALIGN_RIGHT)
	'DESCARGA DE VAGONES
	Call GF_writeTextAlign(oPDF, 200, currentAuxY	, GF_EDIT_DECIMALS(Cdbl(pRs("NETO_DESCARGA_VAG")),3) &" Tn."	, 100	, PDF_ALIGN_RIGHT)
	'DESCARGA TOTAL (CAMIONES + VAGONES)
	Call GF_writeTextAlign(oPDF, 300, currentAuxY	, GF_EDIT_DECIMALS(Cdbl(pRs("NETO_DESCARGA")),3) &" Tn."	, 100	, PDF_ALIGN_RIGHT)
	'EMBARQUES
	Call GF_writeTextAlign(oPDF, 400, currentAuxY	, GF_EDIT_DECIMALS(Cdbl(pRs("NETO_EMBARQUE")),3) &" Tn."    , 150	, PDF_ALIGN_RIGHT)
	
	currentAuxY = currentAuxY + SEPARATION
End function
'-----------------------------------------------------------------------------------
Function armarPDF(pFechaDesde,pFechaHasta,pCdProducto,pVerDetalle,pPto,pTipo)
	Dim filename
	filename   = "DESCARGA_EMBARQUES_" & Left(session("MmtoDato"),8)
	pathPDF = Server.MapPath("../temp/" & filename)
	Set oPDF = GF_createPDF(pathPDF)
	Call GF_setPDFMODE(pTipo)
	Call drawFormato("REPORTE DESCARGAS Y EMBARQUES")
	Call drawFiter(pFechaDesde,pFechaHasta,pPto)
	Call drawBody(pFechaDesde,pFechaHasta,pCdProducto,pVerDetalle,pPto)
	Call GF_closePDF(oPDF)
	armarPDF =pathPDF
end function
'-----------------------------------------------------------------------------------
function nuevaPagina()
	Call GF_newPage(oPDF)
	nroPagina = nroPagina + 1	
	Call drawFormato("REPORTE DESCARGAS Y EMBARQUES")
	currentAuxY = PAGE_TOP_INIT 
end function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim oPDF, g_strPuerto,nroPagina, currentAuxY,fechaDesde,fechaHasta,cdProducto,verDetalle,flagDetalle

nroPagina = 1
   
g_strPuerto = GF_Parametros7("Pto","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
fechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
verDetalle =GF_PARAMETROS7("verDetalle", "", 6)
cdProducto = GF_PARAMETROS7("cdProducto", "", 6)

flagDetalle = false
if (verDetalle = "on") then flagDetalle = true

Call armarPDF(fechaDesde,fechaHasta, cdProducto, flagDetalle,g_strPuerto, PDF_STREAM_MODE)

%>

