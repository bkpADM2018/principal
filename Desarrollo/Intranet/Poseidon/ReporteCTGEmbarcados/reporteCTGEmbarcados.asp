<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosUser.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<%
Const EMBARQUE_ESPECIAL_AJUSTES = 99999

dim cdProducto, dsProducto, cdCliente, cdCamionesDe, cdCosecha, tipo, cdAviso, cdAvisoAnt, cdBuque, dsBuque, kilos,cdProductoAnt, kilosPrevios, mySelected, kilosSobrantes, kilosInformados, msjEmbarcados
dim myTableHTML, dicErr, accion, mySaveText, puerto, myDescomposicion, kilosFinal, controlsState, cargaExclusivaCosecha, secuencia, kilosReales, listOfProducts, listOfHarvest
Set dicErr = Server.CreateObject("Scripting.Dictionary")
puerto = Request("Pto")
controlsState = ""
listOfProducts = ""
listOfHarvest = ""
kilosPrevios = 0
kilosInformados = 0
KilosAlcanzado = True
cdProducto = GF_Parametros7("cdProducto", "", 6)
tipo = GF_Parametros7("tipo", "", 6)
if cdProducto <> "" then
	myDescomposicion = split(cdProducto,"//")
	cdProducto = cint(myDescomposicion(0))
	dsProducto = trim(myDescomposicion(1))
else
    cdProducto = 0    	
end if
cdCliente = GF_Parametros7("cdCliente", 0, 6)
if cdCliente = 0 then cdCliente = 1

cdCamionesDe = GF_Parametros7("cdCamionesDe", 0, 6)

cdCosecha = GF_Parametros7("cdCosecha", 0, 6)

secuencia = GF_Parametros7("secuencia", "", 6) 
accion = GF_Parametros7("accion", "", 6) 

kilos = GF_Parametros7("kilos", 0, 6)

cdAviso = GF_Parametros7("cdAviso", 0, 6)
cdAvisoAnt = GF_Parametros7("cdAvisoAnt", 0, 6)
if clng(cdAviso) <> clng(cdAvisoAnt) then
	cdAvisoAnt = cdAviso 
	kilos = 0
end if	
if cdAviso <> 0 then
	if loadDatosCargas(cdAviso, cdBuque, dsBuque, cdProducto, listOfProducts) then	    
		if cdproducto <> 0 then
			'VER VALORES INGRESADOS
			call loadDatosCosechas(cdAviso, cdProducto, listOfHarvest)
			call loadDatosCargasKilos(cdAviso, cdProducto, cdCosecha, kilosInformados)
			call loadKilosCargados(cdAviso, cdProducto, cdCosecha, kilosPrevios)
			if accion <> "HIS2" then
			    if (cdAviso <> EMBARQUE_ESPECIAL_AJUSTES) then
				    if (abs((cdbl(kilosInformados)-cdbl(kilosPrevios))) = 0)  then
					    setError(CARGA_COMPLETA)
					    kilos = 0
				    else	
					    if kilos > abs((cdbl(kilosInformados)-cdbl(kilosPrevios))) then
						    setError(KILOS_EXCEDEN)
					    end if
				    end if						
				end if
			else
				kilosInformados = kilos
				kilosPrevios = kilos
			end if
		else
			setError(PRODUCTO_REQUERIDO)
		end if	
	else
		setError(FALTAN_DATOS_AVISO)		
	end if		
else	
	listOfProducts = "0"
	setError(AVISO_REQUERIDO)
	kilos = 0
end if

if accion = "SAVE" then
	if not hayError() then	
		kilosReales = ArmarCtaCteCTG(cdBuque, cdProducto, cdCliente, kilos, "INSERTAR", kilosSobrantes, cdAviso, cdCosecha)	
		kilosPrevios = cdbl(kilosPrevios) + cdbl(kilos)
		kilos = kilosReales
		accion = "HIS2"
	else
		accion = ""
	end if	
end if
'Armado de titulos
if accion = "HIST" then
	clearError
	myTableHTML =   myTableHTML &	" <tr class='reg_Header_nav'> " & _
									"	<td width='3%' align='center'>.</td> " & _
									"	<td width='5%' align='center'>" & GF_Traducir("Aviso")			& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Buque")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Kilos")	& "</td> " & _
									"</tr>	"
						
	call MostrarHistorico
	controlsState = " Disabled "
elseif accion = "HIS2" then
	clearError
	myTableHTML =   myTableHTML &	" <tr class='reg_Header_nav'> " & _
									"	<td align='center'>" & GF_Traducir("Fecha")			& "</td> " & _
									"	<td align='center'>" & GF_Traducir("IdCamion")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Camion")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Acoplado")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Carta Porte")	& "</td> " & _
									"	<td align='center'>" & GF_Traducir("CTG")			& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Cliente")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Producto")		& "</td> " & _
									"	<td align='center'>" & GF_Traducir("Kilos")			& "</td> " & _
									"</tr>	"
	Call CargarHistorico(cdAviso, secuencia)
	controlsState = " Disabled "
else
	if not hayError() then
		if kilos > 0 then
			myTableHTML = myTableHTML & " <TR><TD align='center' COLSPAN='9' class='reg_Header_Warning'><img src='../../images/warning-16x16.png' id='war'>  DATOS SIN CONFIRMAR</TD><TR> "
			myTableHTML =   myTableHTML &	" <tr class='reg_Header_nav'> " & _
										"	<td align='center'>" & GF_Traducir("Fecha")			& "</td> " & _
										"	<td align='center'>" & GF_Traducir("IdCamion")		& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Camion")		& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Acoplado")		& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Carta Porte")	& "</td> " & _
										"	<td align='center'>" & GF_Traducir("CTG")			& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Cliente")		& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Producto")		& "</td> " & _
										"	<td align='center'>" & GF_Traducir("Kilos")			& "</td> " & _
										"</tr>	"
			kilosSobrantes = 0
			myQuery = ArmarCtaCteCTG(cdBuque,cdProducto,cdCliente, kilos, "MOSTRAR", kilosSobrantes, cdAviso, cdCosecha)		
			msjEmbarcados = KilosEmbarcados(myQuery, kilos, kilosSobrantes)
			If msjEmbarcados = kilos Then
			    KilosAlcanzado = True
			Else
			    KilosAlcanzado = False
				myTableHTML =   myTableHTML &	" <tr class='reg_Header_navdosHL'> " & _
												"	<td colspan='8' align='right'>" & GF_Traducir("Total kilos disponibles") & "</td> " & _
												"	<td align='right'>" & GF_EDIT_DECIMALS(cdbl(msjEmbarcados)*100,2) & "</td> " & _
												"</tr>	"			
			    msjEmbarcados = clng(kilos) - clng(msjEmbarcados)
			End If
			myTableHTML = myTableHTML & " <TR><TD align='center' COLSPAN='9' class='reg_Header_Warning'><img src='../../images/warning-16x16.png' id='war'>  DATOS SIN CONFIRMAR</TD><TR> "
		end if
	end if
end if
'----------------------------------------------------------------------------------------
'function existeCamionEnCTG(pIdCamion, pFechaContable)
'Dim rsCTG, rtrn
'rtrn = false
'call GF_BD_Puertos(puerto, rsCTG, "OPEN", "SELECT * FROM CTGEMBARCADOS WHERE IDCAMION='" & pIdCamion & "' AND DTCONTABLE=date('" & pFechaContable & "')")
'if not rsCTG.EOF then rtrn = true
'existeCamionEnCTG = rtrn
'end function
'----------------------------------------------------------------------------------------
Function KilosEmbarcados(pSQLQuery, pKilosDefault, pKilosSobrantes)
'Armado de camiones que componen la carga del buque
dim trClass, cdBuqueAux
Dim rs, NColumnas, acum, dif
dim index
index = 0
	call GF_BD_Puertos(puerto, rs, "OPEN",pSQLQuery)
    while not rs.EOF
		index = index + 1
		trClass = "reg_Header_navdos"
        if(cdbl(pKilosSobrantes) <> 0) then
			acum = acum + cdbl(pKilosSobrantes)
		else
			acum = acum + cdbl(rs("KILOSNETOS"))
		end if	
        If cdbl(acum) >= cdbl(pKilosDefault) Then
            Dif = (cdbl(rs("KILOSNETOS")) - (cdbl(acum) - cdbl(pKilosDefault)))
			myTableHTML =   myTableHTML & "<tr class='" & trClass & "' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>" & _
							"	<td align=center>" & GF_FN2DTE(rs("DTCONTABLE")) & "</td>" & _
							"	<td align=center>" & rs("IDCAMION") & "</td>" & _
							"	<td align=center>" & GF_EDIT_PATENTE(rs("CDCHAPACAMION")) & "</td>" & _
							"	<td align=center>" & GF_EDIT_PATENTE(rs("CDCHAPAACOPLADO")) & "</td>" & _
							"	<td align=center>" & GF_EDIT_CTAPTE(rs("CARTAPORTE")) & "</td>" & _
							"	<td align=center>" & rs("CTG") & "</td>" & _
							"	<td>" & rs("DSCLIENTE") & "</td>" & _
							"	<td align=center>" & rs("DSPRODUCTO") & "</td>" & _
							"	<td align=right>" & GF_EDIT_DECIMALS(cdbl(Dif)*100,2) & "</td>" & _
							"</tr>"
            'Guarda los kilos embarcados y Sale de la funcion
            If(Dif <> 0) then 
				KilosEmbarcados = cdbl(Dif) + (cdbl(acum) - cdbl(rs("KILOSNETOS")))
			else
				KilosEmbarcados = acum
			end if	
            Exit Function
        Else
			'VER CASO DE SOBRANTE EN 0. NO DEBERIA IMPRIMIR NADA
			if cdbl(pKilosSobrantes) <> 0 then 
				kilosFinal = pKilosSobrantes 
				pKilosSobrantes = 0
			else
				kilosFinal = rs("KILOSNETOS")
			end if	
			myTableHTML =   myTableHTML & "<tr class='reg_Header_navdos' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>" & _
							"	<td align=center>" & GF_FN2DTE(rs("DTCONTABLE")) & "</td>" & _
							"	<td align=center>" & rs("IDCAMION") & "</td>" & _
							"	<td align=center>" & GF_EDIT_PATENTE(rs("CDCHAPACAMION")) & "</td>" & _
							"	<td align=center>" & GF_EDIT_PATENTE(rs("CDCHAPAACOPLADO")) & "</td>" & _
							"	<td align=center>" & GF_EDIT_CTAPTE(rs("CARTAPORTE")) & "</td>" & _
							"	<td align=center>" & rs("CTG") & "</td>" & _
							"	<td>" & rs("DSCLIENTE") & "</td>" & _
							"	<td align=center>" & rs("DSPRODUCTO") & "</td>" & _
							"	<td align=right>" & GF_EDIT_DECIMALS(cdbl(kilosFinal)*100,2) & "</td>" & _
							"</tr>"
        End If
        rs.MoveNext
    wend
    KilosEmbarcados = acum
    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------------------------------------------
Function ArmarCtaCteCTG(pCdBuque, pCdproducto, pCdCliente, pKilosACargar, pTipoQuery, byref pKilosSobrantes, pCdAviso, pCdCosecha)
Dim rs, strSQL, strSql2, strSqlSobrantes, fechaInicio, idCamionInicio, kilosCargados
dim auxWhere1, auxWhere2, auxSql1
If cdCamionesDe <> 0 Then 
	auxWhere1 = " AND CDCLIENTE=" & cdCamionesDe
	auxWhere2 = " AND HCD.CDCLIENTE=" & cdCamionesDe
end if	
If pCdCosecha <> 0 Then 
	auxWhere1 = auxWhere1 & " AND CDCOSECHA=" & pCdCosecha
	auxWhere2 = auxWhere2 & " AND HCD.CDCOSECHA=" & pCdCosecha
end if	
fechaInicio = "2010-03-01"
	'Mostrar
	If pTipoQuery = "INSERTAR" Then 
	'INSERTAR
		strsql =	"SELECT * FROM (" & _
					"SELECT (YEAR(TG.DTCONTABLE)*10000 + Month(TG.DTCONTABLE)*100 + DAY(TG.DTCONTABLE)) DTCONTABLE, TG.IDCAMION, TG.CDCHAPACAMION, TG.CDCHAPAACOPLADO, TG.CDCOSECHA, " & _
					"	TG.CARTAPORTE, TG.CTG, TG.CDCLIENTE, TG.CDPRODUCTO, CASE WHEN TG.KILOSCARGADOS IS NULL THEN TG.KILOSNETOS ELSE TG.KILOSNETOS-TG.KILOSCARGADOS END AS KILOSNETOS " & _ 
					"	  FROM " & _
					"	( " & _
					"SELECT HCD.CDCLIENTE, HCD.IDCAMION, HC.CDCHAPACAMION, HC.CDCHAPAACOPLADO, HCD.CDCOSECHA, HCD.DTCONTABLE, RTRIM(HCD.NUCARTAPORTE) + RTRIM(HCD.NUCTAPTEDIG) AS CARTAPORTE, HCD.CTG, HC.CDPRODUCTO, " & _
					"( " & _
					"	( " & _
					"		SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = " & _
					"			(SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1) " & _
					"	) " & _
					"	- " & _
					"	( " & _
					"		SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = " & _
					"			(SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2) " & _
					"	) " & _
					"	- " & _
					"	( " & _
					"		SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= " & _
					"			(SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION) " & _
					"	) " & _
					"	        ) KILOSNETOS , EMBARCADOS.KILOSCARGADOS  " & _
					" FROM HCAMIONESDESCARGA HCD " & _
					"	        LEFT JOIN  " & _
					"	            (SELECT IDCAMION, DTCONTABLE, SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS GROUP BY IDCAMION, DTCONTABLE) " & _
					"	                EMBARCADOS ON HCD.IDCAMION = EMBARCADOS.IDCAMION AND HCD.DTCONTABLE=EMBARCADOS.DTCONTABLE " & _
					" LEFT JOIN HCAMIONES HC ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE=HCD.DTCONTABLE "
	else
	'MOSTRAR
		strsql =	"SELECT * FROM (" & _
					"SELECT (YEAR(TG.DTCONTABLE)*10000 + Month(TG.DTCONTABLE)*100 + DAY(TG.DTCONTABLE)) DTCONTABLE, TG.IDCAMION, TG.CDCHAPACAMION, TG.CDCHAPAACOPLADO, TG.CDCOSECHA, " & _
					"	TG.CARTAPORTE, TG.CTG, TG.DSCLIENTE, TG.DSPRODUCTO, CASE WHEN TG.KILOSCARGADOS IS NULL THEN TG.KILOSNETOS ELSE TG.KILOSNETOS-TG.KILOSCARGADOS END AS KILOSNETOS " & _ 
					"	  FROM " & _
					"	( " & _
					"	    SELECT HCD.DTCONTABLE, HCD.IDCAMION, HC.CDCHAPACAMION, HC.CDCHAPAACOPLADO, HCD.CDCOSECHA, " & _ 
					"	        RTRIM(HCD.NUCARTAPORTE) + RTRIM(HCD.NUCTAPTEDIG) AS CARTAPORTE, HCD.CTG, C.DSCLIENTE, P.DSPRODUCTO, " & _ 
					"	        ( " & _ 
					"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) " & _
					"	            -  " & _
					"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2))  " & _
					"	            -  " & _
					"	            ( SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION)) " & _
					"	        ) KILOSNETOS , EMBARCADOS.KILOSCARGADOS  " & _
					"	    FROM HCAMIONESDESCARGA HCD " & _ 
					"	        LEFT JOIN  " & _
					"	            (SELECT IDCAMION, DTCONTABLE, SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS GROUP BY IDCAMION, DTCONTABLE) " & _
					"	                EMBARCADOS ON HCD.IDCAMION = EMBARCADOS.IDCAMION AND HCD.DTCONTABLE=EMBARCADOS.DTCONTABLE " & _
					"	        LEFT JOIN HCAMIONES HC ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE=HCD.DTCONTABLE  " & _
					"	        LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO  " & _
					"	        LEFT JOIN CLIENTES C ON C.CDCLIENTE = HCD.CDCLIENTE "
	end if

    strSql2 = strSQL &  " WHERE HCD.DTCONTABLE >='" & fechaInicio & "'"
   
	strSql2 = strSql2 & auxWhere2
	strSql2 = strSql2 & " AND HC.CDPRODUCTO = " & pCdproducto & _
						" AND HC.CDESTADO IN (6,8) " & _
						" ) TG)TA WHERE KILOSNETOS>0 ORDER BY DTCONTABLE ASC, IDCAMION ASC"	
	'Response.Write strSQL2
	'Response.End 
	If pTipoQuery = "INSERTAR" Then
	    'Inserto los kilos embarcados
	    kilosCargados = GuardarKilosEmbar(strSql2, pCdBuque, pCdCliente, pKilosACargar, pKilosSobrantes, strSqlSobrantes, pCdAviso, pCdCosecha)
	    ArmarCtaCteCTG = cdbl(kilosCargados)
	Else
	    'Muestra
	    ArmarCtaCteCTG = strSql2
	End If

End Function
'---------------------------------------------------------------------------------------
'Guarda los kilos embarcados
Function GuardarKilosEmbar(pSQL, pCdBuque, pCdCliente, pKilosACarga, pKilosSobrante, pSqlSobrantes, pCdAviso, pCdCosecha)
Dim rs, acumulado, Dif, oldCp, a, rtrn, dtContable
secuencia = getSecuencia(pCdAviso)
call GF_BD_Puertos(puerto, rs, "OPEN",pSQL)
while not rs.eof
        'Corrijo el formato de la DTCONTABLE 
        dtContable =  Left(rs("DTCONTABLE"),4) & "-" & Mid(rs("DTCONTABLE"),5, 2) & "-" & Right(rs("DTCONTABLE"), 2)
		acumulado = cdbl(acumulado) + cdbl(rs("KILOSNETOS"))
		If cdbl(acumulado) >= cdbl(pKilosACarga) Then
		    Dif = cdbl(rs("KILOSNETOS")) - (cdbl(acumulado) - cdbl(pKilosACarga))
		    'Guardar Kilos sobrantes
		    If cdbl(Dif) <> 0 Then 
				call GF_BD_Puertos(puerto, rs, "EXEC","INSERT INTO CTGEMBARCADOS VALUES(" & pCdBuque & "," & pCdCliente & ",'" & rs("IDCAMION") & "' , '" & rs("CDCHAPACAMION") & "' , '" & rs("CDCHAPAACOPLADO") & "' , '" & dtContable & "' , '" & rs("CARTAPORTE") & "' , '" & rs("CTG") & "' ," & rs("CDPRODUCTO") & "," & Dif & "," & pCdAviso & "," & secuencia & "," & rs("CDCOSECHA") & ")")
			end if	
		    if(cdbl(Dif) <> 0) then
				rtrn = cdbl(Dif) + (cdbl(acumulado) - cdbl(rs("KILOSNETOS")))
			else
				rtrn = acumulado
			end if	
			GuardarKilosEmbar = rtrn
		    Exit Function
		Else
			Call GF_BD_Puertos(puerto, rs, "EXEC", "INSERT INTO CTGEMBARCADOS VALUES(" & pCdBuque & "," & pCdCliente & ",'" & rs("IDCAMION") & "' , '" & rs("CDCHAPACAMION") & "' , '" & rs("CDCHAPAACOPLADO") & "' , '" & dtContable & "' , '" & rs("CARTAPORTE") & "' , '" & rs("CTG") & "' ," & rs("CDPRODUCTO") & "," & rs("KILOSNETOS") & "," & pCdAviso & "," & secuencia & "," & rs("CDCOSECHA") & ")")
		End If		
    rs.movenext
wend
rs.Close
Set rs = Nothing
rtrn = acumulado
GuardarKilosEmbar = rtrn
End Function
'---------------------------------------------------------------------------------------
function getSecuencia(pCdAviso)
dim rs, strSQL, rtrn
rtrn = 1
strSQL = "SELECT MAX(SECUENCIA) AS MAXSECUENCIA FROM CTGEMBARCADOS WHERE CDAVISO=" & pCdAviso 
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
if not rs.EOF then
	if not isNull((rs("MAXSECUENCIA"))) then rtrn = cint(rs("MAXSECUENCIA")) + 1
end if	
getSecuencia = rtrn	
end function
'---------------------------------------------------------------------------------------
function loadDatosCosechas(pCdAviso, pCdProducto, byref pListOfHarvest)
dim rs, strSQL, rtrn, index
rtrn = false
if (pCdAviso <> EMBARQUE_ESPECIAL_AJUSTES) then
    strSQL =	" SELECT * FROM EMBARQUESDATOS WHERE CDAVISO= " & pCdAviso & _
			    " AND CDPRODUCTO=" & pCdProducto & _
			    " ORDER BY CDCOSECHA ASC "
else
    strSQL =    " SELECT * FROM COSECHAS where COSDEF='1' and CDPRODUCTO=" & pCdProducto & " ORDER BY CDCOSECHA ASC "
end if			    
'Response.Write strSQL 
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
while not rs.EOF
		index = index + 1
		if len(pListOfHarvest) > 0 then pListOfHarvest = pListOfHarvest & ";"
		pListOfHarvest = pListOfHarvest & rs("CDCOSECHA")
	rs.movenext
wend
if index = 0 then pListOfHarvest = "0"
loadDatosCosechas = rtrn	
end function
'---------------------------------------------------------------------------------------
function loadDatosCargas(pCdAviso, byref pCdBuque, byref pDsBuque, ByRef firstProduct, byref pListOfProducts)
dim rs, strSQL, rtrn, index
rtrn = false
pListOfProducts = "0"
if (pCdAviso <> EMBARQUE_ESPECIAL_AJUSTES) then
    strSQL =	" SELECT EMB.CDAVISO, EMB.CDBUQUE, EMB.DSBUQUE, ED.CDPRODUCTO, PR.DSPRODUCTO FROM " & _
		        "	(" & _
		        "	SELECT 1 AS ETAPA, EMB.CDAVISO, BUQ.CDBUQUE, BUQ.DSBUQUE " & _
		        "		FROM EMBARQUES EMB INNER JOIN BUQUES BUQ ON EMB.CDBUQUE=BUQ.CDBUQUE WHERE EMB.CDAVISO=" & pCdAviso & _
		        "	UNION " & _
		        "	SELECT 2 AS ETAPA, EMB.CDAVISO, BUQ.CDBUQUE, BUQ.DSBUQUE " & _
		        "		FROM HEMBARQUES EMB INNER JOIN BUQUES BUQ ON EMB.CDBUQUE=BUQ.CDBUQUE WHERE EMB.CDAVISO= " & pCdAviso & _
		        "	) AS EMB" & _
			    " INNER JOIN " & _
		        "	EMBARQUESDATOS ED ON ED.CDAVISO=EMB.CDAVISO " & _
		        " INNER JOIN " & _
		        "	PRODUCTOS PR ON ED.CDPRODUCTO=PR.CDPRODUCTO " & _
			    " GROUP BY EMB.CDAVISO, EMB.CDBUQUE, EMB.DSBUQUE, ED.CDPRODUCTO, PR.DSPRODUCTO ORDER BY PR.DSPRODUCTO ASC"			    
    Call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
else
    'Se indico el aviso especial para ajustes, se trae el buque especial y la lista completa de productos.
    'Se hace un FULL JOIN entre buques y productos por que solo se necesita el buque indicado y la lista completa de productos.
    strSQL= "Select * from BUQUES, PRODUCTOS" &_
            " where CDBUQUE=" & EMBARQUE_ESPECIAL_AJUSTES            
    Call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)    
end if    
    
if not rs.eof then 
    pCdBuque = rs("CDBUQUE")
    pDsBuque = rs("DSBUQUE")
    rtrn = true	
end if
index = 0
while not rs.EOF
    index = index + 1
	    if ((firstProduct="") or (firstProduct=0)) then firstProduct = rs("CDPRODUCTO")
	    if len(pListOfProducts) > 0 then pListOfProducts = pListOfProducts & ","
	    pListOfProducts = pListOfProducts & rs("CDPRODUCTO")
    rs.movenext
wend
loadDatosCargas = rtrn	
end function
'---------------------------------------------------------------------------------------
function loadDatosCargasKilos(pCdAviso, pCdProducto, pCosecha, byref pKilosInformados)
dim rs, strSQL, rtrn
rtrn = false
pKilosInformados = 0

strSQL = "SELECT SUM(KILOS) AS KILOSINFORMADOS FROM EMBARQUESDATOS WHERE CDAVISO=" & pCdAviso & " AND CDPRODUCTO=" & pCdProducto ' & " AND CDCOSECHA=" & pCosecha
'Response.Write strSQL 
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
if not rs.eof then
	    if not isnull(rs("KILOSINFORMADOS")) then  pKilosInformados = rs("KILOSINFORMADOS")
	    rtrn = true
end if
    
loadDatosCargasKilos = rtrn	
end function
'---------------------------------------------------------------------------------------
function loadKilosCargados(pCdAviso, pCdProducto, pCdCosecha, byref pKilosCargados)
dim rs, strSQL, rtrn
rtrn = false
pKilosCargados = 0
strSQL =	" SELECT SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS CTG WHERE CDAVISO=" & pCdAviso & " AND CDPRODUCTO=" & pCdProducto & _
			" GROUP BY CDAVISO, CDPRODUCTO "
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
if not rs.EOF then
	if not isnull(rs("KILOSCARGADOS")) then pKilosCargados = rs("KILOSCARGADOS")
	rtrn = true
end if	
loadKilosCargados = rtrn	
end function
'---------------------------------------------------------------------------------------
Sub MostrarHistorico()
Dim rs, pQuery
pQuery =" SELECT CDAVISO, CDCOSECHA, SECUENCIA, B.DSBUQUE, C.DSCLIENTE, P.DSPRODUCTO, SUM(KILOSNETOS) KILOSNETOS, B.CDBUQUE, C.CDCLIENTE, P.CDPRODUCTO " & _
		"	FROM CTGEMBARCADOS CTG " & _
		"		LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = CTG.CDPRODUCTO " & _
		"		LEFT JOIN CLIENTES C ON C.CDCLIENTE = CTG.CDCLIENTE " & _
		"		LEFT JOIN BUQUES B ON B.CDBUQUE = CTG.CDBUQUE " & _
		" GROUP BY CDAVISO, B.CDBUQUE, B.DSBUQUE, C.CDCLIENTE, C.DSCLIENTE, P.CDPRODUCTO, P.DSPRODUCTO, CDCOSECHA, SECUENCIA  " & _
		" ORDER BY CDAVISO DESC, CDCLIENTE ASC, CDPRODUCTO ASC, SECUENCIA ASC, CDCOSECHA ASC"
'Response.Write pQuery		
call CargarHistoricoCTG(pQuery, 3)
End Sub
'---------------------------------------------------------------------------------------
Sub CargarHistorico(cdAviso, secuencia)
Dim rs, pQuery, MsjEmbarcados, pKilosSobrantes, myKilosNetos, myAnd
myAnd = ""
if secuencia <> "" then 
	myAnd = " AND SECUENCIA IN (" & secuencia & ")"
end if	
pKilosSobrantes = 0
pQuery = "SELECT SUM(KILOSNETOS) KILOSNETOS FROM CTGEMBARCADOS CTG WHERE CDAVISO= " & cdAviso & myAnd 
call GF_BD_Puertos(puerto, rs, "OPEN",pQuery)
if not rs.EOF then myKilosNetos = rs("KILOSNETOS")
pQuery =	" SELECT (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE, IDCAMION, CDCHAPACAMION, CDCHAPAACOPLADO, NUCARTAPORTE AS CARTAPORTE, CTG, C.CDCLIENTE, C.DSCLIENTE, P.CDPRODUCTO, P.DSPRODUCTO, KILOSNETOS " & _
			"	FROM CTGEMBARCADOS CTG " & _
			"		LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = CTG.CDPRODUCTO " & _
			"		LEFT JOIN CLIENTES C ON C.CDCLIENTE = CTG.CDCLIENTE " & _
			"	WHERE CDAVISO= " & cdAviso & myAnd & _
			"	ORDER BY DTCONTABLE ASC, CDPRODUCTO, CDCLIENTE"
'Response.Write pQuery
MsjEmbarcados = KilosEmbarcados(pQuery, myKilosNetos, pKilosSobrantes)
End Sub
'---------------------------------------------------------------------------------------
Sub CargarHistoricoCTG(SQLQuery, pCantidad)
    Dim rs, NColumnas, acumuladoKN, acumuladoSEC, acumuladoSecCliente, acumuladoSecClienteProducto, acumuladoKNCliente, acumuladoKNClienteProducto, myTrAviso, myTrSecuencias
    Dim Tag, tagx, avisoAnterior, secuenciaAnterior
    dim cdProducto, dsProducto, cdCliente, kilosNetos, auxTextCliente, auxTextClienteProducto
    dim clienteAnterior, productoAnterior, cosechaAnterior
    dim	auxBorderCliente, auxBorderProducto, auxBorderSecuencia, auxBorderCosecha
    avisoAnterior = 0
    clienteAnterior = 0
    productoAnterior = 0 
    cosechaAnterior = 0
    secuenciaAnterior = -1
    'Response.Write  SQLQuery
	call GF_BD_Puertos(puerto, rs, "OPEN",SQLQuery)
    while not rs.EOF
		if cLng(rs("CDAVISO")) <> avisoAnterior then
			if avisoAnterior <> 0 then
				if acumuladoSEC <> 0 then
					if acumuladoKNCliente <> 0 then
						'Imprime celda con total de secuencia anterior de cliente y cliente producto
						auxTextCliente = "<td style='BORDER-BOTTOM: #000000 1px solid;' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNCliente)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('C'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNCliente & ",'" & acumuladoSecCliente & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
						auxTextClienteProducto = "<td style='' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNClienteProducto)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('P'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNClienteProducto & ",'" & acumuladoSecClienteProducto & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
						acumuladoKNCliente = 0
						acumuladoSecCliente = ""
						acumuladoKNClienteProducto = 0
						acumuladoSecClienteProducto = ""						
					else
						auxTextCliente = ""
					end if
					'Imprime celda con total de secuencia anterior
					myTrSecuencias =   myTrSecuencias & "<td style='" & auxBorderCosecha & "' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoSEC)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('S'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoSEC & ",'" & secuenciaAnterior & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>" & auxTextClienteProducto & auxTextCliente & "</tr>" 
					acumuladoSEC = 0
				else
					myTrSecuencias =   myTrSecuencias & " <td style='" & auxBorderCosecha & "' align=center>&nbsp;</td></tr>" 		
				end if
				'Cierre TR aviso al cual le faltaban los kilos e imprime total del aviso
				myTrAviso = myTrAviso & acumuladoKN & ",'')" & chr(34) & " align=right>"  & GF_EDIT_DECIMALS(cdbl(acumuladoKN)*100,2) & "</td></tr>"
				myTableHTML = myTableHTML & myTrAviso & myTrSecuencias & _
					"			<tr style='BORDER-TOP: #000000 1px solid;' class='reg_Headers_navdos'>" & _
					"				<td colspan=8 align=right>" & _
					"				<b>" & GF_Traducir("Total Kg. embarcados") & "</b>" & _
					"				<b>" & GF_EDIT_DECIMALS(cdbl(acumuladoKN)*100,2) & " Kg.</b><img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('A'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKN & ",'')" & chr(34) & " title='Ver Carga Total' src='../images/see_all-16x16.png'></td>" & _
					"			</tr> " & _
					"		</table></td></tr></div>"
				acumuladoKN = 0
				acumuladoSEC = 0
				myTrSecuencias = ""
				acumuladoKNClienteProducto = 0
				acumuladoKNCliente = 0
			end if		
			'Imprime TR de aviso
			myTrAviso =		"<tr class='reg_Header_navdos' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>" & _
							"	<td align=center><img onclick='verSecuencias(TBL_" & rs("CDAVISO") & ", this)' src='../images/mas.gif'></td>" & _
							"	<td align=center>" & rs("CDAVISO") & "</td>" & _
							"	<td align=left>" & rs("DSBUQUE") & "</td>" & _
							"	<td onclick=" & chr(34) & "verHistorico('A'," & rs("CDAVISO") & "," & chr(39) & rs("CDPRODUCTO") & "//" & trim(rs("DSPRODUCTO")) & chr(39) & "," & trim(rs("CDCLIENTE")) & ","
			'Imprime TR de secuencia
			myTrSecuencias =   myTrSecuencias & _
							"<div><tr>" & _
							"	<td align=center colspan=6>" & _
							"		<table border=1 rules='cols' cellpadding=1 cellspacing=0 style='position:absolute;visibility:hidden;' class=reg_Heaaders width='95%' id=TBL_" & rs("CDAVISO") & ">" & _
							"			<tr class='reg_Header_nav'>" & _
							"				<td width='20%' align=center>" & GF_Traducir("Cliente") & "</td>" & _
							"				<td width='10%' align=center>" & GF_Traducir("Producto") & "</td>" & _
							"				<td width='2%'  align=center>" & GF_Traducir("Sec") & "</td>" & _
							"				<td width='8%'  align=center>" & GF_Traducir("Cosecha") & "</td>" & _
							"				<td width='15%' align=center>" & GF_Traducir("Kilos Netos") & "</td>" & _
							"				<td width='15%' align=center>" & GF_Traducir("Total Secuencia") & "</td>" & _
							"				<td width='15%' align=center>" & GF_Traducir("Total Producto") & "</td>" & _
							"				<td width='15%' align=center>" & GF_Traducir("Total Cliente") & "</td>" & _
							"			</tr> " & _  
							"			<tr class='reg_Headers_navdos'>" & _
							"				<td align=left>" & rs("DSCLIENTE") & "</td>" & _
							"				<td align=center>" & rs("DSPRODUCTO") & "</td>" & _
							"				<td align=center>" & rs("SECUENCIA") & "</td>" & _
							"				<td align=center>" & rs("CDCOSECHA") & "</td>" & _
							"				<td align=right>" & GF_EDIT_DECIMALS(cdbl(rs("KILOSNETOS"))*100,2) & " Kg.</td>"
			acumuladoSecClienteProducto = cLng(rs("SECUENCIA"))							
			acumuladoSecCliente = cLng(rs("SECUENCIA"))
		else
			'Cierre de la fila anterior
			if cLng(rs("SECUENCIA")) <> secuenciaAnterior then
				if secuenciaAnterior <> -1 then
					if cLng(rs("CDCLIENTE")) <> clienteAnterior then
						auxTextClienteProducto = "<td style='BORDER-BOTTOM: #000000 1px solid;' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNClienteProducto)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('P'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNClienteProducto & ",'" & acumuladoSecClienteProducto  & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
						auxTextCliente = "<td style='BORDER-BOTTOM: #000000 1px solid;' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNCliente)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('C'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNCliente & ",'" & acumuladoSecCliente & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
						acumuladoKNCliente = 0
						acumuladoSecCliente = ""
						acumuladoKNClienteProducto = 0
						acumuladoSecClienteProducto = ""
					else
						if trim(rs("DSPRODUCTO")) <> trim(productoAnterior) then
							auxTextClienteProducto = "<td style='BORDER-BOTTOM: #000000 1px dotted;' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNClienteProducto)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('P'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNClienteProducto & ",'" & acumuladoSecClienteProducto & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
							acumuladoKNClienteProducto = 0
							acumuladoSecClienteProducto = ""
						else
							auxTextClienteProducto = "<td align=right>&nbsp;</td>"
						end if
						auxTextCliente = "<td style='' align=right>&nbsp;</td>"
					end if
					myTrSecuencias =   myTrSecuencias & "<td style='" & auxBorderCosecha & "' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoSEC)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('S'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoSEC & ",'" & secuenciaAnterior & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>" & auxTextClienteProducto & auxTextCliente & "</tr>" 
				else
					myTrSecuencias =   myTrSecuencias & " <td style='" & auxBorderCosecha & "' align=center>&nbsp;</td></tr>" 				
				end if
			else
					myTrSecuencias =   myTrSecuencias & " <td style='" & auxBorderCosecha & "' align=center>&nbsp;</td></tr>" 				
			end if
			'Fin Cierre de la fila anterior		
			
			if cLng(rs("CDCLIENTE")) <> clienteAnterior then
				auxTextCliente = rs("DSCLIENTE")
				productoAnterior = ""
				auxBorderCliente = " BORDER-TOP: #000000 1px solid;"
				auxBorderProducto = " BORDER-TOP: #000000 1px solid;"
				auxBorderSecuencia = " BORDER-TOP: #000000 1px solid;"
				auxBorderCosecha = " BORDER-TOP: #000000 1px solid;"
			else
				auxTextCliente = "&nbsp;"	
				auxBorderCliente = ""
				auxBorderProducto = ""
				auxBorderSecuencia = ""
				auxBorderCosecha = ""
			end if
			
			if trim(rs("DSPRODUCTO")) <> trim(productoAnterior) then
				auxTextProducto = rs("DSPRODUCTO")
				auxBorderProducto = " BORDER-TOP: #000000 1px dotted;"
				auxBorderSecuencia = " BORDER-TOP: #000000 1px dotted;"
				auxBorderCosecha = " BORDER-TOP: #000000 1px dotted;"				
			else
				auxTextProducto = "&nbsp;"
				auxBorderProducto = ""
				auxBorderSecuencia = ""
				auxBorderCosecha = ""
			end if		
	
			if cLng(rs("SECUENCIA")) <> secuenciaAnterior then
				if secuenciaAnterior <> -1 then
					acumuladoSEC = 0
					auxTextSecuencia = rs("SECUENCIA")
				end if
				auxBorderSecuencia = " BORDER-TOP: #000000 1px dotted;"
				auxBorderCosecha = " BORDER-TOP: #000000 1px dotted;"
			else
				auxTextSecuencia = "&nbsp;"
				auxBorderSecuencia = ""
				auxBorderCosecha = ""
			end if
			myTrSecuencias =   myTrSecuencias & _
							"			<tr style='" & auxBorderCliente & "' class='reg_Headers_navdos'>" & _
							"				<td style='" & auxBorderCliente & "' align=left>" & auxTextCliente & "</td>" & _
							"				<td style='" & auxBorderProducto & "' align=center>" & auxTextProducto & "</td>" & _
							"				<td style='" & auxBorderSecuencia & "' align=center>" & auxTextSecuencia & "</td>" & _
							"				<td style='" & auxBorderCosecha & "' align=center>" & rs("CDCOSECHA") & "</td>" & _							
							"				<td style='" & auxBorderCosecha & "' align=right>" & GF_EDIT_DECIMALS(cdbl(rs("KILOSNETOS"))*100,2) & " Kg.</td>"
		end if	
		
		kilosNetos = CLng(rs("KILOSNETOS"))
		acumuladoKN = acumuladoKN + kilosNetos
		acumuladoSEC = acumuladoSEC + kilosNetos
		acumuladoKNCliente = acumuladoKNCliente + kilosNetos
		acumuladoKNClienteProducto = acumuladoKNClienteProducto + kilosNetos
				

		secuenciaAnterior = cLng(rs("SECUENCIA"))
		if len(acumuladoSecCliente)>0 then acumuladoSecCliente = acumuladoSecCliente & ","
		acumuladoSecCliente = acumuladoSecCliente & cLng(rs("SECUENCIA"))
		if len(acumuladoSecClienteProducto)>0 then acumuladoSecClienteProducto = acumuladoSecClienteProducto & ","
		acumuladoSecClienteProducto = acumuladoSecClienteProducto & cLng(rs("SECUENCIA"))
		

		avisoAnterior = cLng(rs("CDAVISO")) 
		
		clienteAnterior = rs("CDCLIENTE")
		productoAnterior = rs("DSPRODUCTO")
		cosechaAnterior = rs("CDCOSECHA")	
		
		cdProducto = rs("CDPRODUCTO")
		dsProducto = rs("DSPRODUCTO")
		cdCliente = rs("CDCLIENTE")

        rs.MoveNext
    wend
    rs.Close
    Set rs = Nothing



	'Imprimir aumulados cargados
	if acumuladoSEC <> 0 then
		if acumuladoKNCliente <> 0 then
			auxTextCliente = "<td style='BORDER-BOTTOM: #000000 1px solid;' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNCliente)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('C'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNCliente & ",'" & acumuladoSecCliente & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
			auxTextClienteProducto = "<td style='' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoKNClienteProducto)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('P'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKNClienteProducto & ",'" & acumuladoSecClienteProducto & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>"
			acumuladoKNCliente = 0
			acumuladoSecCliente = ""
		else
			auxTextCliente = ""
		end if
		myTrSecuencias =   myTrSecuencias & "<td style='" & auxBorderCosecha & "' align=right>" & GF_EDIT_DECIMALS(cdbl(acumuladoSEC)*100,2) & " Kg.<img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('S'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoSEC & ",'" & secuenciaAnterior & "')" & chr(34) & " title='Ver detalle' src='../images/see_all-16x16.png'></td>" & auxTextClienteProducto & auxTextCliente & "</tr>" 
		acumuladoSEC = 0
	else
		myTrSecuencias =   myTrSecuencias & " <td style='" & auxBorderCosecha & "' align=center>&nbsp;</td></tr>" 				
	end if	
	acumuladoSEC = 0
	if acumuladoKN <> 0 then
		myTrAviso = myTrAviso & acumuladoKN & ",'')" & chr(34) & " align=right>"  & GF_EDIT_DECIMALS(cdbl(acumuladoKN)*100,2) & "</td></tr>"
		myTableHTML = myTableHTML & myTrAviso & myTrSecuencias & _
			"			<tr style='BORDER-TOP: #000000 1px solid;' class='reg_Headers_navdos'>" & _
			"				<td colspan=8 align=right>" & _
			"				<b>" & GF_Traducir("Total Kg. embarcados") & "</b>" & _
			"				<b>" & GF_EDIT_DECIMALS(cdbl(acumuladoKN)*100,2) & " Kg.</b><img style='cursor:pointer;' onclick=" & chr(34) & "verHistorico('A'," & avisoAnterior & "," & chr(39) & cdProducto & "//" & trim(dsProducto) & chr(39) & "," & trim(cdCliente) & "," & acumuladoKN & ",'')" & chr(34) & " title='Ver Carga Total' src='../images/see_all-16x16.png'></td>" & _
			"			</tr> " & _
			"		</table></td></tr></div>"	

		acumuladoKN = 0
	end if		    
end sub
'---------------------------------------------------------------------------------------
%>
<HTML>
<HEAD>
   <TITLE>Reporte de CTGs Embarcados</TITLE>
</HEAD>
<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
	var toolBarGrupos = new Toolbar("toolBarGrupos",5);	
	<%select case accion%>
		<%case "HIS2"%>
			toolBarGrupos.addButton("Previous-16x16.png", "Volver", "submitInfo('HIST');");	
			toolBarGrupos.addButton("Excel-16x16.png", "Exportar", "createXLS('H');");	
		<%case "HIST"%>
			toolBarGrupos.addButton("cancel-16x16.png", "Cancelar", "cancelar();");	
		<%case "SAVE"%>
			toolBarGrupos.addButton("Excel-16x16.png", "Exportar", "createXLS('H');");	
			toolBarGrupos.addButton("Historico-16x16.png", "Historico", "clearAndSubmit();");	
		<%case else%>			
			toolBarGrupos.addButton("accept-16x16.png", "Confirmar", "submitInfo('SAVE');");	
			toolBarGrupos.addButton("Historico-16x16.png", "Historico", "clearAndSubmit();");	
			toolBarGrupos.addButton("Excel-16x16.png", "Exportar", "createXLS('N');");	
			toolBarGrupos.addButton("Stock-16x16.png", "Stock", "loadPopUpDisponibilidad();");	
	<%end select%>		
	function loadPopUpDisponibilidad() {				
		var puw = new winPopUp('popupDisponibles','CtgDisponibles.asp?Pto=<%=puerto%>','500','420','CTGs Disponibles', '');
	}	
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function verSecuencias(pElement, pImg){
		//alert(pImg)
		if (pElement.style.visibility == "hidden") {
			pElement.style.visibility = "visible";
			pElement.style.position = "relative";
			pImg.src = "../images/menos.gif"
		}
		else{
			pElement.style.visibility = "hidden";
			pElement.style.position = "absolute";
			pImg.src = "../images/mas.gif"
		}
	}
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function submitInfo(pAccion){
		if (pAccion=='SAVE' || pAccion=='SEARCH'){
			if (document.getElementById("results")){
				document.getElementById("results").innerHTML = "";
			}
			if (document.getElementById("loading")){
				document.getElementById("loading").style.visibility = "visible";
				document.getElementById("loading").style.position = "relative";
			}
		}
		document.getElementById("accion").value = pAccion;
		document.form1.submit();
	}
	function createXLS(pOption){
		if (pOption == "H"){
			window.open("reporteCTGEmbarcadosXLS.asp?cdProducto=<%=cdProducto%>&dsProducto=<%=dsProducto%>&cdCliente=<%=cdCliente%>&cdAviso=<%=cdAviso%>&accion=HIS2&kilos=<%=kilos%>&Pto=<%=puerto%>&secuencia=<%=secuencia%>&tipo=<%=tipo%>");
		}	
		else{
			window.open("reporteCTGEmbarcadosXLS.asp?cdProducto=<%=cdProducto%>&dsProducto=<%=dsProducto%>&cdCliente=<%=cdCliente%>&cdAviso=<%=cdAviso%>&accion=PREV&kilos=<%=kilos%>&Pto=<%=puerto%>&secuencia=<%=secuencia%>&tipo=<%=tipo%>");		
		}
		
	}
	function verHistorico(pTipo, pCdAviso, pCdProducto, pCdCliente, pKilos, pSec){
		document.form1.action = "reporteCTGEmbarcados.asp?cdAviso=" + pCdAviso + "&cdAvisoAnt=" + pCdAviso + "&cdProducto=" + pCdProducto + "&cdCliente=" + pCdCliente + "&kilos=" + pKilos + "&secuencia=" + pSec;
		document.getElementById("accion").value = "HIS2";
		document.getElementById("tipo").value = pTipo;
		document.form1.submit();
	}
	function cancelar(){
		document.form1.submit();
	}
	function clearAndSubmit(){
		submitInfo('HIST');
	}
	function bodyOnLoad(){
		toolBarGrupos.draw();
	}
</script>	
<BODY onload="bodyOnLoad()">

<div align="center">
   <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
	     <td width="90%"><b><font class="Birthday">Reporte de CTGs embarcados</font></b></td>
	     <td align="center"><img SRC="../Images/Iconos/Barco64x64.png"></td>
      </tr>
      <tr>
         <td>&nbsp;</td>
	     <td align="center" valign="top"><b><font name="Times New Roman" face="Verdana" size="4"><%Response.Write(puerto)%></font></b></td>
      </tr>
   </table>
</div>
<br>
<div id="toolBarGrupos"></div>
<br>
<FORM action="reporteCTGEmbarcados.asp" method=POST id=form1 name=form1>

	<table id="tblBusqueda" width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="../images/marcos/marco_r1_c1.gif"></td>
           <td width="25%"><img src="../images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="../images/marcos/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="../images/marcos/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
           <td width="8"><img src="../images/marcos/marco_r2_c3.gif"></td>
           <td align="right">
           		<% 
           		if not KilosAlcanzado then
					Response.Write "<font color=red><b>ATENCION: No se pudo alcanzar la cantidad de kilos solicitada</b></font>"
				End If
				%>
           </td>
           <td></td>
       </tr>
       <tr>
           <td><img src="../images/marcos/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="../images/marcos/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="../images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="../images/marcos/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="../images/marcos/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="100%" align="center" border="0">
                            <tr>
								<td width="12%" align="right"><% = GF_TRADUCIR("Aviso") %>:</td>
								<td width="38%">
									<input type="text" SIZE="3" MAXLENGTH="5" id="cdAviso" name="cdAviso" value="<% =cdAviso %>" <%=controlsState%>>
									<input type="hidden" SIZE="3" MAXLENGTH="5" id="cdAvisoAnt" name="cdAvisoAnt" value="<% =cdAvisoAnt %>">
								</td>
                                <td width="12%" align="right"><% =GF_TRADUCIR("Buque") %>:</td>
                                <td width="38%">                                
										<b><%=cdBuque & "-" & dsBuque%> </b>
                                </td>	
                            </tr>                     
							<tr>
								<td align="right"><% = GF_TRADUCIR("Producto") %>:</td>
								<td>
									<select style="z-index:-1;" id="cdProducto" onchange="submitInfo('CMB')" name="cdProducto" <%=controlsState%>>
										<%
										strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS WHERE CDPRODUCTO IN(" & listOfProducts & ") ORDER BY DSPRODUCTO"
										Call executeQueryDb(puerto, Rs, "OPEN", strsql)										
										while not rs.eof 
											if cint(cdProducto) = cint(rs("CDPRODUCTO")) then
												mySelected = "SELECTED"
											else
												mySelected = ""
											end if	
												%>
												<option value="<%=rs("CDPRODUCTO") & "//" & rs("DSPRODUCTO")%>" <%=mySelected%>><%=rs("DSPRODUCTO")%></option>
												<%			
											rs.movenext
										wend	
											%>							
									</select>
								</td>
								<td align="right"><%=GF_TRADUCIR("Cosecha") %></td>
								<td align="left">
									<select style="z-index:-1;" onchange="submitInfo('CMB')" name="cdCosecha" <%=controlsState%>>
										<option value="0" ><%=GF_Traducir("Cualquiera...")%></option>
										<% 
										if len(listOfHarvest) > 1 then
											myDescomposicion = split(listOfHarvest,";")
											for i=0 to ubound(myDescomposicion)
												if clng(myDescomposicion(i)) = cdCosecha then 
													mySelected = "SELECTED"
												else
													mySelected = ""
												end if	
												%>
													<option value="<%=myDescomposicion(i)%>" <%=mySelected%>><%=myDescomposicion(i)%></option>
												<%
											next
										END IF	
										%>
									</select>	
								</td>
							</tr>	
							<tr>	
								<td align="right"><% = GF_TRADUCIR("Cliente") %>:</td>
								<td>
									<select style="z-index:-1;" name="cdCliente" <%=controlsState%>>
										<!--<option value=0><%=GF_Traducir("TODOS")%></option>-->
										<%
										strSQL = "SELECT CDCLIENTE, DSCLIENTE FROM dbo.CLIENTES ORDER BY DSCLIENTE"
											Call executeQueryDb(puerto, Rs, "OPEN", strsql)											
											while not rs.eof 
												if cdCliente = rs("CDCLIENTE") then
													mySelected = "SELECTED"
												else
													mySelected = ""
												end if												
													%>
													<option value="<%=rs("CDCLIENTE")%>" <%=mySelected%>><%=rs("DSCLIENTE")%></option>
													<%			
												rs.movenext
											wend	
										%>							
									</select>
								</td>
								<td align="right"><% = GF_TRADUCIR("Camiones de") %>:</td>
								<td>
									<select style="z-index:-1;" name="cdCamionesDe" <%=controlsState%>>
										<option value="0" ><%=GF_Traducir("Cualquiera...")%></option>
										<%
										if not isNull(rs) then
											rs.movefirst
											while not rs.eof 
												if cdCamionesDe = rs("CDCLIENTE") then
													mySelected = "SELECTED"
												else
													mySelected = ""
												end if												
													%>
													<option value="<%=rs("CDCLIENTE")%>" <%=mySelected%>><%=rs("DSCLIENTE")%></option>
													<%			
												rs.movenext
											wend	
										end if		
										%>							
									</select>
								</td>								
							</tr>
							<tr>
								<td align="right"><%=GF_TRADUCIR("Kilos") %>:</td>
								<td colspan="3">
									<input type="text" SIZE="10" id="kilos" name="kilos" value="<% =kilos %>" <%=controlsState%>> 
								</td>
							</tr>
                            <tr>
								<td colspan="2">
									<table width="80%" border=0>
										<tr>
											<td width="15%" align="LEFT"></td>
											<td width="40%" align="LEFT">
												<font color="green">
													<%=GF_Traducir("Carga Informada: ")%> 
												</font>
											</td>
											<td width="45%" align="right">
												<%=GF_EDIT_DECIMALS(cdbl(kilosInformados)*100,2)%> Kg.	
											</td>	
												
										</tr>
										<tr>
											<td align="LEFT"></td>
											<td>	
												<font color="green">
													<%=GF_Traducir("Carga Actual: ")%> 
												</font>
											</td>
											<td align="right">
												<%=GF_EDIT_DECIMALS(cdbl(kilosPrevios)*100,2)%> Kg.
											</td>
										</tr>
										<tr>
											<td align="LEFT"></td>
											<td>	
												<font color="red">
													<%=GF_Traducir("Carga Restante: ")%> 
												</font>
											</td>	
											<td align="right">
												<%=GF_EDIT_DECIMALS((cdbl(kilosInformados)-cdbl(kilosPrevios))*100,2)%> Kg.
											</td>	
										</tr>		
									</table>	
								</td>
								<td colspan="2" align="center">
									<input type="SUBMIT" value="Buscar..." id=cmdSearch name=cmdSearch onclick="submitInfo('SEARCH');">
								</td>	
                            </tr>								                            
                     </table>
	           </td>
	           <td height="100%"><img src="../images/marcos/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="../images/marcos/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="../images/marcos/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="../images/marcos/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	<br>
	<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
		<tr>
			<td colspan=3>
			<%
			if hayError() then 
				call showErrors()
			end if
			%>
			</td>
		</tr>	
	</table>			

    <INPUT type="hidden" id="Pto" name="Pto" value=<%= Request("Pto")%>>
    <INPUT type="hidden" id="accion" name="accion">
    <INPUT type="hidden" id="tipo" name="tipo" value=<%=tipo%>>
	<%
    if not hayError() then 
		%>
		<div id="loading" style="position:absolute;visibility:hidden;">
			<table border=0 align="center" width="90%" cellpadding=1 cellspacing=0>
				<tr>
					<td align="center">
						<img src="../images/loading_circle_green.gif"><br><%=GF_Traducir("Aguarde por favor...")%>
					</td>
				</tr>
			</table>	
		</div>	
		<table id="results" border=0 align="center" width="90%" class="reg_Header" cellpadding=1 cellspacing=0>
		<%
	    if myTableHTML = "" then
			
			if mySaveText = "" then
				Response.Write "<tr><td align=center>No se encontraron camiones</td></tr>"
			else
				Response.Write "<tr><td align=center><font size=4><b>" & mySaveText & "</b></font></td></tr>"
			end if	
		else	
			Response.Write myTableHTML
		end if	
		%>
		</table>	
		<%
    end if
    %>
</form>
</body>
</html>
