<%
'--------------------------- CAMARA ---------------------------
'***** Codigo del pagador *********'
CONST CAMARA_CODIGO_COMPRADOR = 1
'***** Muestra Lacrada *********
CONST CAMARA_MUESTRA_LACRADA  = "L"
CONST CAMARA_SERVICIO_LACRADO = "N"
'***** Codigos de Ensayo habilitados por la Camara *********
CONST CAMARA_ENSAYO_FISICO     = "F"
CONST CAMARA_ENSAYO_QUIMICO    = "Q"
CONST CAMARA_ENSAYO_S  = "S"
CONST CAMARA_ENSAYO_CONDICION  = "C"
CONST CAMARA_ENSAYO_G  = "G"
'***** Parametros de Camara  *********
Const CAMARA_PARAMETER_DIRECCIONPTO = "LEYENDADIRECCIONPTO" 'Direccion postal del puerto
Const CAMARA_PARAMETER_PLANTAONCCA = "LEYENDAPLANTAONCCA"
Const CAMARA_PARAMETER_ESTADOEGRESADO = "CDESTADOEGRESADO"
Const CAMARA_PARAMETER_ESTADOTARA = "CDESTADOTARA"
Const CAMARA_PARAMETER_MAIL_EXPORTACION = "MAILFILECAMARAEXPORT"
'***** Ruta origen y nombre de los archivos  *********
Const CAMARA_EXPORT_FILENAME_CABECERA		   = "SOLICI01.TXT"
Const CAMARA_EXPORT_FILENAME_ANALISIS		   = "SOLICI02.TXT"
Const CAMARA_EXPORT_FILENAME_CUENTAYORDEN      = "SOLICI03.TXT"
Const CAMARA_FILE_ERRORES			           = "ErroresImportacion" 
Const CAMARA_EXPORT_FILENAME_REPORTE           = "ReporteExportacion.TXT"
Const CAMARA_EXPORT_FILENAME_ERROR             = "SolicitudesReporteErrores.TXT"
Const CAMARA_EXPORT_TEMP_REPORTE               = "TemporalExportacion"
' ************************ Clases de analisis ************************
Const CAMARA_ANALISIS_COMUN = 1
Const CAMARA_ANALISIS_RECONSIDERACION = 4
Const CAMARA_EXPORTACION_ANALISIS = "EXPORTACION"
Const CAMARA_IMPORTACION_ANALISIS = "IMPORTACION"

Const CODIGO_CAMARA_ROSARIO = 2
Const CODIGO_CAMARA_BAHIA = 3

Const RESULTADO_NO_VALIDAR_VALOR = "NO_VAL"

Const ANALISIS_SEGMENT_SIZE = 10
Const XML_SOLICITUDES_TAG   = "Solicitud"
Const XML_RESULTADOS_TAG    = "EnsayoTecnica"

' ************************ Path del archivo ************************
Const FOLDER_FILE_ANALISIS = "Archivos"
Const FOLDER_FILE_HISTORICO = "Historico"
Const FILE_IMPORT_ANALISIS_02 = "Datos.XML"
Const FILE_IMPORT_ANALISIS_03 = "RESULTA.TXT"
' ************************ Codigos de Bolsa ************************
Const BOLSA_ARROYO = 90
Const BOLSA_TRANSITO = 91
Const BOLSA_PIEDRABUENA = 92
' *********************  Exportacion de archivos *******************
Const CAMARA_PARAMETER_DIRCAMARA = "DIREXPOCAMARA"
' *********************  Puerto en camara  *******************
Const CAMARA_PARAMETER_PUERTOENCAMARA = "PUERTOENCAMARA"
' *********************  Cuenta y Orden de intermediario  ******************
Const CUENTA_ORDEN_INTERMEDIARIO = 1
'**************  Producto habilitado para la rebaja convenida en Piedrabuena (exportacion de Analisis) **************
Const CAMARA_PARAMETER_PRODPROTEINA = "CDPRODPROTEINA" ' Proteina
'------------------------------------------------------------------------------------------------------------
'JAS - DEPRECATED
'Devuelve el numero de puerto para operar en la Camara(Solo para los puertos de Rosario)
'Function getNroPuertoForCamara(pPto)
'	Dim rtrn
'	rtrn = 0
'	Select Case UCase(pPto)
'		Case TERMINAL_TRANSITO
'			rtrn = 24
'		Case TERMINAL_ARROYO
'			rtrn = 63
'	End Select	
'	getNroPuertoForCamara = rtrn
'End Function 
'------------------------------------------------------------------------------------------------------------
'Devuelve los datos del Oncca para una determinada carta de porte y fecha
'	-Cuit y Razon Social del titular de la Carta de Porte
'	-Tipo Transporte
'	-Codigo de localidad destino
'	-Codigo de localidad procedencia
'	-Codigo de establecimiento
Function getDatosOncca(rsTit)
	Dim strSQL ,rtrn()
	redim rtrn(7)
	rtrn(0) = ""
	rtrn(1) = ""
	rtrn(2) = ""
	rtrn(3) = ""
	rtrn(4) = ""
	rtrn(5) = ""
	rtrn(6) = ""	
	if (not rsTit.Eof) then
		rtrn(0) = rsTit("NCUITREMITENTE")
		rtrn(1) = rsTit("CRAZONREMITENTE")
		rtrn(2) = ""
		if (isNumeric(rsTit("NCTIPOTRANSPORTE"))) then
			if (CInt(rsTit("NCTIPOTRANSPORTE")) = TIPO_TRANSPORTE_VAGON) then 			
				rtrn(2) = "V"
			else
				rtrn(2) = "C"
			end if
		end if
		rtrn(3) = rsTit("NCLOCDEST")
		rtrn(4) = rsTit("NCLOCPROCE")
		rtrn(5) = rsTit("ESTABLEPROCE")
		rtrn(6) = 0
		if (isNumeric(rsTit("NCTIPOTRANSPORTE"))) then rtrn(6) = CInt(rsTit("NCTIPOTRANSPORTE"))	
	end if	
	getDatosOncca = rtrn
End Function
'------------------------------------------------------------------------------------------------------------
Function validarCamposObligatorios(pAceptacion,pCuit1,pCuit2,pNroPtoCamara,pNeto,pDsVendedor,pCtaPorte,pArrOncca,pCtg,pCdPlantaOncca,pTipoTransporte)		 
	Dim msgErr 	
	msgErr = ""	
	if (CInt(pAceptacion) = 0) then msgErr = "Proceso interrumpido: No se encontraron datos del Codigo de aceptacion"	
	if (Trim(pCuit1) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del Cuit(comprador)"	
	if (Trim(pCuit2) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del Cuit(vendedor)"
	if (Cdbl(pNroPtoCamara) = 0) then msgErr = "Proceso interrumpido: No se encontraron datos Codigo de Puerto"
	if (Cdbl(pNeto) <= 0 ) then msgErr = "Proceso interrumpido: Se encontro un error en el peso Neto"
	if (Trim(pDsVendedor) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del Remitente comercial"
	if (Trim(pCtaPorte) = "") then msgErr = "Proceso interrumpido: No se encontraron datos de la Carta de Porte"
	if (pTipoTransporte = TIPO_TRANSPORTE_CAMION) then
        if (Trim(pCtg) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del CTG"
    end if
	if (Trim(pCdPlantaOncca) = "") then msgErr = "Proceso interrumpido: No se encontraron datos de la Planta de Oncca Destino"
	if (Trim(pArrOncca(0)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del Cuit del Titular de la Carta de Porte"
	if (Trim(pArrOncca(1)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos de la razon social del Titular de la Carta de Porte"
	if (Trim(pArrOncca(2)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del tipo de Transporte"
	if (Trim(pArrOncca(3)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos de la localidad Oncca Destino"
	if (Trim(pArrOncca(4)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos de la localidad Oncca Procedencia"
	'if (Trim(pArrOncca(5)) = "") then msgErr = "Proceso interrumpido: No se encontraron datos del establecimiento procedencia"
	validarCamposObligatorios = msgErr
end Function
'------------------------------------------------------------------------------------------------------------
Function getProductoCamara2ACTI(pCamara, pProducto)
    Dim rtrn, strSQL, rs
    
    rtrn = 0
    if (session("LAB_C" & pCamara & "_P" & pProducto) <> "") then
        rtrn = session("LAB_C" & pCamara & "_P" & pProducto )
    else
        strSQL="Select * from MERFL.MER119F1 where CAMAPE=" & pCamara & " and PRODPE=" & pProducto 
        Call executeQuery(rs, "OPEN", strSQL)
        if (not rs.eof) then rtrn = CLng(rs("EQUIPE"))        
    end if
    session("LAB_C" & pCamara & "_P" & pProducto) = rtrn
    getProductoCamara2ACTI = rtrn
    
End Function
'------------------------------------------------------------------------------------------------------------
Function getTipoAnalisisCamara2ACTI(pTipoAnalisis)
    Dim rtrn
    
    rtrn = 0
    Select Case (pTipoAnalisis)
        Case CAMARA_ENSAYO_FISICO
            rtrn = 3
        Case CAMARA_ENSAYO_QUIMICO
            rtrn = 4
        Case CAMARA_ENSAYO_CONDICION
            rtrn = 6
        Case CAMARA_ENSAYO_G
            rtrn = 7
        Case CAMARA_ENSAYO_S
            rtrn = 5            
    End Select
    getTipoAnalisisCamara2ACTI = rtrn    
End Function
'------------------------------------------------------------------------------------------------------------
Function getTipoAnalisisACTICamara(pTipoAnalisis)
    Dim rtrn
    rtrn = ""
    Select Case (pTipoAnalisis)
        Case 3
            rtrn = CAMARA_ENSAYO_FISICO
        Case 4
            rtrn = CAMARA_ENSAYO_QUIMICO
        Case 6
            rtrn = CAMARA_ENSAYO_CONDICION
        Case 7
            rtrn = CAMARA_ENSAYO_G
        Case 5
            rtrn = CAMARA_ENSAYO_S
    End Select
    getTipoAnalisisACTICamara = rtrn    
End Function
'------------------------------------------------------------------------------------------------------------
Function getPuertoCamara2ACTI(pCamara, pPuerto)
    Dim rtrn, strSQL, rs
    
    rtrn = 0
    if (isNumeric(pPuerto)) then
        if (session("LAB_C" & pCamara & "_PTO" & pPuerto) <> "") then
            rtrn = session("LAB_C" & pCamara & "_PTO" & pPuerto)
        else
            strSQL="Select DEACDE from MERFL.MER19AF1 where CAMADE=" & pCamara & " and DECADE=" & pPuerto
            Call executeQuery(rs, "OPEN", strSQL)
            if (not rs.eof) then rtrn = rs("DEACDE")
        end if
    end if        
    session("LAB_C" & pCamara & "_PTO" & pPuerto) = rtrn
    getPuertoCamara2ACTI = rtrn
    
End Function
'------------------------------------------------------------------------------------------------------------
Function getConceptoCamara2ACTI(pCamara, pProducto, pTipoAnalisis, pConcepto)
    Dim rtrn, strSQL, rs
    
    rtrn = 0
    if (session("LAB_C" & pCamara & "_P" & pProducto & "_T" & pTipoAnalisis & "_E" & pConcepto) <> "") then
        rtrn = session("LAB_C" & pCamara & "_P" & pProducto & "_T" & pTipoAnalisis & "_E" & pConcepto)
    else
        strSQL="Select * from MERFL.MER2E9F1 where CAMAAE=" & pCamara & " and PRODAE in (0, " & pProducto & ") and TIPAAE=" & pTipoAnalisis & " and ANCAAE=" & pConcepto & " order by PRODAE desc"
        Call executeQuery(rs, "OPEN", strSQL)
        if (not rs.eof) then 
            rtrn = CLng(rs("ANACAE"))
            'Conversiones especiales de codigos (tomados del progrmaa original MER59C)
            if ((CLng(rtrn) = 5) and (CInt(pProducto) = 23)) then rtrn = 59
            if ((CLng(rtrn) = 56) and (CInt(pProducto) = 20)) then rtrn = 62        
        end if
    end if        
    session("LAB_C" & pCamara & "_P" & pProducto & "_T" & pTipoAnalisis & "_E" & pConcepto) = rtrn
    getConceptoCamara2ACTI = rtrn
    
End Function
'------------------------------------------------------------------------------------------------------------
Function getProcedencia(pProc)
	Dim rtrn()
	redim rtrn(2)
	if (pProc <> "") then
		rtrn(0) = Mid(pProc,1,4)
		rtrn(1) = Mid(pProc,6,2)		
	end if
	getProcedencia = rtrn
End Function
'--------------------------------------------------------------------------------------------------------------------
Function validarCartaPorte(pCartaPorte, ByRef msg)
    Dim rtrn, strSQL, rs
    rtrn = false
    msg = ""
    strSQL="Select * from MERFL.MER311F6 where CPORR6=" & pCartaPorte
    Call executeQuery(rs, "OPEN", strSQL)
    if (not rs.eof) then 
        strSQL="Select * from MERFL.MER591CA where CPORCA=" & pCartaPorte
        Call executeQuery(rs, "OPEN", strSQL)
        if (rs.eof) then
            rtrn=true
        else
            msg = "La carta de porte ya tiene analisis registrados. (Carta de Porte:" & pCartaPorte & " | Certificado Archivo: " & g_certificado & " | Certificado AS400: " & rs("NROACA") & ")"                 
        end if
    else
        msg = "La carta de porte no se encuentra aplicada a ningun contrato. (Certificado: " & g_certificado & " | Carta de Porte:" & pCartaPorte & " | Muestra: " & g_muestra & ")" 
    end if
    validarCartaPorte = rtrn    
    
End Function
'------------------------------------------------------------------------------------------------------------
Function controlarCertificadoDuplicado(pCamara, pDestino, pProducto, pCertificado, pFechaAnalisis)
	Dim strSQL ,rsCer
	
	controlarCertificadoDuplicado = false		
	strSQL = "SELECT CPORCA FROM MERFL.MER591CA WHERE CDESCA=" & pDestino & " AND COBECA=" & pCamara & " AND CPROCA=" & pProducto & " AND NROACA=" & pCertificado & " AND FANACA=" & pFechaAnalisis
    Call executeQuery(rsCer, "OPEN", strSQL)    
    If Not rsCer.EOF Then controlarCertificadoDuplicado = true
End Function
'------------------------------------------------------------------------------------------------------------
Function existeEnsayoResultadosCamaraPuerto(pFechaDescarga,pEnsayo,pMuestra)
	Dim strSQL ,rsCer, myFechaDescarga
	
	myFechaDescarga = Left(pFechaDescarga, 4) & "-" & Mid(pFechaDescarga, 5, 2) & "-" & Right(pFechaDescarga, 2)
	existeEnsayoResultadosCamaraPuerto = false
	strSQL = "Select * from ResultadosCamara Where CdEnsayo='" & pEnsayo & "' and nuBArras ='" & pMuestra & "' and dtContable ='" & myFechaDescarga & "'"	    
    Call executeQueryDb(g_strPuerto, rsCer, "OPEN", strSQL)
    If Not rsCer.EOF Then existeEnsayoResultadosCamara = true
End Function
'------------------------------------------------------------------------------------------------------------
Function cargarGruposEnsayosCamiones(pIdCamion,pDtContable,pPto)
	Dim strsql,iMaxSqCalada,rsEns,rsEns1,i,myDtContable
	myDtContable = ""
	If (pDtContable <> "0") Then myDtContable = GF_FN2DTCONTABLE(pDtContable)
	
	If (myDtContable = "") Then
         strsql = "Select Max(sqCalada) as max From CaladaDecamiones where IdCamion ='"&pIdCamion&"' and IcTipoCalada = 'V'"
    Else
         strsql = "Select Max(sqCalada) as max From HCaladaDecamiones where IdCamion ='"&pIdCamion&"' and IcTipoCalada = 'V'  and DtContable = '"&myDtContable&"'"
    End If
    Call executeQueryDb(pPto, rsEns, "OPEN", strsql)	
	If Not rsEns.EOF Then 
		if not(IsNull(rsEns("max"))) then iMaxSqCalada = Clng(rsEns("max"))
	end if
	
	If myDtContable = "" Then
        strsql = "select case when g.cdgrupo is null then '' else Ltrim(g.cdgrupo) end as Grupo, '' as Ensayo  from gruposensayoscamiones g "
        strsql = strsql & " where g.IdCamion = '"&pIdCamion&"' and g.sqCalada= " & iMaxSqCalada
        strsql = strsql & " union all select '' as Grupo, case when e.cdensayo is null then '' else Ltrim(e.cdensayo) end as Ensayo from ensayoscamiones e "
        strsql = strsql & " where e.IdCamion = '"&pIdCamion&"' and e.sqCalada=" & iMaxSqCalada
    Else
        strsql = "select case when g.cdgrupo is null then '' else Ltrim(g.cdgrupo) end as Grupo, '' as Ensayo from hgruposensayoscamiones g "
        strsql = strsql & " where g.IdCamion = '"&pIdCamion&"' and g.dtContable = '"& myDtContable &"'"
        strsql = strsql & " and g.sqCalada= " & iMaxSqCalada
        strsql = strsql & " union all select '' as Grupo, case when e.cdensayo is null then '' else Ltrim(e.cdensayo) end as Ensayo from hensayoscamiones e "
        strsql = strsql & " where e.IdCamion = '"&pIdCamion&"' and e.dtContable = '"& myDtContable &"'"
        strsql = strsql & " and e.sqCalada=" & iMaxSqCalada
    End If
	Call executeQueryDb(pPto, rsEns1, "OPEN", strsql)
	i = 1
	While (not rsEns1.EOF)
        If (rsEns1("Grupo") <> "") Then
            if (not oDiccGruposEnsayosCamion.Exists(rsEns1("Grupo") & "|")) then oDiccGruposEnsayosCamion.Add rsEns1("Grupo") & "|", "k" & i
        Else
            if (not oDiccGruposEnsayosCamion.Exists("|" & rsEns1("Ensayo"))) then  oDiccGruposEnsayosCamion.Add "|" & rsEns1("Ensayo"), "k" & i			
        End If
        i = i + 1
        rsEns1.MoveNext()
	wend
End Function
'------------------------------------------------------------------------------------------------------------
Function cargarGruposEnsayosVagones(pCdVagon,pCartaPorte,pDtContable,pPto)
	Dim strsql,iMaxSqCalada,rsEns,rsEns1,i,myDtContable
	myDtContable = ""
	If (pDtContable <> "0") Then myDtContable = GF_FN2DTCONTABLE(pDtContable)
	
	If (myDtContable = "") Then
         strsql = "Select Max(sqCalada) as max From CALADADEVAGONES where nucartaPorte ='"&pCartaPorte&"' and cdvagon = '"& pCdVagon &"' and IcTipoCalada = 'A'"
    Else
         strsql = "Select Max(sqCalada) as max From HCALADADEVAGONES where nucartaPorte ='"&pCartaPorte&"' and cdvagon = '"& pCdVagon &"' and IcTipoCalada = 'A'  and DtContable = '"&myDtContable&"'"
    End If      
    Call executeQueryDb(pPto, rsEns, "OPEN", strsql)	
	If Not rsEns.EOF Then 
		if not(IsNull(rsEns("max"))) then iMaxSqCalada = Clng(rsEns("max"))
	end if
	
	If myDtContable = "" Then
        strsql = "select case when g.cdgrupo is null then '' else Ltrim(g.cdgrupo) end as Grupo, '' as Ensayo from GRUPOSENSAYOSVAGONES g "
        strsql = strsql & " where g.nucartaPorte = '"&pCartaPorte&"' and g.cdvagon = '"& pCdVagon &"' and g.sqCalada= " & iMaxSqCalada
        strsql = strsql & " union all select '' as Grupo, case when e.cdensayo is null then '' else Ltrim(e.cdensayo) end as Ensayo from ENSAYOSVAGONES e "
        strsql = strsql & " where e.nucartaPorte = '"&pCartaPorte&"' and e.cdvagon = '"&pCdVagon&"' and e.sqCalada=" & iMaxSqCalada
    Else
        strsql = "select case when g.cdgrupo is null then '' else Ltrim(g.cdgrupo) end as Grupo, '' as Ensayo from HGRUPOSENSAYOSVAGONES g "
        strsql = strsql & " where g.nucartaPorte = '"&pCartaPorte&"' and g.cdVagon = '"&pCdVagon&"' and g.dtContable = '"& myDtContable &"' and g.sqCalada= " & iMaxSqCalada
        strsql = strsql & " union all select '' as Grupo, case when e.cdensayo is null then '' else Ltrim(e.cdensayo) end as Ensayo from HENSAYOSVAGONES e "
        strsql = strsql & " where e.nucartaPorte = '"&pCartaPorte&"' and e.cdvagon = '"&pCdVagon&"' and e.dtContable = '"& myDtContable &"' and e.sqCalada=" & iMaxSqCalada
    End If        
	Call executeQueryDb(pPto, rsEns1, "OPEN", strsql)
	i = 1
	While (not rsEns1.EOF)
        If (rsEns1("Grupo") <> "") Then
            if (not oDiccGruposEnsayosCamion.Exists(rsEns1("Grupo") & "|")) then oDiccGruposEnsayosCamion.Add rsEns1("Grupo") & "|", "k" & i
        Else
            if (not oDiccGruposEnsayosCamion.Exists("|" & rsEns1("Ensayo"))) then  oDiccGruposEnsayosCamion.Add "|" & rsEns1("Ensayo"), "k" & i			
        End If
        i = i + 1
        rsEns1.MoveNext()
	wend
End Function
'------------------------------------------------------------------------------------------------------------
Function getDatosCartaPorte(pCartaPorte,ByRef pArray,pPto)
	Dim auxEstadoEgresado,strSQL,rsCP
	getDatosCartaPorte = false
	if (IsNumeric(pCartaPorte)) then
	    strSQL = "Select COALESCE(ca.IdCamion,'') IdCamion, (YEAR(ca.dtContable)*10000 + Month(ca.dtContable)*100 + DAY(ca.dtContable)) as dtContable, COALESCE(cc.NuBarras,'') NuBarras,COALESCE(cc.cdAceptacion,0) cdAceptacion, COALESCE(ca.cdChapacamion,'') cdChapacamion "&_
			     "from HCamionesDescarga cd "&_
			     "	  left join HCamiones ca on cd.IdCamion = ca.IdCamion and cd.DtContable = ca.DtContable "&_
			     "	  left join HCaladadeCamiones cc on cc.DtContable = cd.dtcontable and cc.IdCamion = cd.idcamion "&_
                 "where RIGHT(Rtrim(cd.nuCartaPorte),9) = '"& right(pCartaPorte, 9) & "'" &_
			     "		and cc.SqCalada = (Select Max(sqCalada) from HcaladadeCamiones where dtContable = cd.dtcontable and IdCamion = cd.idcamion) " &_			     
			     "      order by CA.DTEGRESO DESC"
	    Call executeQueryDb(pPto, rsCP, "OPEN", strSQL)
	    pArray(0) = ""
        pArray(1) = ""
        pArray(2) = "0"
        pArray(3) = 0
        pArray(4) = ""
	    if not rsCP.Eof then
		    pArray(0) = rsCP("IdCamion")
            pArray(1) = rsCP("dtContable")
            pArray(2) = Left(rsCP("NuBarras"), Len(rsCP("NuBarras"))-1)
            pArray(3) = rsCP("cdAceptacion")
            pArray(4) = rsCP("cdChapacamion")
            getDatosCartaPorte = true
	    end if	    
	end if
end function
'------------------------------------------------------------------------------------------------------------
Function CargarGruposEnsayosDef(pCdProducto, pAceptacion,pPto)
	 Dim strsql,rsEns,i
     strsql = "select case when cdgrupo is null then '' else Ltrim(cdgrupo) end as Grupo , '' as Ensayo from Defgrupos g "
     strsql = strsql & " where g.cdproducto = " & pCdProducto & " and g.cdAceptacion = " & pAceptacion
     strsql = strsql & " union all select '' as Grupo, case when e.cdensayo is null then '' else Ltrim(e.cdensayo) end as Ensayo from DefEnsayos e "
     strsql = strsql & " where e.cdproducto = " & pCdProducto & " and e.cdAceptacion= " & pAceptacion  
	Call executeQueryDb(pPto, rsEns, "OPEN", strsql)
	i = 1
	While (Not rsEns.EOF)
	    If (rsEns("Grupo") <> "" ) then
			if (not oDiccGruposEnsayos.Exists(rsEns("Grupo") & "|")) then oDiccGruposEnsayos.Add rsEns("Grupo") & "|", "k" & i			
		Else
			if (not oDiccGruposEnsayos.Exists("|" & rsEns("Ensayo"))) then oDiccGruposEnsayos.Add "|" & rsEns("Ensayo"), "k" & i
	    End If
		i = i + 1
	    rsEns.MoveNext
	wend
End Function
'------------------------------------------------------------------------------------------------------------
Function getGrupoCamara()
	Dim flagSeguir,rtrn,auxKey,arrayProductosGrupos,auxProd 
	'GRUPO DE PRODUCTOS HABILITADOS POR LA CAMARA PARA MUESTRA COMERCIAL
	rtrn = ""
		If (oDiccGruposEnsayosCamion.Count > 0) then
			for each key in oDiccGruposEnsayosCamion.Keys
				'obtengo el grupo guardado (1�:grupo - 2�:ensayo)
				auxKey = Split(key,"|")			
				if (Trim(auxKey(0)) <> "")then 
					rtrn = Trim(auxKey(0))
					Exit For
				end if	
			next
		else
			If (oDiccGruposEnsayos.Count > 0) then
				for each key in oDiccGruposEnsayos.Keys
					'obtengo el grupo guardado (1�:grupo - 2�:ensayo)
					auxKey = Split(key,"|")			
					if (Trim(auxKey(0)) <> "")then 
						rtrn = Trim(auxKey(0))
						Exit For
					end if
				next
			end if
		end if
	getGrupoCamara = rtrn
End Function
'------------------------------------------------------------------------------------------------------------
Function getEnsayosAnalisis(Byref objDic,pNroMuestra,pGrupo,pPto)
	Dim obj,iContadorC,auxKey,str,auxTipoEnsayo
	str = ""
	For Each obj In objDic
		auxKey = Split(obj,"|")
		if (Trim(auxKey(1)) <> "") then
			auxTipoEnsayo = Left(Trim(auxKey(1)),1)
			auxCdEnsayo   = Right(Trim(auxKey(1)),Len(Trim(auxKey(1)))-1)
			if (validarCdEnsayoCamara(auxTipoEnsayo,auxCdEnsayo)) then
				If (Not existEnsayoIncluido(Trim(pGrupo), Trim(auxKey(1)),pPto)) Then
					str = str & pNroMuestra &_
							    auxTipoEnsayo &_
								GF_nDigits(auxCdEnsayo,3) & "|"
				end if				
			end if
		end if
	Next
	if (Len(str) > 0) then str = Replace(left(str,len(str)-1),"|",vbcrlf)
	getEnsayosAnalisis = str
End Function
'------------------------------------------------------------------------------------------------------------
'Valida que el Ensayo(Tipo-Codigo) tengan el formato adecuado para la Camara
Function validarCdEnsayoCamara(pTipoEnsayo,pCdEnsayo)
	validarCdEnsayoCamara = false
	'Controlo que el tipo de ensayo este en el rango de los que pide la Camara
	if ((pTipoEnsayo = CAMARA_ENSAYO_FISICO)or(pTipoEnsayo = CAMARA_ENSAYO_QUIMICO)or(pTipoEnsayo = CAMARA_ENSAYO_CONDICION)) then		
		'Controlo que el codigo de ensayo tenga la cantidad de numeros que pide la Camara
		if (Len(pCdEnsayo) = 3) then validarCdEnsayoCamara = true
	end if
End Function
'------------------------------------------------------------------------------------------------------------
Function existEnsayoIncluido(pGrupo,pEnsayo,pPto)
    Dim rsEns,strsql
    existEnsayoIncluido = false
    If (pGrupo <> "") Then
        strsql = "Select * from EnsayosDeGrupos where Ltrim(cdgrupo) ='" & pGrupo & "' and Ltrim(cdensayo) ='" & pEnsayo & "'"
        Call executeQueryDb(pPto, rsEns, "OPEN", strsql)
        if (not rsEns.Eof) then existEnsayoIncluido = true        
    End If
End Function
'------------------------------------------------------------------------------------------------------------
Function getCuentayOrdenes(pIdCamion,pDtContable,pNroMuestra,pPto)
	Dim rsCyO,strSQL,rtrn
	rtrn = ""	
	if (pDtContable <> "0") then
        strsql = "Select c.sqOrden,Case when c.CdVendedor is null then 0 else c.CdVendedor end as CdVendedor, case when ve.DsVendedor is null then '' else Ltrim(ve.DsVendedor) end as DsVendedor,case when ve.NUDOCUMENTO is null then '' else Ltrim(ve.NUDOCUMENTO) end as NUDOCUMENTO, CASE when cl.NROSUCURSAL is null then 0 else cl.NROSUCURSAL end as CDSUCURSAL From HCuentayOrdenesCamiones c "
		strsql = strsql & "Left Join "
		strsql = strsql & "(Select 	 CDVENDEDOR,	DSVENDEDOR,	DSDOMICILIO,	NUTELEFONO,	CDTIPODOC,	NUDOCUMENTO,	DSOBSERVACIONES "        
        strsql = strsql & " from VENDEDORES) ve on ve.cdVendedor=c.cdVendedor "
        strsql = strsql & " left join TBLCAMARARELACIONCLIENTES re on re.idclientepto = ve.cdVendedor "
        strsql = strsql & " left join TBLCAMARACLIENTES cl on cl.idcliente = re.idclientecamara "
		strsql = strsql & " Where c.IdCamion = '"& pIdCamion &"' and c.DtContable = '" & GF_FN2DTCONTABLE(pDtContable) & "'"
	    strsql = strsql & " order by c.sqOrden"
    else
		strsql = "Select c.sqOrden,Case when c.CdVendedor is null then 0 else c.CdVendedor end as CdVendedor, case when ve.DsVendedor is null then '' else Ltrim(ve.DsVendedor) end as DsVendedor,case when ve.NUDOCUMENTO is null then '' else Ltrim(ve.NUDOCUMENTO) end as NUDOCUMENTO, CASE when cl.NROSUCURSAL is null then 0 else cl.NROSUCURSAL end as CDSUCURSAL From CuentayOrdenesCamiones c "
		strsql = strsql & "(Select 	 CDVENDEDOR,	DSVENDEDOR,	DSDOMICILIO,	NUTELEFONO,	CDTIPODOC,	NUDOCUMENTO,	DSOBSERVACIONES "	
        strsql = strsql & " from VENDEDORES) ve on ve.cdVendedor=c.cdVendedor "
		strsql = strsql & " left join TBLCAMARARELACIONCLIENTES re on re.idclientepto = ve.cdVendedor "
        strsql = strsql & " left join TBLCAMARACLIENTES cl on cl.idcliente = re.idclientecamara "
        strsql = strsql & " Where c.IdCamion = '"& pIdCamion &"'"
	    strsql = strsql & " order by c.sqOrden"
    end if    
    Call executeQueryDb(pPto, rsCyO, "OPEN", strSQL)
    while not rsCyO.Eof 
		If (Cdbl(rsCyO("CdVendedor")) <> 0) Then
			rtrn = rtrn & pNroMuestra &_
						  GF_nDigits(rsCyO("NUDOCUMENTO"),11) &_
						  GF_nChars(Left(rsCyO("DSVENDEDOR"),40),40," ",CHR_AFT) &_
						  GF_nDigits(rsCyO("CDSUCURSAL"),3)
		End If
		rsCyO.MoveNext()
		if not rsCyO.eof then 
			If (Cdbl(rsCyO("CdVendedor")) <> 0) Then rtrn = rtrn & vbcrlf
		end if	
    wend 
    getCuentayOrdenes = rtrn
End Function
'------------------------------------------------------------------------------------------------------------
Public Function getProximoSticker(pPto, dtContable, NuCartaPorte, tipoTransprote, idTransporte)
    
    Dim strDate, bok, sSticker, strSQL, rs, sDig	
    
	strDate = GF_FN2DTCONTABLE(Left(session("MmtoDato"), 8))
	If dtContable <> "" Then strDate = dtContable	
	
	bok = False
	sSticker = ""	
	
	strSQL = "Select STICKER from STICKERSCAMARA where DTCONTABLE='" & strDate & "' and NUCARTAPORTE='" & NuCartaPorte & "' and TIPOTRANSPORTE=" & tipoTransprote & " and IDTRANSPORTE='" & idTransporte & "'"
	Call executeQueryDb(pPto, rs, "OPEN", strSQL)
	If (Not rs.EOF) Then
		sSticker = CStr(rs("STICKER"))
	Else
		strSQL = "Insert into STICKERSCAMARA(DTCONTABLE, NUCARTAPORTE, TIPOTRANSPORTE,IDTRANSPORTE, DIGITOVERIFICADOR) values('" & strDate & "', '" & NuCartaPorte & "', " & tipoTransprote & ", '" & idTransporte & "', '')"
		Call executeQueryDb(pPto, rs, "EXEC", strSQL)
		strSQL = "Select STICKER from STICKERSCAMARA where DTCONTABLE='" & strDate & "' and NUCARTAPORTE='" & NuCartaPorte & "' and TIPOTRANSPORTE=" & tipoTransprote & " and IDTRANSPORTE='" & idTransporte & "'"
		Call executeQueryDb(pPto, rs, "OPEN", strSQL)
		If (Not rs.EOF) Then sSticker = CStr(rs("STICKER"))
	End If
	If (sSticker = "") Then
		err.Raise 1000, "DB2Ado", "No se ha podido asignar un nuevo sticker. Consulte con el administrador."
	Else
		sDigito = getDigitoVdorSticker(sSticker, pPto)
		sSticker = sSticker & sDigito
		sSQL = "Update STICKERSCAMARA set DigitoVerificador='" & sDigito & "' where DTCONTABLE='" & strDate & "' and NUCARTAPORTE='" & NuCartaPorte & "' and TIPOTRANSPORTE=" & tipoTransprote & " and IDTRANSPORTE='" & idTransporte & "'"
		Call executeQueryDb(pPto, rs, "EXEC", strSQL)
	End If
	getProximoSticker = sSticker
    
End Function    
'------------------------------------------------------------------------------------------------------------
Function getDigitoVdorSticker(sUltSticker, pPto)    
    Dim Dig, iTam, sngSuma, sDig, sngPar, sngImpar, i, sMultiplo, sngTotal
    Dim AuxSticker, sumDig, mod43
        
    AuxSticker = GF_nDigits(sUltSticker, 9)
    if (pPto = TERMINAL_PIEDRABUENA) then
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 1, 1)) * 4
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 2, 1)) * 3
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 3, 1)) * 2
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 4, 1)) * 7
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 5, 1)) * 6
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 6, 1)) * 5
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 7, 1)) * 4
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 8, 1)) * 3
            sngSuma = sngSuma + CInt(Mid(AuxSticker, 9, 1)) * 2
            Dig = 11 - (sngSuma Mod 11)
            sDig = Trim(CStr(Dig))
            sDig = Right(sDig, 1)
    Else
            'Nuevo calculo para UpRiver por Intacta
            
            'Sumar todos los digitos del codigo
            i = 1
            Do Until i > 9            
                sumDig = sumDig + CInt(Mid(AuxSticker, i, 1))                
                i = i + 1
            Loop
            'Obtener modulo 43
            mod43 = sumDig Mod 43            
            sDig = getCode39(mod43)
    End if
    getDigitoVdorSticker = sDig
End Function
'------------------------------------------------------------------------------------------------------------
Function getCode39(pDig)
    Dim asciiNumber
    
    If pDig >= 0 And pDig <= 9 Then
        asciiNumber = pDig + 48
    ElseIf pDig >= 10 And pDig <= 35 Then
        asciiNumber = pDig + 55
    ElseIf pDig = 36 Then
        asciiNumber = pDig + 9
    ElseIf pDig = 37 Then
        asciiNumber = pDig + 9
    ElseIf pDig = 38 Then
        asciiNumber = pDig - 6
    ElseIf pDig = 39 Then
        asciiNumber = pDig - 3
    ElseIf pDig = 40 Then
        asciiNumber = pDig + 7
    ElseIf pDig = 41 Then
        asciiNumber = pDig + 2
    ElseIf pDig = 42 Then
        asciiNumber = pDig - 5
    End If
    If asciiNumber > 0 Then
        getCode39 = Chr(asciiNumber)
    Else
        getCode39 = "?"
    End If
End Function
'------------------------------------------------------------------------------------------------------------
Function getDatosSticker(pBarra,ByRef pArr,pPto)
    Dim auxEstadoEgresado,rsSti
    getDatosSticker = False
    if (IsNumeric(pBarra)) then
        strSQL = "Select COALESCE(cc.IdCamion,'') IdCamion, (YEAR(cc.DtContable)*10000 + Month(cc.DtContable)*100 + DAY(cc.DtContable)) as dtContable,COALESCE(cc.cdAceptacion,0) cdAceptacion, "&_
			     "(Select COALESCE(cd.nuCartaPorte,'') nuCartaPorte from HCamionesDescarga cd where cd.DtContable = cc.DtContable and cd.IdCamion = cc.Idcamion) as cartaPorte ," &_
			     "(Select COALESCE(ca.cdChapaCamion,'') cdChapaCamion from HCamiones ca where ca.DtContable = cc.DtContable and ca.IdCamion = cc.Idcamion) as ChapaCamion "&_
			     " from HCaladadeCamiones cc "&_
			     " where cc.nuBarras= '"& pBarra &"'"&_
			     " and cc.sqCalada = (Select max(sqCalada) from HCaladadeCamiones where dtContable = cc.dtContable and idCamion =cc.idCamion)"	
	    Call executeQueryDb(pPto, rsSti, "OPEN", strSQL)
        If Not rsSti.EOF Then
            pArr(0) = Trim(rsSti("IdCamion"))
            pArr(1) = rsSti("dtContable")
            pArr(2) = Trim(rsSti("cartaPorte"))
            pArr(3) = rsSti("cdAceptacion")
            pArr(4) = Trim(rsSti("ChapaCamion"))
            getDatosSticker = True
        End If
    end if
End Function
'------------------------------------------------------------------------------------------------------------
Function armarSQLCabecera(pFechaDesde,pFechaHasta,pProducto,pIdCamion,pLstMuestras,pPto)
	Dim strSQL,myWhere, myWhereVagon, myWhereMuestra
	if (pFechaDesde <> "") then  
	    Call mkWhere(myWhere, "ca.DtContable", pFechaDesde, ">=", 3)
	    Call mkWhere(myWhereVagon, "ca.DtContableVagon", pFechaDesde, ">=", 3)
	end if
	if (pFechaHasta <> "") then  
	    Call mkWhere(myWhere, "ca.DtContable", pFechaHasta, "<=", 3)
	    Call mkWhere(myWhereVagon, "ca.DtContableVagon", pFechaHasta, "<=", 3)
    end if	    
	if (pProducto <> "")then	 
	    Call mkWhere(myWhere, "ca.cdProducto", pProducto, "=", 1)
	    Call mkWhere(myWhereVagon, "ca.cdProducto", pProducto, "=", 1)
	end if
	if (pIdCamion <> "") then	Call mkWhere(myWhere, "ca.idcamion", pIdCamion, "=", 3)    	
	if (pLstMuestras <> "") then
		if (myWhere = "") then 
			myWhere = " where "
		else
			myWhere = myWhere & " and "
		end if
		myWhere = myWhere & " SC.STICKER in (" & pLstMuestras & ")"
	end if
	strSQL = "Select T.*, "&_
			 "		 case when v.dsvendedor is null then '' else Rtrim(v.dsvendedor) end as dsvendedor, "&_
			 "		 case when v.NUDOCUMENTO is null then '' else v.NUDOCUMENTO end as NUDOCUMENTO, "&_
			 "       Case when v.CDSUCURSAL is null then 0 else v.CDSUCURSAL end as cdsucursalVendedor, "&_
			 "       Case when c.dsCliente is null then '' else Rtrim(c.dsCliente) end as dsCliente, "&_
			 "       Case when c.nucuit is null then '' else c.nucuit end as nucuit, "&_
			 "       Case when c.cdsucursal is null then 0 else c.cdsucursal end as cdsucursalCliente, "&_
			 "       p.cdproductocamara, "&_
			 "       Rtrim(p.dsproducto) as dsproducto, "&_
			 "       case when p.IcTipoEnvio is null then 0 else p.IcTipoEnvio end as IcTipoEnvio, "&_
			 "       Case when cor.CDSUCURSAL is null then 0 else cor.CDSUCURSAL end as cdsucursalCorredor, "&_
			 "       case when cor.NUCUIT is null then '' else cor.NUCUIT end as cuitcorredor, "&_
			 "       Case when pro.cdprocedenciaCamara is null then '' else pro.cdprocedenciaCamara end as cdprocedenciaCamara, "&_
             "       CASE WHEN pro.dsprocedencia IS NULL THEN '' ELSE Rtrim(pro.dsprocedencia) END AS dsprocedencia, "&_
             "       case when cor.dsCorredor is null then '' else Rtrim(cor.dsCorredor) end as dsCorredor,  "&_
             "       case when ent.dsEntregador is null then '' else Rtrim(ent.dsEntregador) end as dsEntregador, "&_
             "       case when ent.NUCUIT is null then '' else ent.NUCUIT end as cuitEntregador, "&_
             "       case when vi.dsvendedor is null then '' else Rtrim(vi.dsvendedor) end as dsIntermediario, "&_
             "       case when em.dsempresa is null then '' else Rtrim(em.dsempresa) end as dsempresa, "&_
             "       case when vi.NUDOCUMENTO is null then '' else vi.NUDOCUMENTO end as cuitIntermediario, "&_
             "       case when DO.NCUITREMITENTE is null then '' else DO.NCUITREMITENTE end as NCUITREMITENTE, "&_
             "       case when DO.CRAZONREMITENTE is null then '' else DO.CRAZONREMITENTE end as CRAZONREMITENTE, "&_
             "       DO.NCTIPOTRANSPORTE, "&_
             "       Case when DO.NCLOCDEST is null then '' else DO.NCLOCDEST end as NCLOCDEST, "&_
             "       Case when DO.NCLOCPROCE is null then '' else DO.NCLOCPROCE end as NCLOCPROCE, "&_
             "       Case when DO.NCESTABLEPROCE is null then '' when DO.NCESTABLEPROCE = '1' then '' when DO.NCESTABLEPROCE = '000001' then '' else Ltrim(DO.NCESTABLEPROCE) end as ESTABLEPROCE "&_
			 " from ( "&_
			 "   select (YEAR(ca.DtContable)*10000 + Month(ca.DtContable)*100 + DAY(ca.DtContable)) DtContable, "&_
			 "          (YEAR(ca.DtContable)*10000 + Month(ca.DtContable)*100 + DAY(ca.DtContable)) DtContableDescarga, "&_
             "          "& TIPO_TRANSPORTE_CAMION &" as tipoTransporte, "&_ 
             "          '0' cdoperativo, " &_
			 "          ca.idCamion as idTransporte, "&_ 
			 "          ca.cdProducto, "&_			 
			 "          cd.cdVendedor, "&_ 			 
			 "          cd.cdCorredor, "&_ 			 
			 "          cd.cdEntregador, "&_
			 "          cd.nuCartaPorte nuCartaPorte, "&_
			 "          cd.cdProcedencia, "&_
			 "          cd.cdCliente, "&_			 
			 "          cd.cdCosecha, "&_
			 "          ca.cdChapaCamion,  "&_
			 "          case when cd.ctg is null then '' else cd.ctg end as ctg, "&_
			 "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from HPesadasCamion pc  "&_
			 "          where pc.dtContable = ca.dtContable and pc.Idcamion = ca.Idcamion and pc.cdPesada = 1 "&_
			 "                and pc.sqpesada =  (select max(sqPesada) from HPesadasCamion "&_ 
			 "                                    where dtcontable = pc.DtContable and pc.Idcamion = Idcamion and cdPesada = 1)) as Bruto, "&_
			 "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from HPesadasCamion pc "&_ 
			 "          where pc.dtContable = ca.dtContable  and pc.Idcamion = ca.Idcamion and pc.cdPesada = 2 "&_
			 "               and pc.sqpesada =  (select max(sqPesada) from HPesadasCamion where dtcontable = pc.DtContable and Idcamion = pc.Idcamion and cdPesada = 2)) as Tara,   "&_
			 "         (select case when mc.vlMermaKilos is null then 0 else mc.vlMermaKilos end as vlMermaKilos from HMermasCamiones mc "&_ 
			 "          where mc.dtContable = ca.dtContable  and mc.Idcamion = ca.Idcamion  "&_
			 "                   and mc.sqpesada =  (select max(sqPesada) from HPesadasCamion where dtcontable = mc.DtContable and Idcamion = mc.Idcamion and cdPesada = 2)) as Merma,   "&_
			 "         concat(SC.STICKER, SC.DIGITOVERIFICADOR) as Barra, "&_
             "         (Select DISTINCT case when BIO.MUESTRA is null then '' else BIO.MUESTRA end as MUESTRA " &_
             "          from TBLMUESTRASBIOTECNOLOGIA BIO " &_
             "          where BIO.MUESTRA = (Select Max(MUESTRA) as MUESTRA from TBLMUESTRASBIOTECNOLOGIA TMB where TMB.NUCARTAPORTE=BIO.NUCARTAPORTE) " &_
			 "                  and BIO.NUCARTAPORTE=cd.NUCARTAPORTE) as BarraBio," &_
			 "         case when BIO.IDBIOTECNOLOGIA is null then 0 else BIO.IDBIOTECNOLOGIA end as IDBIOTECNOLOGIA, " &_			 
			 "          cd.cdEmpresa, "&_
			 "         (select case when c.cdAceptacion is null then 0 else c.cdAceptacion end as cdAceptacion from HCaladadeCamiones c  "&_
			 "                   where c.dtContable = ca.dtContable and c.Idcamion = ca.Idcamion  "&_
			 "                   and c.sqCalada = (select max(sqcalada) from HCaladadeCamiones where dtcontable = c.DtContable and Idcamion = c.Idcamion )) as Aceptacion,   "&_
			 "         (select c.icCamara from HCaladadeCamiones c "&_ 
			 "                   where c.dtContable = ca.dtContable and c.Idcamion = ca.Idcamion  "&_
			 "                   and c.sqCalada = (select max(sqcalada) from HCaladadeCamiones where dtcontable = c.DtContable and Idcamion = c.Idcamion )) as Camara,   "&_
			 "          ca.nuAutSalida as RECIBIDO, "&_ 
             "          case when cyo.cdvendedor is null then 0 else cyo.cdvendedor end as intermediario, "&_
             "          0 QTVAGONES " &_
			 "   from HCamiones ca "&_
			 "       inner join HCamionesDescarga Cd on  cd.dtContable = ca.DtContable and cd.Idcamion = ca.idcamion " &_
			 "		 inner join STICKERSCAMARA SC on SC.NUCARTAPORTE=CD.NUCARTAPORTE and SC.IDTRANSPORTE=CA.IDCAMION and SC.TIPOTRANSPORTE=1" &_
			 "       left join TBLBIOTECNOLOGIASDECLARADAS BIO on BIO.NUCARTAPORTE=cd.NUCARTAPORTE "&_ 
             "       left join HCUENTAYORDENESCAMIONES cyo on cyo.dtcontable = ca.DtContable and cyo.idcamion = cd.idcamion and sqorden = "& CUENTA_ORDEN_INTERMEDIARIO &_
             "       "&  myWhere
             'Si es piedrabuena agrego la consulta a los vagones
             'IMPORTANTE: Para el c�digo de barras de los Vagones siempre se utiliza el mismo c�digo que para los an�lisis de camara, el sticker se coloca a mano en la pantalla
             '            y se toma directamente de la tabla caladade vagones. Esto es as� ya que en vagones no hay una impresora y el sistema no emite los sobres directamente.
             If (pPto = TERMINAL_PIEDRABUENA) then
             strSQL = strSQL & " UNION "&_
             "     SELECT 0 DtContable, "&_
             "          (YEAR(ca.DtContableVagon)*10000 + Month(ca.DtContableVagon)*100 + DAY(ca.DtContableVagon)) DtContableDescarga, "&_
             "          "& TIPO_TRANSPORTE_VAGON &" as tipoTransporte, "&_
             "          op.cdoperativo cdoperativo, " &_
             "          ca.cdvagon AS idTransporte, "&_
             "          ca.cdproducto, "&_
             "          op.cdvendedor, "&_
             "          op.cdcorredor, "&_
             "          op.cdentregador, "&_
             "          CONCAT(op.NUCARTAPORTESERIE, SUBSTRING(op.NUCARTAPORTE, 1, 8)) nuCartaPorte, "&_             
             "          op.cdprocedencia, "&_
             "          op.cdcliente, "&_
             "          ca.cdcosecha, "&_
             "          '' as cdchapacamion, "&_
             "          '' as ctg, "&_
             "          (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   PESADASVAGON pc "&_
             "           WHERE  pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 1 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 1)) AS Bruto, "&_
             "         (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   PESADASVAGON pc "&_
             "           WHERE pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 2 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Tara, "&_
             "          (SELECT CASE WHEN mc.vlmermakilos IS NULL THEN 0 ELSE mc.vlmermakilos END AS vlMermaKilos "&_
             "           FROM   MERMASVAGONES mc "&_
             "           WHERE  mc.CDOPERATIVO = ca.CDOPERATIVO AND mc.CDVAGON = ca.CDVAGON "&_
             "                  AND mc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   PESADASVAGON "&_
             "                                     WHERE  mc.CDOPERATIVO = CDOPERATIVO AND mc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Merma, "&_
             "          concat(SC.STICKER, SC.DIGITOVERIFICADOR) as Barra, " &_
             "          (SELECT CASE WHEN c.nubarras IS NULL THEN '' ELSE Ltrim(c.nubarras) END AS nuBarras "&_
             "           FROM   CALADADEVAGONES c "&_
             "           WHERE  c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   CALADADEVAGONES "&_
             "                                    WHERE  CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS BarraBio, "&_                          
             "          CASE WHEN BIO.idbiotecnologia IS NULL THEN 0 ELSE BIO.idbiotecnologia END AS IDBIOTECNOLOGIA, "&_
             "          op.cdempresa, "&_
             "          (SELECT CASE WHEN c.cdaceptacion IS NULL THEN 0 ELSE c.cdaceptacion END AS cdAceptacion "&_
             "           FROM   CALADADEVAGONES c "&_
             "           WHERE  c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   CALADADEVAGONES "&_
             "                                    WHERE  CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Aceptacion, "&_
             "          (SELECT c.iccamara "&_
             "           FROM   CALADADEVAGONES c "&_
             "           WHERE  c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   CALADADEVAGONES "&_
             "                                    WHERE  CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Camara, "&_
             "          ca.NURECIBO as RECIBIDO, "&_
             "          0 AS intermediario, "&_
             "          QTVAGONES "&_
             "   FROM   VAGONES ca "&_
        	 "     INNER JOIN OPERATIVOS op "&_
             "          ON op.nucartaporte = ca.nucartaporte AND op.CDOPERATIVO = ca.CDOPERATIVO and op.CDESTADO not in (" & OPERATIVOS_ESTADO_TERMINADO & ") "&_
			 "		 inner join STICKERSCAMARA SC on SC.NUCARTAPORTE=CA.NUCARTAPORTE and CA.CDVAGON=SC.IDTRANSPORTE and SC.TIPOTRANSPORTE=2" &_
             "     LEFT JOIN tblbiotecnologiasdeclaradas BIO "&_
             "          ON BIO.nucartaporte = op.nucartaporte "&_
             "       " &  myWhereVagon
             strSQL = strSQL & " UNION "&_
             "     SELECT ( Year(ca.dtcontable) * 10000 + Month(ca.dtcontable) * 100 + Day(ca.dtcontable) ) DtContable, "&_
             "          (YEAR(ca.DtContableVagon)*10000 + Month(ca.DtContableVagon)*100 + DAY(ca.DtContableVagon)) DtContableDescarga, "&_
             "          "& TIPO_TRANSPORTE_VAGON &" as tipoTransporte, "&_
             "          op.cdoperativo cdoperativo, " &_
             "          ca.cdvagon AS idTransporte, "&_
             "          ca.cdproducto, "&_
             "          op.cdvendedor, "&_
             "          op.cdcorredor, "&_
             "          op.cdentregador, "&_
             "          CONCAT(op.NUCARTAPORTESERIE, SUBSTRING(op.NUCARTAPORTE, 1, 8)) nuCartaPorte, "&_             
             "          op.cdprocedencia, "&_
             "          op.cdcliente, "&_
             "          ca.cdcosecha, "&_
             "          '' as cdchapacamion, "&_
             "          '' as ctg, "&_
             "          (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   HPESADASVAGON pc "&_
             "           WHERE  pc.dtcontable = ca.dtcontable AND pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 1 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = pc.dtcontable AND pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 1)) AS Bruto, "&_
             "         (SELECT CASE WHEN pc.vlpesada IS NULL THEN 0 ELSE pc.vlpesada END AS vlPesada "&_
             "           FROM   HPESADASVAGON pc "&_
             "           WHERE  pc.dtcontable = ca.dtcontable AND pc.CDOPERATIVO = ca.CDOPERATIVO AND pc.CDVAGON = ca.CDVAGON AND pc.cdpesada = 2 "&_
             "                  AND pc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = pc.dtcontable AND pc.CDOPERATIVO = CDOPERATIVO AND pc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Tara, "&_
             "          (SELECT CASE WHEN mc.vlmermakilos IS NULL THEN 0 ELSE mc.vlmermakilos END AS vlMermaKilos "&_
             "           FROM   HMERMASVAGONES mc "&_
             "           WHERE  mc.dtcontable = ca.dtcontable AND mc.CDOPERATIVO = ca.CDOPERATIVO AND mc.CDVAGON = ca.CDVAGON "&_
             "                  AND mc.sqpesada = (SELECT Max(sqpesada) "&_
             "                                     FROM   HPESADASVAGON "&_
             "                                     WHERE  dtcontable = mc.dtcontable AND mc.CDOPERATIVO = CDOPERATIVO AND mc.CDVAGON = CDVAGON AND cdpesada = 2)) AS Merma, "&_
             "			 concat(SC.STICKER, SC.DIGITOVERIFICADOR) as Barra, " &_
             "          (SELECT CASE WHEN c.nubarras IS NULL THEN '' ELSE Ltrim(c.nubarras) END AS nuBarras "&_
             "           FROM   HCALADADEVAGONES c "&_
             "           WHERE  c.dtcontable = ca.dtcontable AND c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   HCALADADEVAGONES "&_
             "                                    WHERE  dtcontable = c.dtcontable AND CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS BarraBio , "&_
             "          CASE WHEN BIO.idbiotecnologia IS NULL THEN 0 ELSE BIO.idbiotecnologia END AS IDBIOTECNOLOGIA, "&_
             "          op.cdempresa, "&_
             "          (SELECT CASE WHEN c.cdaceptacion IS NULL THEN 0 ELSE c.cdaceptacion END AS cdAceptacion "&_
             "           FROM   HCALADADEVAGONES c "&_
             "           WHERE  c.dtcontable = ca.dtcontable AND c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   HCALADADEVAGONES "&_
             "                                    WHERE  dtcontable = c.dtcontable AND CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Aceptacion, "&_
             "          (SELECT c.iccamara "&_
             "           FROM   HCALADADEVAGONES c "&_
             "           WHERE  c.dtcontable = ca.dtcontable AND c.CDOPERATIVO = ca.CDOPERATIVO AND c.CDVAGON = ca.CDVAGON "&_
             "                  AND c.sqcalada = (SELECT Max(sqcalada) "&_
             "                                    FROM   HCALADADEVAGONES "&_
             "                                    WHERE  dtcontable = c.dtcontable AND CDOPERATIVO = c.CDOPERATIVO AND CDVAGON = c.CDVAGON)) AS Camara, "&_
             "          ca.NURECIBO as RECIBIDO, "&_
             "          0 AS intermediario, "&_
             "          QTVAGONES   " &_
             "   FROM   HVAGONES ca "&_
        	 "     INNER JOIN HOPERATIVOS op "&_
             "          ON op.dtcontable = ca.dtcontable AND op.nucartaporte = ca.nucartaporte AND op.CDOPERATIVO = ca.CDOPERATIVO "&_
			 "	   inner join STICKERSCAMARA SC on SC.NUCARTAPORTE=CA.NUCARTAPORTE  and CA.CDVAGON=SC.IDTRANSPORTE and SC.TIPOTRANSPORTE=2" &_
             "     LEFT JOIN tblbiotecnologiasdeclaradas BIO "&_
             "          ON BIO.nucartaporte = op.nucartaporte "&_
             "       " &  myWhereVagon
             end if
             strSQL = strSQL & " )T  "&_
             "      left join DATOSONCCA DO on NCCARTAPORTE = T.nuCartaPorte" &_				 
			 "		left join PRODUCTOS p on p.CDPRODUCTO = T.cdProducto "&_
			 "		left join clientes c on c.cdcliente = T.cdCliente "&_
			 "		left join vendedores v on v.cdvendedor = T.cdvendedor " &_
			 "		left join corredores cor on cor.cdcorredor = T.cdcorredor " &_
			 "		left join procedencias pro on pro.cdprocedencia = T.cdprocedencia " &_
             "      left join ENTREGADORES ent on ent.cdentregador = T.cdentregador "&_
		     "      left join vendedores vi on vi.cdvendedor = T.intermediario  "&_
             "      left join empresas em ON em.cdempresa = T.cdempresa "&_
             " where (T.BARRA <> '' or T.BARRABIO <> '') and T.tara <> 0 " & myWhereMuestra &_
			 " order by T.tipoTransporte,T.BARRA "	 
			 'response.write strSQL
			 'response.end
	Call executeQueryDB(pPto, rs, "OPEN", strSQL)		
	Set armarSQLCabecera = rs
End Function
'------------------------------------------------------------------------------------------------------------
Function getDsPuertoByPuertoCamara(pCamara, pPtoCamara)
	Dim rtrn,ptoActi
	ptoActi = getPuertoCamara2ACTI(pCamara, pPtoCamara)
    rtrn = getDsPuertoByNro(ptoActi)
	getDsPuertoByPuertoCamara = rtrn
End Function 
'------------------------------------------------------------------------------------------------------------
Function tieneBiotecnologia(cdProducto, pPto)    
    DIm rs, ret
    
    if (session("BIOTEC_" & cdProducto) = "") then
        ret = false
        strSQL="Select * from TBLBIOTECNOLOGIAS where IDPRODUCTO = " & cdProducto
        Call executeQueryDb(pPto, rs, "OPEN", strSQL)    
        if (not rs.eof) then ret = true
        session("BIOTEC_" & cdProducto) = ret
    else
        ret = session("BIOTEC_" & cdProducto)
    end if 
    tieneBiotecnologia = ret
    
End Function    
'------------------------------------------------------------------------------------------------------------
'Funcion que determina si se deben enviar o no los datos de un camion a camara.
Function enviarAnalisisACamra(tipoEnvio, icCamara, cdProducto, pPto, pMuestraBiotecnologia, pMuestraComercial, pBarra, pBarraBio)
    Dim flagBiotecnologia,flagComercial, ret
    ret = false
    
    flagBiotecnologia = false
    if ((tieneBiotecnologia(cdProducto, pPto)) and (Trim(pBarraBio) <> "")) then flagBiotecnologia = true    
    
    flagComercial = ((Cdbl(tipoEnvio) = 1)or(Cdbl(tipoEnvio) = 2 and CStr(icCamara) = "S"))and(pBarra <> "")

	'Si se pidieron las muestras comerciales y el camion tiene muestra comercial, se envia el camion.
	if ((pMuestraComercial = TIPO_AFIRMACION) and (flagComercial)) then ret = true
	'Si se pidieron las muestras comerciales y el camion tiene muestra comercial, se envia el camion.
	if ((pMuestraBiotecnologia = TIPO_AFIRMACION) and (flagBiotecnologia)) then ret = true
    enviarAnalisisACamra = ret
    
End Function
'------------------------------------------------------------------------------------------------------------------
'Rutina que elimina los codigos de analisis del puerto que trajo la camara.
Function fixAnalisis(logMig)
    strSQL= " Select CAMARA.COBECA, CAMARA.CDESCA, CAMARA.CPROCA, CAMARA.NROACA, CAMARA.FANACA, CAMARA.CPORCA, CAMARA.NSANCA, CAMARA.OBSECA, CAMARA.GRADCA, CAMARA.GRASCA, CAMARA.IMPACA, CAMARA.KGMOCA, CAMARA.FACTCA, CAMARA.USERCA, CAMARA.FECHCA, CAMARA.HORACA, " &_
            " PUERTO.COBECA COBEPU, PUERTO.CDESCA CDESPU, PUERTO.CPROCA CPROPU, PUERTO.NROACA NROAPU, PUERTO.FANACA FANAPU, PUERTO.CPORCA CPORPU, PUERTO.NSANCA NSANPU, PUERTO.OBSECA OBSEPU, PUERTO.GRADCA GRADPU, PUERTO.GRASCA GRASPU, PUERTO.IMPACA IMPAPU, PUERTO.KGMOCA KGMOPU, PUERTO.FACTCA FACTPU, PUERTO.USERCA USERPU, PUERTO.FECHCA FECHPU, PUERTO.HORACA HORAPU " &_
            " from " &_
            " (Select * from MERFL.MER591CA where COBECA < 90) CAMARA " &_
            " inner join " &_
            " (Select * from MERFL.MER591CA where COBECA >= 90) PUERTO " &_
            " on CAMARA.CPORCA = PUERTO.CPORCA and CAMARA.NSANCA=PUERTO.NSANCA "    
    logMig.Info(strSQL)                    
    Call executeQuery(rs, "OPEN", strSQL)
     
    while (not rs.eof)
        strSQL= "Select PUERTO.* from " &_
                " (Select * from MERFL.MER591DA where COBEDA=" & rs("COBECA") & " and NROADA=" & rs("NROACA") & " and CDESDA=" & rs("CDESCA") & " and CPRODA=" & rs("CPROCA") & " and FANADA=" & rs("FANACA") & ") CAMARA" &_
                " inner join " &_
                " (Select * from MERFL.MER591DA where COBEDA=" & rs("COBEPU") & " and NROADA=" & rs("NROAPU") & " and FANADA=" & rs("FANAPU") & " and CDESDA=" & rs("CDESPU") & " and CPRODA=" & rs("CPROPU") & ") PUERTO " &_
                " on CAMARA.CDESDA=PUERTO.CDESDA and CAMARA.CPRODA=PUERTO.CPRODA and CAMARA.COANDA=PUERTO.COANDA"
        logMig.Info(strSQL)                
        Call executeQuery(rs2, "OPEN", strSQL)
        while (not rs2.eof)
            strSQL= "Delete from MERFL.MER591DA where COBEDA=" & rs2("COBEDA") & " and NROADA=" & rs2("NROADA") & " and FANADA=" & rs2("FANADA") & " and CDESDA=" & rs2("CDESDA") & " and CPRODA=" & rs2("CPRODA") & " and COANDA=" & rs2("COANDA")
            logMig.Info(strSQL)
            Call executeQuery(rsX, "EXEC", strSQL)
            rs2.MoveNext()
        wend 
        rs.MoveNext()
    Wend        
End Function
'------------------------------------------------------------------------------------------------------------------
'Funcion responsable por incorporar los analisis del puerto a los de camara para las descargas incorporadas desde archivo.
Function reprocesar(logMig)
        
    logMig.info("####################################################")
    logMig.info("************* REPROCESO DE ANALISIS ***************")
    logMig.info("		-MOMENTO       :  " & GF_FN2DTE(Left(session("MmtoSistema"),8)))
    logMig.info("		-USUARIO       :  " & session("Usuario"))	    
    logMig.info("####################################################")	        
       
    Call fixAnalisis(logMig)
    
    'Obtengo las descargas que requieren reprocesar sus an�lisis desde el archivo de descargas.                
    strSQL= " Select CAMARA.COBECA, CAMARA.CDESCA, CAMARA.CPROCA, CAMARA.NROACA, CAMARA.FANACA, CAMARA.CPORCA, CAMARA.NSANCA, CAMARA.OBSECA, CAMARA.GRADCA, CAMARA.GRASCA, CAMARA.IMPACA, CAMARA.KGMOCA, CAMARA.FACTCA, CAMARA.USERCA, CAMARA.FECHCA, CAMARA.HORACA, " &_
	        " PUERTO.COBECA COBEPU, PUERTO.CDESCA CDESPU, PUERTO.CPROCA CPROPU, PUERTO.NROACA NROAPU, PUERTO.FANACA FANAPU, PUERTO.CPORCA CPORPU, PUERTO.NSANCA NSANPU, PUERTO.OBSECA OBSEPU, PUERTO.GRADCA GRADPU, PUERTO.GRASCA GRASPU, PUERTO.IMPACA IMPAPU, PUERTO.KGMOCA KGMOPU, PUERTO.FACTCA FACTPU, PUERTO.USERCA USERPU, PUERTO.FECHCA FECHPU, PUERTO.HORACA HORAPU " &_
            " from " &_
            " (Select * from MERFL.MER591CA where COBECA < 90) CAMARA " &_
            " inner join " &_
            " (Select * from MERFL.MER591CA where COBECA >= 90) PUERTO " &_
            " on CAMARA.CPORCA = PUERTO.CPORCA and CAMARA.NSANCA=PUERTO.NSANCA "            
    logMig.Info(strSQL)
    Call executeQuery(rs, "OPEN", strSQL)
    if (not rs.eof) then
        On error Resume Next
        while (not rs.eof)
	        'Comparo y defino el grado de la descarga.
	        GRADCA = CInt(rs("GRADCA"))
            GRASCA = CInt(rs("GRASCA"))
            GRADPU = CInt(rs("GRADPU"))
            GRASPU = CInt(rs("GRASPU"))
            'Siempre me quedo con el grado determinado por la c�mara por eso se comentan estas lineas.               
            'if (GRASCA < GRASPU) then GRASCA = GRASPU
            'if (GRASCA < GRADCA) then GRASCA = GRADCA
            'strSQL="Update MERFL.MER591CA set GRADCA=" & GRADCA & ", GRASCA=" & GRASCA & " where COBECA=" & rs("COBECA") & " and CPORCA=" & rs("CPORCA") & " and NSANCA=" & rs("NSANCA")
            'logMig.Info(strSQL)
            'Call MYexecuteQuery(rsI, "EXEC", strSQL)
            
            'Actualizo el detalle
            If Err.Number = 0 Then
                strSQL="Update MERFL.MER591DA set COBEDA=" & rs("COBECA") & ", NROADA=" & rs("NROACA") & ", FANADA=" & rs("FANACA") & " where CPRODA=" & rs("CPROPU") & " and CDESDA=" & rs("CDESPU") & " and COBEDA=" & rs("COBEPU") & " and NROADA=" & rs("NROAPU") & " and FANADA=" & rs("FANAPU")
                logMig.Info(strSQL)
                Call MYexecuteQuery(rsI, "EXEC", strSQL)
            end if
                                
            'Elimino la cabecera de la camara de puerto
            If Err.Number = 0 Then
                strSQL="Delete from MERFL.MER591CA where COBECA=" & rs("COBEPU") & " and CPROCA=" & rs("CPROPU") & " and CDESCA=" & rs("CDESPU") & " and CPORCA=" & rs("CPORPU") & " and NSANCA=" & rs("NSANPU")
                logMig.Info(strSQL)
                Call MYexecuteQuery(rsI, "EXEC", strSQL)
            end if                            
            
            If Err.Number <> 0 Then                                                                                                                 
                'Si hubo error.
                 logMig.Info("ERROR! Carta de Porte: " & rs("CPORCA") & " - " & Err.Description)
                 Call Err.Clear()
            end if
            rs.MoveNext()
        wend 
   else
        logMig.Info("No hay registros para reprocesar.")
   end if                      
   logMig.Info("---------- PROCESO FINALIZADO ----------")
   'Call GP_ENVIAR_MAIL_ATTACHMENT("Sincro A. Camara-Puerto", "Resultados del dia." , obtenerMail("7431"), "scalisij@toepfer.com", server.MapPath("logs/") & "/" & logMig.fileName & ".txt")
End Function
'------------------------------------------------------------------------------------------------------------------
function MYexecuteQuery(byref pRS, pOperacion, pSql)
on error resume next
    Dim con
	MYexecuteQuery = false
    if (IsEmpty(session("conn" & CONEXION_AS400 &  "Alias"))) then Call loadConfigFile(CONEXION_AS400)			
	Set con = server.CreateObject("ADODB.connection")
	con.open session("conn" & CONEXION_AS400 &  "Alias"),  session("conn" & CONEXION_AS400 &  "User"), session("conn" & CONEXION_AS400 &  "Key")			
	con.execute pSql
	con.close
	MYexecuteQuery = true		
end function

%>