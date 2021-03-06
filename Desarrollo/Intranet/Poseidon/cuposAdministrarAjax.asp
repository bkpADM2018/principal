<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/includeGeneracionArchivos.asp"-->
<%
'-------------------------------------------------------------------------------------------------------------------------            
'Controla que la cantidad de cupos ingresadas cubran correctamente el rango establecido en el contrato
Function controlarCantidadCupos(p_ArrCupo, p_ArrFecha, p_ArrMaximo, p_ArrAsignado)
    Dim i,auxCantidad, flagControl
    flagControl = true
    i = 0
    logMig.info(" Iniciando control")        
    while ((i <= UBound(p_ArrCupo))and(flagControl))
        if (CInt(p_ArrCupo(i)) > 0) then           
           auxCantidad = Cdbl(p_ArrAsignado(i)) + Cdbl(p_ArrCupo(i))
           logMig.info(" ------> Fecha: "& GF_FN2DTE(p_ArrFecha(i)) )
           logMig.info(" --------------> Cupos Maximos     : " & p_ArrMaximo(i))
           logMig.info(" --------------> Cupos ya Asignados: " & p_ArrAsignado(i))       
           logMig.info(" --------------> Cupos nuevos      : " & p_ArrCupo(i) )
           if (Cdbl(auxCantidad) > Cdbl(p_ArrMaximo(i))) then            
                flagControl = false
                logMig.info(" --------------> RESULTADO CONTROL : ERROR!")
           else
                logMig.info(" --------------> RESULTADO CONTROL : OK")
           end if
        end if
        i = i + 1
    wend    
    logMig.info(" Finalizando control")
    controlarCantidadCupos = flagControl
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function deleteNominacion(p_Pto, p_cuitCupeador, p_cuitCliente, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_CdCupo, p_esCorredor)
    Dim strSQL ,rs, max, cant
    logMig.info(" Liberando Cupos ")
    logMig.info(" --------> Fecha desde: "& GF_FN2DTE(p_fechaDesde))
    logMig.info(" --------> Fecha hasta: "& GF_FN2DTE(p_fechaHasta))
    logMig.info(" --------> Producto: "& p_CdProducto)
    strSQL = "UPDATE CODIGOSCUPO SET CDVENDEDOR=0, MOVIL='', PATENTE='' "
	if ((not p_esCorredor) or (CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER))) then strSQL = strSQL & ", CDCORREDOR=0"
    if (CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER)) then 
        strSQL = strSQL & ", ESTADO=" & CUPO_CANCELADO
    else
        if (p_esCorredor) then  strSQL = strSQL & ", ESTADO=" & CUPO_PROVISORIO	
    end if
    strSQL = strSQL & " WHERE FECHACUPO >= "& p_fechaDesde & " AND FECHACUPO <= "& p_fechaHasta &_             
             " AND CDPRODUCTO=" & p_CdProducto    &_
             " AND ESTADO <= " & CUPO_PUBLICADO_AFIP  &_
             " AND CODIGOCUPO NOT IN (Select NUCUPO from CAMIONES)"
    if (p_cuitCliente <> "") then 
        strSQL = strSQL & " AND CUITCLIENTE = '"& p_cuitCliente & "'"
        logMig.info(" --------> Cliente: "& p_cuitCliente)
    end if                          
    if (p_CdCorredor <> "") then         
        strSQL = strSQL & " AND CDCORREDOR = "& defineCdCorredor(p_cuitCliente, p_CdCorredor)
        logMig.info(" --------> Corredor: "& p_CdCorredor)
    end if
    if (p_CdVendedor <> "") then 
        logMig.info(" --------> Vendedor: "& p_CdVendedor)
        strSQL = strSQL & " AND CDVENDEDOR = "&p_CdVendedor    
    end if
    if (p_CdCupo <> "") then
        strSQL = strSQL & " AND CODIGOCUPO = '"& p_CdCupo & "'"
        logMig.info(" --------> Codigo de cupo: "& p_CdCupo)
    end if    
    Call executeQueryDb(p_Pto, rs, "EXEC", strSQL)
	'Se eliminan los cupos especiales. Los cupos especiales son los primeros que se eliminan ante una quita de cupos.
	max= "*"
	'Si se esta eliminando un solo codigo, debo restar solo 1 cupo especial.
	if (p_CdCupo <> "") then max = "TOP 1 *"
	strSQL="Select " & max & " from CODIGOSCUPOESPECIALES "  &_
			" where CUITCLIENTE='" & p_cuitCliente & "' " &_
            " and FECHACUPO>=" & p_fechaDesde & " and FECHACUPO<=" & p_fechaHasta &_
            " and CDPRODUCTO=" & p_CdProducto &_
			" and CDCORREDOR=" & defineCdCorredor(p_cuitCliente, p_CdCorredor) &_
			" and CDVENDEDOR=" & p_CdVendedor
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    while (not rs.eof)
		'con cant=0 se borran todos los cupos especiales.
		cant = 0
		if (p_CdCupo <> "") then 
			'Si se esta eliminando un solo codigo, debo restar solo 1 cupo especial.
			cant = CLng(rs("Asignados")) - 1
			'Si ya ingresaron todos los cupos asignados, no puedo quitar ninguno.
			if (cant < CLng(rs("Ingresados"))) then cant = -1
		end if
		if (cant >= 0) then Call CrearCupoEspecial(p_Pto, p_cuitCliente, p_CdCorredor, p_CdVendedor, rs("FECHACUPO"), p_CdProducto, Trim(rs("CONDICION")), cant)
		rs.MoveNext()
	wend
End function 
'-------------------------------------------------------------------------------------------------------------------------
Function enviarMailCupos(p_Pto, p_cuitCupeador, p_cuitCliente, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_YaEnviados, p_mrecep, p_esCorredor)
    Dim strSQL, rsC, ret
    if (p_cuitCliente <> "") then
        ret = enviarMailCliente(p_Pto, p_cuitCupeador, p_cuitCliente, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_YaEnviados, p_mrecep, p_esCorredor)
    else
        'Si no especifica fecha, entonces esta mandando todos los mails de un dia en particular (selecciono columna)
        strSQL="Select CUITCLIENTE, CDCORREDOR, CDVENDEDOR from CODIGOSCUPO C where FECHACUPO = " & p_fechaDesde & " AND C.CDPRODUCTO=" & p_CdProducto & " group by CUITCLIENTE, CDCORREDOR, CDVENDEDOR"        
        Call executeQueryDb(p_Pto, rsC, "OPEN", strSQL)
        while (not rsC.eof) 
            ret = ret & enviarMailCliente(p_Pto, p_cuitCupeador, rsC("CUITCLIENTE"), defineCdCorredor(rsC("CUITCLIENTE"), rsC("CDCORREDOR")), rsC("CDVENDEDOR"), p_fechaDesde, p_fechaHasta, p_CdProducto, p_YaEnviados, p_mrecep, p_esCorredor) & "<br>"
            rsC.MoveNext()
        wend
        ret = Left(ret, Len(ret)-4)
    end if 
    enviarMailCupos = ret       
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function enviarMailCliente(p_Pto, p_cuitCupeador, p_cuitCliente, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_YaEnviados, p_mrecep, p_esCorredor)
    Dim strSQL ,rs, myWhere, receptor, fileAtt, hayCorredor, hayVendedor, strAsunto, dsProducto, cantCupos
    Dim msgBody, fs, f, fileCupos
	
    logMig.info(" Enviando cupos provisorios por mail: ")
    logMig.info(" --------> Fecha desde: "& GF_FN2DTE(p_fechaDesde))
    logMig.info(" --------> Fecha hasta: "& GF_FN2DTE(p_fechaHasta))
    logMig.info(" --------> Producto: "& p_CdProducto)    
    logMig.info(" --------> Cliente: "& p_cuitCliente)
    logMig.info(" --------> Incluirlos ya Enviados: " & CBool(p_YaEnviados))
    logMig.info(" --------> Es Corredor: " & p_esCorredor)

    myWhere =   " where FECHACUPO >= "& p_fechaDesde & " AND FECHACUPO <= "& p_fechaHasta &_             
                " AND CUITCLIENTE = '"& p_cuitCliente & "'" &_
                " AND C.CDPRODUCTO=" & p_CdProducto    
    if (CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER))then 
        if (CBool(p_YaEnviados)) then
            myWhere =   myWhere & " and ESTADO >= " & CUPO_PROVISORIO
        else
            myWhere =   myWhere & " and ESTADO = " & CUPO_PROVISORIO    
        end if            
    else
        myWhere =   myWhere & " and ESTADO >=" & CUPO_OTORGADO
    end if
    
    if (p_CdCorredor <> "") then 
		if (CLng(p_CdCorredor) > 0) then
			myWhere = myWhere & " AND CDCORREDOR = "& defineCdCorredor(p_cuitCliente, p_CdCorredor)
			logMig.info(" --------> Corredor: "& p_CdCorredor)
		else
		    if (p_esCorredor) then myWhere = myWhere & " AND CDCORREDOR = "& session("KCOrganizacion")
		end if
	else
	    if (p_esCorredor) then myWhere = myWhere & " AND CDCORREDOR = "& session("KCOrganizacion") 
    end if
    
    if (p_CdVendedor <> "") then 
		if (CLng(p_CdVendedor) > 0) then
			myWhere = myWhere & " AND CDVENDEDOR = "&p_CdVendedor    
			logMig.info(" --------> Vendedor: "& p_CdVendedor)
		end if
    end if

	'--- Se determina el receptor del mail ---
	hayCorredor = False
	if ((p_CdCorredor <> "") and (CLng(p_CdCorredor) <> 0) and  (Clng(p_CdCorredor) <> SIN_CORREDOR)) then hayCorredor = True
	hayVendedor = False
	if ((p_CdVendedor <> "") and (CLng(p_CdVendedor) <> 0) and  (Clng(p_CdVendedor) <> SIN_CORREDOR)) then hayVendedor = True
	receptor = p_cuitCliente
	receptorDs = getDsClienteByCUIT(p_cuitCliente)
    if (p_mrecep = "") then	
		if (CDbl(p_cuitCupeador) = CDbl(p_cuitCliente)) then 		
			if (hayCorredor) then
				receptor = getCuitCorredorByCd(p_Pto, p_CdCorredor)
				receptorDs = getDsCorredor(p_CdCorredor)
			else
				if (hayVendedor) then
					receptor = getCuitVendedorByCd(p_Pto, p_CdVendedor)
					receptorDs = getDsVendedor(p_CdVendedor)
				end if
			end if    		
		end if      
	else
		if (hayCorredor and p_mrecep = "R") then
			receptor = getCuitCorredorByCd(p_Pto, p_CdCorredor)
			receptorDs = getDsCorredor(p_CdCorredor)
		end if    
		if (hayVendedor and p_mrecep = "V") then
			receptor = getCuitVendedorByCd(p_Pto, p_CdVendedor)
			receptorDs = getDsVendedor(p_CdVendedor)
		end if
	end if
	'-------------------------------------
	
    strSQL = "Select C.*, P.DSPRODUCTO, E.DSESTADO from CODIGOSCUPO C inner join PRODUCTOS P on P.CDPRODUCTO=C.CDPRODUCTO left join CAMIONES D on D.NUCUPO=C.CODIGOCUPO" &_
             " left join ESTADOS E on D.CDESTADO=E.CDESTADO" & myWhere & " order by CDCORREDOR, CDVENDEDOR, FECHACUPO"             
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (not rs.eof) then   	
        fileCupos = generarArchivoCupos(rs, p_Pto, p_CdProducto, receptor, receptorDs, getDsClienteByCUIT(p_cuitCliente))
        'ADM recibe un segundo archivo para importar en su sistema.
		fileAtt=""
        if (receptor = CUIT_ADM) then
            Call cargarTablaConversion(receptor, p_Pto)
            fileAtt = generarArchivoADM(rs, p_Pto, p_CdProducto, receptor)
        end if
               
        'Se envia el mail.
        auxOrigen = getTaskMailList(TASK_POS_ADMIN_CUPOS, MAIL_TASK_SENDER)
        auxDestino = getTaskMailList(TASK_POS_ADMIN_CUPOS, receptor)
		auxDestino = auxDestino & ";" & getTaskMailList(TASK_POS_ADMIN_CUPOS, p_cuitCupeador)
        if (CDbl(p_cuitCupeador) <> CDbl(CUIT_TOEPFER)) then auxDestino = auxDestino & ";" & getTaskMailList(TASK_POS_ADMIN_CUPOS, CUIT_TOEPFER)
		
        logMig.info("Enviando Mail:")
        logMig.info("ORIGEN : " & auxOrigen)
        logMig.info("DESTINO: " & auxDestino)
        '   Envia Mail		
        retMsg = ""
		set fs=Server.CreateObject("Scripting.FileSystemObject")
		Set f=fs.OpenTextFile(fileCupos,1,false)
		msgBody=f.ReadAll
		f.close		
		mail_config_Type = MAIL_TYPE_TEXT		
        if (GP_ENVIAR_MAIL_ATTACHMENT("ADM Agro S.R.L. - Cupos Asignados - " & p_Pto & " - " & getDsProducto(p_CdProducto) & " - " & receptorDs, msgBody, auxOrigen, auxDestino, fileAtt)) then                
            if (CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER)) then 
                'Se marcan los cupos informados como OTORGADOS
                strSQL = "Update CODIGOSCUPO SET ESTADO=" & CUPO_OTORGADO & replace(myWhere , "C.", "") & " and ESTADO = " & CUPO_PROVISORIO
                Call executeQueryDb(p_Pto, rs, "EXEC", strSQL)
            end if
            
            retMsg = "El mail con los cupos de " & receptorDs & "  fue enviado con exito!"
        end if            
     else
        retMsg = "No hay cupos provisorios para informar."        
     end if               
        
     logMig.info(retMsg)   
     enviarMailCliente = retMsg    
End function 
'------------------------------------------------------------------------------------------	
Function descargaArchivo(p_Pto, p_cuitCupeador, p_cuitCliente, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_esCorredor)
    Dim strSQL ,rs, myWhere, receptor, fileAtt
        
    logMig.info(" Descargando Archivo cupos: ")
    logMig.info(" --------> Fecha desde: "& GF_FN2DTE(p_fechaDesde))
    logMig.info(" --------> Fecha hasta: "& GF_FN2DTE(p_fechaHasta))
    logMig.info(" --------> Producto: "& p_CdProducto)    
    logMig.info(" --------> Cliente: "& p_cuitCliente)
    logMig.info(" --------> Cupeador: "& p_cuitCupeador)
    logMig.info(" --------> Es corredor: "& p_esCorredor)

    myWhere =   " where FECHACUPO >= "& p_fechaDesde & " AND FECHACUPO <= "& p_fechaHasta
    if (p_CdCorredor <> "") then
        myWhere = myWhere & " AND CDCORREDOR = '"& p_CdCorredor & "'"
    else        
        if (p_esCorredor) then myWhere = myWhere & " AND CDCORREDOR = "& session("KCOrganizacion") 
    end if
    if (CDbl(p_cuitCliente) > 0) then
        myWhere = myWhere & " AND CUITCLIENTE = '"& p_cuitCliente & "'"
    end if
    myWhere = myWhere & " AND C.CDPRODUCTO=" & p_CdProducto    &_
                        " and ESTADO >=" & CUPO_OTORGADO    

    if ((CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER)) and (CDbl(p_cuitCliente) = CDbl(CUIT_TOEPFER)))then 
        if ((p_CdCorredor <> "") and (CLng(p_CdCorredor) <> 0) and  (Clng(p_CdCorredor) <> SIN_CORREDOR)) then
            receptor = getCuitCorredorByCd(p_Pto, p_CdCorredor)
            receptorDs = getDsCorredor(p_CdCorredor)
        else
            receptor = getCuitVendedorByCd(p_Pto, p_CdVendedor)
            receptorDs = getDsVendedor(p_CdVendedor)
        end if
    else
        receptor = p_cuitCupeador
        if (p_CdCorredor <> "") then
            receptorDs = getDsCorredor(p_CdCorredor)            
        else
            receptorDs = getDsClienteByCUIT(p_cuitCupeador)
        end if
    end if        
    
    strSQL = "Select C.*, P.DSPRODUCTO, E.DSESTADO from CODIGOSCUPO C inner join PRODUCTOS P on P.CDPRODUCTO=C.CDPRODUCTO left join CAMIONES D on D.NUCUPO=C.CODIGOCUPO" &_
             " left join ESTADOS E on D.CDESTADO=E.CDESTADO" & myWhere & " order by CDCORREDOR DESC, CDVENDEDOR DESC, FECHACUPO"
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (not rs.eof) then   
        retMsg = generarArchivoCupos(rs, p_Pto, p_CdProducto, receptor, receptorDs, getDsClienteByCUIT(p_cuitCliente))
        logMig.info(retMsg)
     else
        logMig.info("No hay cupos")
        retMsg = ""        
     end if     
     descargaArchivo = retMsg    
End function 
'------------------------------------------------------------------------------------------	
Function dibujarCampo(line, val, sz)

    dibujarCampo = line & GF_nChars(val,sz," ", CHR_AFT)
    
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function generarArchivoCupos(pRs, pPto, pCdProducto, pCuitReceptor, pDsReceptor, pDsCliente)
        
    Dim fso, f, fname
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")      
    fname = Server.MapPath(".") & "\temp\CuposAsignados_" & pPto & "_" & pCdProducto & "_" & pCuitReceptor & ".txt"
    Set f = fso.OpenTextFile(fname, 2, true)    

    if (not pRs.Eof) then        
        f.WriteLine "Cupos otorgados a " & pDsReceptor
		f.WriteLine "Destinatario de la Mercaderia: " & pDsCliente
		f.WriteLine ""
        logMig.info("Cupos otorgados a " & pDsReceptor)
        linea = ""       
        linea = dibujarCampo(linea, "Fecha", 17)        
        linea = dibujarCampo(linea, "Cupo Asignado", 17)        
        linea = dibujarCampo(linea, "Producto", 15)
        linea = dibujarCampo(linea, "Corredor", 77)        
        linea = dibujarCampo(linea, "Vendedor", 64)
        linea = dibujarCampo(linea, "Estado", 50)
        f.WriteLine linea
        logMig.info(linea)
        linea = ""       
        'linea = dibujarCampo(linea, "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", 145)
        f.WriteLine linea
        logMig.info(linea)
        while (not pRs.Eof)            
            linea = ""
            linea = dibujarCampo(linea, GF_FN2DTE(pRs("FECHACUPO")), 13)
            linea = dibujarCampo(linea, pRs("CODIGOCUPO"), 20)
            linea = dibujarCampo(linea, Left(pRs("DSPRODUCTO"), 15), 15)
            if ((CLng(pRs("CDCORREDOR")) > 0) and (CLng(pRs("CDCORREDOR")) <> SIN_CORREDOR)) then            
                linea = dibujarCampo(linea, defineCdCorredor(pRs("CUITCLIENTE"), pRs("CDCORREDOR")) &"-"& Left(getDsCorredor(defineCdCorredor(pRs("CUITCLIENTE"), pRs("CDCORREDOR"))), 50), 50)
            else
                linea = dibujarCampo(linea, "", 50)
            end if
            if (CLng(pRs("CDVENDEDOR")) > 0) then
                linea = dibujarCampo(linea, pRs("CDVENDEDOR") &"-"& Left(getDsVendedor(pRs("CDVENDEDOR")), 50), 50)
            else
                linea = dibujarCampo(linea, "", 50)
            end if
            auxEstado = "OTORGADO"
            if (pRs("DSESTADO") <> "") then auxEstado = pRs("DSESTADO")
            linea = dibujarCampo(linea, auxEstado, 50)
            f.WriteLine linea
            logMig.info(linea)
            pRs.MoveNext()
        wend        
    end if    
    
    f.close()
    
    Set f = Nothing
    Set fso = Nothing
    
    generarArchivoCupos = fname
    
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function generarArchivoADM(pRs, pPto, pCdProducto, pCuitReceptor)
        
    Dim fso, f, fname, registro, msg
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")      
    fname = Server.MapPath(".") & "\temp\DataCupos_" & pPto & "_" & pCdProducto & "_" & pCuitReceptor & ".txt"
    Set f = fso.OpenTextFile(fname, 2, true)    
    
    pRs.MoveFirst()
    if (not pRs.Eof) then        
        logMig.info("Generadno archivo para importar en Sistema.")
        while (not pRs.Eof)
            '1.- C�digo de Cupo (Alfanum�rico. m�x 11 posiciones)
            registro = pRs("CODIGOCUPO") & "|"
            '2.- Fecha del Cupo (Formato AAAAMMDD)
            registro = registro & pRs("FECHACUPO") & "|"
            '3.- C�digo de Producto (Seg�n tabla ADM)
            registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_PRODUCTO, pCdProducto, msg), 3) & "|"
            '4.- C�digo de Puerto (Seg�n tabla ADM)
            registro = registro & GF_nDigits(convertirDatoPuerto(CONV_KEY_PUERTO, pPto, msg), 2)
            f.WriteLine registro
            logMig.info(registro)
            pRs.MoveNext()
        wend        
    end if    
    
    f.close()
    
    Set f = Nothing
    Set fso = Nothing
    
    generarArchivoADM = fname
    
End Function
'-------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------CORREDORES----------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'Administro y genero el codigo de corredor para los tres puertos
Function generarCodigoCorredor(p_cuitCorredor, p_dsCorredor, p_Pto, p_codigo)
    Dim auxCdCorredor, myCdCorredor
    
    myCdCorredor = p_codigo
    if (myCdCorredor = "") then myCdCorredor = 0
    
    'Compruebo para cada puerto si existe el cuit ingresado, si no existe lo agrego. Luego devuelvo el codigo del corredor para el puerto en cuestion (parametro p_Pto)
    'Arroyo:
    if (p_Pto = TERMINAL_ARROYO) and (myCdCorredor > 0) then
        generarCodigoCorredor = myCdCorredor
    else               
        auxCdCorredor = myCdCorredor        
        if ((myCdCorredor = 0) and (p_cuitCorredor <> "")) then
            auxCdCorredor = getCdCorredorByCuit(TERMINAL_ARROYO, p_cuitCorredor)
            if (Cdbl(auxCdCorredor) = 0) then 
                auxCdCorredor = getNextCdCorredor(TERMINAL_ARROYO)
                Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_ARROYO)
            end if                
        end if        
        if (p_Pto = TERMINAL_ARROYO) then generarCodigoCorredor = auxCdCorredor
    end if
    
    'Transito:
    if (p_Pto = TERMINAL_TRANSITO) and (myCdCorredor > 0) then
        generarCodigoCorredor = myCdCorredor
    else               
        auxCdCorredor = myCdCorredor
        if ((myCdCorredor = 0) and (p_cuitCorredor <> "")) then
            auxCdCorredor = getCdCorredorByCuit(TERMINAL_TRANSITO, p_cuitCorredor)
            if (Cdbl(auxCdCorredor) = 0) then 
                auxCdCorredor = getNextCdCorredor(TERMINAL_TRANSITO)
                Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_TRANSITO)
            end if                
        end if
        if (p_Pto = TERMINAL_TRANSITO) then generarCodigoCorredor = auxCdCorredor
    end if
    
    'Piedrabuena:
    if (p_Pto = TERMINAL_PIEDRABUENA) and (myCdCorredor > 0) then
        generarCodigoCorredor = myCdCorredor
    else               
        auxCdCorredor = myCdCorredor
        if ((myCdCorredor = 0) and (p_cuitCorredor <> "")) then
            auxCdCorredor = getCdCorredorByCuit(TERMINAL_PIEDRABUENA, p_cuitCorredor)
            if (Cdbl(auxCdCorredor) = 0) then 
                auxCdCorredor = getNextCdCorredor(TERMINAL_PIEDRABUENA)
                Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_PIEDRABUENA)
            end if
        end if
        if (p_Pto = TERMINAL_PIEDRABUENA) then generarCodigoCorredor = auxCdCorredor
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Obtenemos el proximo Codigo de corredor libre para agregar
Function getNextCdCorredor(p_Pto)
    Dim strSQL, rs 
    getNextCdCorredor = 100000
    strSQL = "Select NEXTID from "&_
             "(Select 99999 + ROW_NUMBER ( ) over (ORDER BY CDCORREDOR) NEXTID , *   from CORREDORES C where C.CDCORREDOR>=100000) T where NEXTID <>CDCORREDOR order by NEXTID"
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (rs.Eof) then         
        strSQL="Select (MAX(CDCORREDOR)+1) NEXTID from CORREDORES C where C.CDCORREDOR>=100000"
        Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    end if        
    if (not rs.Eof) then 
        if (rs("NEXTID") <> "") then getNextCdCorredor = rs("NEXTID")
    end if        
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Agrego un nuevo corredor a la planta
Function addNewCorredor(p_cdCorredor,p_cuitCorredor, p_dsCorredor, p_Pto)
    Dim strSQL
    strSQL = "INSERT INTO CORREDORES (CDCORREDOR, DSCORREDOR, CDTIPODOC, NUCUIT ) "&_
             "VALUES ("& p_cdCorredor &",'"& UCase(Trim(p_dsCorredor)) &"','C.U.I.T.','"& Trim(p_cuitCorredor) &"')"
    Call executeQueryDb(p_Pto, rs, "EXEC", strSQL)
End function
'-------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------VENDEDORES----------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'Administro y genero el codigo de vendedor para los tres puertos
Function generarCodigoVendedor(p_cuitVendedor, p_dsVendedor, p_Pto, p_codigo)
    Dim auxCdVendedor, myCdVendedor
    
    myCdVendedor = p_codigo
    if (myCdVendedor = "") then myCdVendedor = 0
    'Compruebo para cada puerto si existe el cuit ingresado, si no existe lo agrego. Luego devuelvo el codigo del Vendedor para el puerto en cuestion (parametro p_Pto)
    'Arroyo:
    if (p_Pto = TERMINAL_ARROYO) and (myCdVendedor > 0) then
        generarCodigoVendedor = myCdVendedor
    else                
        auxCdVendedor = myCdVendedor
        if ((myCdVendedor = 0) and (p_cuitVendedor <> "")) then
            auxCdVendedor = getCdVendedorByCuit(TERMINAL_ARROYO, p_cuitVendedor)
            if (Cdbl(auxCdVendedor) = 0) then 
                auxCdVendedor = getNextCdVendedor(TERMINAL_ARROYO)
                Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_ARROYO)
            end if                
        end if
        if (p_Pto = TERMINAL_ARROYO) then generarCodigoVendedor = auxCdVendedor
    end if

    'Transito:
    if (p_Pto = TERMINAL_TRANSITO) and (p_codigo > 0) then
        generarCodigoVendedor = myCdVendedor
    else                
        auxCdVendedor = myCdVendedor
        if ((myCdVendedor = 0) and (p_cuitVendedor <> "")) then
            auxCdVendedor = getCdVendedorByCuit(TERMINAL_TRANSITO, p_cuitVendedor)
            if (Cdbl(auxCdVendedor) = 0) then 
                auxCdVendedor = getNextCdVendedor(TERMINAL_TRANSITO)
                Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_TRANSITO)
            end if                
        end if
        if (p_Pto = TERMINAL_TRANSITO) then generarCodigoVendedor = auxCdVendedor
    end if
    
    'Piedrabuena:
    if (p_Pto = TERMINAL_PIEDRABUENA) and (myCdVendedor > 0) then
        generarCodigoVendedor = myCdVendedor
    else                
        auxCdVendedor = myCdVendedor
        if ((myCdVendedor = 0) and (p_cuitVendedor <> "")) then
            auxCdVendedor = getCdVendedorByCuit(TERMINAL_PIEDRABUENA, p_cuitVendedor)
            if (Cdbl(auxCdVendedor) = 0) then 
                auxCdVendedor = getNextCdVendedor(TERMINAL_PIEDRABUENA)
                Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_PIEDRABUENA)
            end if                
        end if
        if (p_Pto = TERMINAL_PIEDRABUENA) then generarCodigoVendedor = auxCdVendedor
    end if
    
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Obtenemos el proximo Codigo de Vendedor libre para agregar
Function getNextCdVendedor(p_Pto)
    Dim strSQL, rs 
    getNextCdVendedor = 100000
    strSQL = "Select NEXTID from "&_
             "(Select 99999 + ROW_NUMBER ( ) over (ORDER BY CDVENDEDOR) NEXTID , *   from VENDEDORES C where C.CDVENDEDOR>=100000) T where NEXTID <>CDVendedor order by NEXTID"
    Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    if (rs.Eof) then         
        strSQL="Select (MAX(CDVENDEDOR)+1) NEXTID from VENDEDORES C where C.CDVENDEDOR>=100000"
        Call executeQueryDb(p_Pto, rs, "OPEN", strSQL)
    end if        
    if (not rs.Eof) then 
        if (rs("NEXTID") <> "") then getNextCdVendedor = rs("NEXTID")
    end if        
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Agrego un nuevo Vendedor a la planta
Function addNewVendedor(p_cdVendedor,p_cuitVendedor, p_dsVendedor, p_Pto)
    Dim strSQL
    strSQL = "INSERT INTO VENDEDORES (CDVENDEDOR, DSVENDEDOR, CDTIPODOC, NUDOCUMENTO ) "&_
             "VALUES ("& p_cdVendedor &",'"& UCase(Trim(p_dsVendedor)) &"','C.U.I.T.','"& Trim(p_cuitVendedor) &"')"
    Call executeQueryDb(p_Pto, rs, "EXEC", strSQL)
End function
'----------------------------------------------------------------------------------------------------------------------------
Function cargarArrayCupos(p_Pto, p_cuitCupeador, p_FechaDesde, p_FechaHasta, p_cdProducto)
    Dim i,posicionesArray, maxCuposDisponibles, rsMax, rsAsig, strSQL, strSQL2, rsAsigOP
    
    if (CDbl(p_cuitCupeador) = CDbl(CUIT_TOEPFER)) then
        maxCuposDisponibles = getValueParametro(CUPOS_MAX_DISPONIBLES, p_Pto)        
        strSQL= " Select FECHACUPO, count(*) CANTIDAD from CODIGOSCUPO C " &_
                "   where C.FECHACUPO <= " & p_FechaHasta &_
                "       and C.FECHACUPO >= " & p_FechaDesde &_            
                "       and ESTADO <> " & CUPO_CANCELADO &_
                "   group by FECHACUPO " &_
                "   order by FECHACUPO"            
        Call executeQueryDb(p_Pto, rsAsig, "OPEN", strSQL)        
    else
        maxCuposDisponibles = 0
        strSQL2= " Select FECHACUPO, count(*) CANTIDAD from CODIGOSCUPO C " &_
                "   where ((C.CUITCLIENTE = " &  p_cuitCupeador & " and C.ESTADO >= " & CUPO_OTORGADO & ")" &_
				"			or (C.CUITCLIENTE = '" & CUIT_TOEPFER & "' and C.CDCORREDOR = " & session("KCOrganizacion") & " and C.ESTADO >= " & CUPO_PROVISORIO & "))" &_
                "       and C.FECHACUPO <= " & p_FechaHasta &_
                "       and C.FECHACUPO >= " & p_FechaDesde &_
                "       and C.CDPRODUCTO = " & p_cdProducto &_                 
                "       <ESPECIAL>" &_
                "   group by FECHACUPO " &_
                "   order by FECHACUPO"            
        strSQL = Replace(strSQL2, "<ESPECIAL>", "")
        Call executeQueryDb(p_Pto, rsMax, "OPEN", strSQL)
        strSQL = Replace(strSQL2, "<ESPECIAL>", "and CDVENDEDOR <> 0")
        Call executeQueryDb(p_Pto, rsAsig, "OPEN", strSQL)
    end if
    
    posicionesArray = GF_DTEDIFF(p_FechaDesde,p_FechaHasta,"D")
    redim arrMaximo(posicionesArray)
    redim arrAsignado(posicionesArray)
    redim arrCupo(posicionesArray)
    redim arrFecha(posicionesArray)
    i = 0
    while (i <= posicionesArray)
        arrCupo(i)  = GF_PARAMETROS7("cupo_" & i, 0, 6)
        arrFecha(i) = GF_DTEADD(p_FechaDesde, i, "D")                
        arrMaximo(i) = maxCuposDisponibles            
        if (CDbl(p_cuitCupeador) <> CDbl(CUIT_TOEPFER)) then
            if (not rsMax.eof) then
                if (CDbl(rsMax("FECHACUPO")) = CDbl(arrFecha(i))) then 
                    arrMaximo(i) = rsMax("CANTIDAD")
                    rsMax.MoveNext()
                end if               
            end if
        end if            
        arrAsignado(i) = 0        
        if (not rsAsig.eof) then
            if (CDbl(rsAsig("FECHACUPO")) = CDbl(arrFecha(i))) then 
                arrAsignado(i) = rsAsig("CANTIDAD")
                rsAsig.MoveNext()
            end if
        end if
        i = i + 1
    wend
End function 
'----------------------------------------------------------------------------------------------------------------------------
Function CrearCupos(pPto, pFecha, pCdProducto, pCuitCliente, pCdCorredor, pCdVendedor, pCantidad)

    Dim strSQL, rsX, h, myKey, colName, dteDiff
    
    logMig.info(" Creando Cupos")
    For h = 1 to pCantidad
        strSQL="Insert into CODIGOSCUPO(FECHACUPO, CDPRODUCTO, CUITCLIENTE, CDCORREDOR, CDVENDEDOR, CODIGOCUPO, PATENTE, MOVIL, ESTADO, MMTO, " &_
               "idcupoWS, cuitOrigenWS, cuitIntermediarioWS, cuitRemComercialWS, cuitRepresentanteEntregadorWS, cuitTransportistaWS, cuitChoferWS, idCuitOrigenWS, " &_
               "idCuitIntermediarioWS, idCuitRemComercialWS, idCuitCorredorVWS, idCuitRepresentanteEntregadorWS, idCuitTransportistaWS, idCuitChoferWS, ctgWS, " &_
		       "fechaCTG_desdeWS, fechaCTG_HastaWS, cartaporteWS, fechaCP_cargaWS, fechaCP_VtoWS, codLocalidadOrigenWS, desvioWS, cosechaWS, nroEstablecimientoOrigenWS, pesoNetoEstimadoWS, kmRecorrerWS, dominioWS, " &_
		       "cuitMercadoATerminoWS, cuitCorredorCWS, cuitIntermediarioFleteWS, idCuitMercadoATerminoWS, idCuitCorredorCWS, idCuitDestinatarioWS, idCuitIntermediarioFleteWS, idTurnoDetalleWS, TurnoDetalleWS, renspaWS, cantHorasSalidaCamionWS, nroContratoWS, idEstadoEnPlantaWS, IDSector, CDUsuario) " &_
		       "values (" & pFecha & ", " & pCdProducto & ", '" & pCuitCliente & "', " & defineCdCorredor(pCuitCliente, pCdCorredor) & ", " & pCdVendedor & ",'PROVISORIO', '', '', " & CUPO_PROVISORIO & ", " & session("MmtoDato") &_
		       ", 0, '', '', '', '', '', '', 0, 0, 0, 0, 0, 0, 0, '', '2000-01-01 00:00:00.000', '2000-01-01 00:00:00.000', '', '2000-01-01 00:00:00.000', '2000-01-01 00:00:00.000', 0, '', '', 0, 0, 0, ''" &_
		       ", '', '', '', 0, 0, 0, 0, 0, '', '', 0, '', 0, " & session("UsuarioSector") & ", '" & session("Usuario") & "')"
        Call executeQueryDb(pPto, rsX, "EXEC", strSQL)
    Next
    	
    'Se codifican los cupos tomando como indice el ID asignado en el campo identidad.    
	logMig.info(" Asignando nuevos codigos:")
	myKey = getLetraCupo2(pPto)
	strSQL="Select DSPRODUCTO from PRODUCTOS where CDPRODUCTO=" & pCdProducto        
	Call executeQueryDb(pPto, rs, "OPEN", strSQL)
	if (not rs.eof) then 
		myKey = myKey & Left(Trim(rs("DSPRODUCTO")), 1)
	else
		myKey = myKey & "X"
	end if
	dteDiff = GF_DTEDIFF(CUPOS_FECHA_BASE, pFecha, "D")	
	colName = "M" & Right("0" & CStr(((CLng(pFecha) mod 100) mod 12) + 1), 2)		
	strSQL ="UPDATE  CODIGOSCUPO SET CODIGOCUPO = concat('" & myKey & "', Right(concat('0000', '" & dteDiff & "'), 4), Right(concat('0000', B." & colName & "), 4)) " &_
			" FROM    CODIGOSCUPO a " &_
			"	INNER JOIN CODIGOSCUPOMTX b " &_
			"	ON (A.IDCUPO % 10000) = B.IDX " &_
			" where CODIGOCUPO='PROVISORIO' and CDPRODUCTO=" & pCdProducto
	Call executeQueryDb(pPto, rsX, "EXEC", strSQL)
    logMig.info(" Creacion Finalizada")
End Function
'----------------------------------------------------------------------------------------------------------------------------
Function controlarCorVen(pCuitCliente, pCdCorredor, pCdVendedor)
'/*** SE LEVANTA EL CONTROL HASTA SINCRONIZAR LOS PROVEEDORES DE LOS PUERTOS CON BS AS ***/
    controlarCorVen = true
    if (CDbl(pCuitCliente) = CDbl(CUIT_TOEPFER)) then
        'Si el cliente es TOEPFER el corredor y vendedor no puede ser mayor a 100000
        'if (CLng(pCdCorredor) >= 100000) or (CLng(pCdVendedor) >= 100000) then controlarCorVen = false    
    end if
    
End Function
'----------------------------------------------------------------------------------------------------------------------------
Function CrearCupoEspecial(pPto, pCuitCliente,pCdCorredor,pCdVendedor,pFecha, pCdProducto, pCondicion, pCantidad)
    Dim msg, strSQL, rs, myWhere, myCor, myVen
    
    myCor = 0
    myVen = 0    
    myWhere = " where CUITCLIENTE='" & pCuitCliente & "' " &_
                    " and FECHACUPO=" & pFecha &_
                    " and CDPRODUCTO=" & pCdProducto &_
					" and CDCORREDOR=" & defineCdCorredor(pCuitCliente, pCdCorredor) &_
					" and CDVENDEDOR=" & pCdVendedor
	myCor = pCdCorredor
	myVen = pCdVendedor
		
    msg = ValidarCupoEspecial(pPto, pCuitCliente,pCdCorredor,pCdVendedor,pFecha, pCdProducto, pCondicion, pCantidad, myWhere)
    if (msg = "") then
        if (CLng(pCantidad) = 0) then
            'Esta queriendo borrar los especieales
            strSQL="Delete from CODIGOSCUPOESPECIALES "  & myWhere & " and CONDICION = '" & pCondicion & "'"
        else
            strSQL="Select * from CODIGOSCUPOESPECIALES "  & myWhere & " and CONDICION = '" & pCondicion & "'"
            Call executeQueryDb(pPto, rs, "OPEN", strSQL)
            if (not rs.eof) then
                strSQL="Update CODIGOSCUPOESPECIALES Set QTASIGNADOS=" & pCantidad & myWhere & " and CONDICION = '" & Trim(pCondicion) & "'"
            else            
                strSQL="Insert into CODIGOSCUPOESPECIALES values (" & pFecha & ", " & pCdProducto & ", '" & pCuitCliente & "', " & defineCdCorredor(pCuitCliente, myCor) & ", " & pCdVendedor & ", '" & Trim(pCondicion) & "', " & pCantidad & ", 0)"
            end if
        end if                             
        Call executeQueryDb(pPto, rsX, "EXEC", strSQL)                
    end if
    CrearCupoEspecial = msg
End Function    
'----------------------------------------------------------------------------------------------------------------------------
Function ValidarCupoEspecial(pPto, pCuitCliente,pCdCorredor,pCdVendedor,pFecha, pCdProducto, pCondicion, pCantidad, pWhere)
    Dim strSQL, rs, saldo, ingresados
    
    if (pCondicion <> "") then
        if (CLng(pCantidad) >= 0) then        
            'Valido que la cantidad no supere los cupos asignados y las condiciones ya cargadas.            
            strSQL= "Select count(*) CANTIDAD from CODIGOSCUPO " & pWhere                          
            Call executeQueryDb(pPto, rs, "OPEN", strSQL)
            saldo = 0            
            if (not rs.eof) then saldo = rs("CANTIDAD")
            strSQL="Select case when SUM(QTASIGNADOS) is Null then 0 else SUM(QTASIGNADOS) end CANTIDAD from CODIGOSCUPOESPECIALES "  & pWhere & " and CONDICION <> '" & pCondicion & "'"
            Call executeQueryDb(pPto, rs, "OPEN", strSQL)
            if (not rs.eof) then saldo = CLng(saldo) - CLng(rs("CANTIDAD"))            
            if (CLng(saldo) >= CLng(pCantidad)) then
                'VErifico que la cantidad no sea menor que los ya ingresdos (Por si se quiere eliminar)
                strSQL="Select case when SUM(QTINGRESADOS) is Null then 0 else SUM(QTINGRESADOS) end CANTIDAD from CODIGOSCUPOESPECIALES "  & pWhere & " and CONDICION = '" & pCondicion & "'"                
                Call executeQueryDb(pPto, rs, "OPEN", strSQL)
                if (not rs.eof) then ingresados = CLng(rs("CANTIDAD"))
                if (CLng(ingresados) <= CLng(pCantidad)) then
                    ValidarCupoEspecial = ""
                else                
                    ValidarCupoEspecial = "La cantidad no puede ser menor a los cupos ya ingresados."
                end if                    
            else
                ValidarCupoEspecial = "La cantidad supera los cupos disponibles."
            end if            
        else
            ValidarCupoEspecial = "La cantidad ingresada es incorrecta."
        end if    
    else
        ValidarCupoEspecial = "Falta especificar una condici�n."
    end if    
End Function
'*****************************************************************************************************************************
'**************************************************** INICIO DE PAGINA *******************************************************
'*****************************************************************************************************************************
Dim cuitCliente, rs,fecha, i,cdVendedor,cdCorredor,cantidad,fechaDesde,fechaHasta,logMig,numeroPto,idCupo, arrCupo(),arrFecha()
Dim maxIndice, cdProducto, cuitCupeador, arrMaximo(), arrAsignado(), dicConv, esEspecial, condicion, myLckKey, mrecep, flagSeguir
Dim flagEsCorredor

cuitCupeador = GF_PARAMETROS7("cuitCupeador", "",6)

'Se controla el acceso - Solo se permite elegir el proveedor por parametro si el usuario de la session es TOEPFER
if (CDbl(cuitCupeador) <> CDbl(session("CuitOrganizacion"))) then
    Response.Write "Faltan permisos para completar la operacion."
    response.end
end if

Call GP_CONFIGURARMOMENTOS()

cuitCliente = GF_PARAMETROS7("cuitCliente", "",6)
accion = GF_PARAMETROS7("accion","",6)
cdProducto = GF_PARAMETROS7("cdProducto",0,6)
fechaHasta = GF_PARAMETROS7("fechaHasta","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde","",6)
cdCorredor = GF_PARAMETROS7("cdCorredor", "",6)
cuitCorredor = GF_PARAMETROS7("cuitCorredor","",6)
dsCorredor = GF_PARAMETROS7("dsCorredor","",6)
cdVendedor = GF_PARAMETROS7("cdVendedor", "",6)
cuitVendedor = GF_PARAMETROS7("cuitVendedor","",6)
dsVendedor = GF_PARAMETROS7("dsVendedor","",6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
cdCupo = GF_PARAMETROS7("cdCupo","",6)
esEspecial = GF_PARAMETROS7("especial",0,6)
condicion = UCase(GF_PARAMETROS7("cond", "",6))
cantidad = GF_PARAMETROS7("cant", 0,6)
yaEnviados = GF_PARAMETROS7("forzar", 0,6)
mrecep	= GF_PARAMETROS7("mrecep","",6)
fc	= GF_PARAMETROS7("fc","",6)

'Solo pueden cupear el cliente, el corredor o Toepfer =>
'Si no es Toepfer ni el cleinte tiene que ser el corredor.
flagEsCorredor = false
if ((fc = "C") and (CDbl(cuitCupeador) <> CDbl(CUIT_TOEPFER))) then flagEsCorredor=true

if (CDbl(cuitCupeador) <> CDbl(CUIT_TOEPFER)) then 
    if (flagEsCorredor) then 
        cuitCliente = CUIT_TOEPFER
    else
        cuitCliente = cuitCupeador 
    end if
end if

Set logMig = new classLog
Call startLog(HND_FILE, MSG_INF_LOG)
logMig.fileName = "ADMINISTRACION_CUPOS-" & cuitCupeador & "-" & Left(Session("MmtoDato"),8)
logMig.info("****************************************** INICIA *********************************************************")
logMig.info("-----> ACCION  : " & UCase(accion))    
logMig.info("-----> PRODUCTO: " & cdProducto)

'Primero verifico ellock de seguridad para evitar accesos multiples.
'myLckUsr = getLckUser(session("Usuario"))
'myLckKey = LCK_LOGISTICA & "_" & cuitCupeador & "_" & cdProducto
flagSeguir = true
'if (CDbl(cuitCupeador) <> CDbl(CUIT_TOEPFER)) then flagSeguir = checkLckKey(g_strPuerto, myLckKey, myLckUsr)
if (flagSeguir) then
    select case accion
        case ACCION_VISUALIZAR
            Call Descargar(descargaArchivo(g_strPuerto, cuitCupeador,cuitCliente,cdCorredor,cdVendedor,fechaDesde,fechaHasta, cdProducto, flagEsCorredor))
        case ACCION_BORRAR
            Call deleteNominacion(g_strPuerto, cuitCupeador,cuitCliente,cdCorredor,cdVendedor,fechaDesde,fechaHasta, cdProducto,cdCupo, flagEsCorredor)
            Response.Write RESPUESTA_OK
        case ACCION_EMAIL
            response.write enviarMailCupos(g_strPuerto, cuitCupeador,cuitCliente,cdCorredor,cdVendedor,fechaDesde,fechaHasta, cdProducto, yaEnviados, mrecep, flagEsCorredor)
        case ACCION_GRABAR
            if (esEspecial = 1) then
                'Se agrega el cupo especial
                Response.Write CrearCupoEspecial(g_strPuerto, cuitCliente,cdCorredor,cdVendedor,fechaDesde, cdProducto, condicion, cantidad)
            else
                'Primero cargamos los array con los valores recibidos por parametros , estos array son globales
                Call cargarArrayCupos(g_strPuerto, cuitCupeador, fechaDesde,fechaHasta, cdProducto)
                'Controlo que la cantidad asignada sea correcta con el historico de nominaciones para el contrato y su rango (Desde-Hasta)
                if (controlarCantidadCupos(arrCupo,arrFecha,arrMaximo,arrAsignado)) then            

                    'Obtengo los codigo de vendedor y de corredor para grabar la nominacion
                    cdVendedor = generarCodigoVendedor(cuitVendedor,dsVendedor,g_strPuerto, cdVendedor)
                    cdCorredor = generarCodigoCorredor(cuitCorredor,dsCorredor,g_strPuerto, cdCorredor)
                    
                    if (controlarCorVen(cuitCliente,cdCorredor,cdVendedor)) then
                        for i = 0 to UBound(arrFecha)
                            'Recorro los valores ingresados para las fechas, solo se calcula cuando los kilos son mayores a 0
                            if (CInt(arrCupo(i)) > 0) then
                                'Cantidad de cupos a nominar
                                cantidad = arrCupo(i)
                                if (CDbl(cuitCupeador) = CDbl(CUIT_TOEPFER)) then
                                    'Es toepfer, crea nuevos cupos                                                            
                                    Call CrearCupos(g_strPuerto, arrFecha(i), cdProducto, cuitCliente, cdCorredor, cdVendedor, cantidad)
                                else
                                    'Es un tercero, actualiza los suyos cn corredor y vendedor.
									if (CDbl(cuitCupeador) <> CDbl(cuitCliente)) then
										'Si es un corredor, actualiza los cupos provisorios.									
										sqlEstado = "ESTADO in (" & CUPO_PROVISORIO & ", " & CUPO_OTORGADO & ", " & CUPO_PUBLICADO_AFIP & ")"
										sqlEstado2 = ", ESTADO = " & CUPO_OTORGADO
										sqlCorredor = cdCorredor
									else
										'Si es un exportador, solo los OTORGADOS y Enviados a AFIP.
										sqlEstado = "ESTADO in (" & CUPO_OTORGADO & ", " & CUPO_PUBLICADO_AFIP & ")"
										sqlEstado2 = ""
										sqlCorredor = "0"
									end if
                                    strSQL="Update CODIGOSCUPO SET CDCORREDOR=" & cdCorredor & ", CDVENDEDOR=" & cdVendedor & sqlEstado2 & " where CODIGOCUPO in (Select TOP(" & cantidad & ") CODIGOCUPO from CODIGOSCUPO where CUITCLIENTE='" & cuitCliente & "' and FECHACUPO=" & arrFecha(i) & " and CDPRODUCTO=" & cdProducto & " and CDCORREDOR=" & sqlCorredor & " and CDVENDEDOR=0 and " & sqlEstado & ")"
                                    Call executeQueryDb(g_strPuerto, rsX, "EXEC", strSQL)
                                end if                    
                            end if
                        next
                        'Si grabo los datos correctamente envio el codigo de vendedor/comprador por si es nuevo alguno de los dos y fijar los valores visualmente en la pagina
                        Response.Write RESPUESTA_OK &"|"& cuitCliente &"|"& cdVendedor &"|"& cdCorredor
                    else
                        Response.Write "El Corredor/Vendedor no pueden ser utilizados para el Destinatario Elegido."
                    end if                
                else
                   Response.Write "La cantidad excede el maximo diario."
                end if
            end if            
    end select
else
    logMig.info("Accion Rechazada debido a que el usuario notiene permisos LCK para operar")
    Response.Write "LCK"
end if    
logMig.info("****************************************** FINALIZA *******************************************************")
%>