<%
Const PROV_MULTILATERAL = "M"

'Estados del proveedor
Const ESTADO_HABILITADO = ""
Const ESTADO_DESHABILITADO = "*"

'Relacion del proveedor.
Const PROV_HEAD_OFFICE = "M"
Const PROV_BRANCH = "S"
Const PROV_OVERSEAS = "E"

'Modo de operacion respecto de los datos de AFIP
Const PROV_AFIP_MANUAL = "M"

'Codigos Impositivos
Const PROV_COD_MONOTRIBUTO = "M"

'Const EMAIL_ALTA = "bacarinie@toepfer.com;"
'Const EMAIL_BAJA = "bacarinie@toepfer.com;"
'Const EMAIL_MODI = "bacarinie@toepfer.com;"
'Const EMAIL_IMPUESTOS = "bacarinie@toepfer.com;"
'Const EMAIL_ERROR = "bacarinie@toepfer.com;"
Const PROV_SEC_NVA = 1
Const FILE_TYPE_NVA = "NVA"
Const FILE_TYPE_FILE = "FILE"
Const FILE_TYPE_LOGO = "LOGO"

Const LISTA_PRV_LEGALES = "LEGAL_PROV"

Const EXCENTO = "E"
Const NO_CUIT = "F"
Const NO_INSCRIPTO = "N"
Const PROV_PROFORMA = "P"
Const TABLA_CTZ_CABECERA     = "TBLCTZCABECERA"
Const TABLA_PCP_DETALLE      = "TBLPCPDETALLE"
Const TABLA_PCT_COTIZACIONES = "TBLPCTCOTIZACIONES"
Const TABLA_PCT_PROVEEDORES  = "TBLPCTPROVEEDORES"
Const TABLA_REM_CABECERA     = "TBLREMCABECERA"
Const TABLA_PCT_CABECERA	 = "TBLPCTCABECERA"

Dim nropro,razsoc,domici,dslocali,locali,codpos,codpro,tiprov,sector,sucurs,nomamp,tipdoc
dim emplea, sochec, cooper, profor
Dim nrodoc,codiga,codiva,insiga,nroibr,nrocml,controlAFIP,dirext,locext,visimp,marmal,opcion
Dim	fecalt,fecbaj,estado, reqLegales, habContratos
Dim peribr,filler,secto(9), provincia, usoRRHH
Dim cantFilasMulti,itemsMulti()
Dim TABLA_MAESTRO_PROV_UNIFICADO, TABLA_MAESTRO_PROFORMAS, TABLA_MAESTRO_PROVEEDORES


Dim prov_rsSectores
prov_rsSectores 	= null
Dim prov_rsProvincias
prov_rsProvincias 	= null

'Const TABLA_MAESTRO_PROV_UNIFICADO = "merfl.tcb6a1f1" '<- vista logica 
TABLA_MAESTRO_PROVEEDORES = "QS36F."&CHR(034)&"TG.6A1F1"&CHR(034) 'QS36F.TG.6A1F1 <- archivo fisico 
TABLA_MAESTRO_PROV_UNIFICADO = "VWEMPRESAS" 
TABLA_MAESTRO_PROFORMAS = "TBLEMPRESAS" 
TABLA_MAESTRO_PROVEEDORES_LOG = "PROVFL.LOG6A1F1" 

'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Grabar los datos actuales del proveedor en el LOG de cambios.
' Parametros:
'				pNroPro
' Devuelve:
'				nada
'--------------------------------------------------------------------------------------------------
Function grabarProveedorLog(pNroPro)
    Dim strSQL, rs, strParams
    
    Call GP_ConfigurarMomentos()    
    strParams = pNroPro & "||" & session("Usuario") & "||" & Left(session("MmtoDato"), 8) & "||" & Right(session("MmtoDato"), 6)
    Call executeSP(rs, "PROVFL.LOG6A1F1_INS", strParams)

End Function
'--------------------------------------------------------------------------------------------------
function esProforma(pNroProveedor)
    esProforma = false
    if ((pNroProveedor >= 100000) or (pNroProveedor = 0)) then
		esProforma = true
	else
		'JAS -if nrodoc = "" then loadDataDB(pNroProveedor)
		if ((profor = PROV_PROFORMA) and (estado=ESTADO_DESHABILITADO)) then esProforma = true		
	end if
end function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Reactivar un proveedor
' Parametros:
'				pNroPro
' Devuelve:
'				nada
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function reactivarProveedor(pNroPro)
	Dim strSQL, rs
	
	Call loadDataDB(pNroPro)

	'JAS - fecbaj = 0
	profor = PROV_PROFORMA 
	Call updateProveedor()
	
	'Quitar Firmas siempre	
	Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLEMPRESASFIRMAS_DEL", pNroPro)
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Graba los datos del proveedor
' Parametros:
'				ninguno
' Devuelve:
'				[array]
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function grabarProforma()
	Dim rs, rtrn, strParams, sp_ret
	
	Call GP_ConfigurarMomentos()
	fecalt = mid(session("MmtoDato"),3,6)

	strParams =  ucase(razsoc) & "||" & ucase(domici) & "||" & left(dslocali,14) & "||" & codpos & "||" & codpro & "||" &_
	            tiprov & "||" & sector & "||" & emplea & "||" & sochec & "||" & cooper & "||" & sucurs & "||" & ucase(nomamp) & "||" &_
	            tipdoc & "||" & nrodoc & "||" & codiga & "||" & codiva & "||" & insiga & "||" & nroibr & "||" & nrocml & "||" &_
	            fecalt & "||" & peribr & "||" & secto(2) & "||" & secto(3) & "||" & secto(4) & "||" & secto(5) & "||" & secto(6) & "||" &_ 
	            secto(7) & "||" & secto(8) & "||" & secto(9) & "||" & habContratos & "||" & UCase(dirext) & "||" & UCase(locext) & "||" &_ 
	            visimp & "||" & marmal & "||" & controlAFIP & "$$IDPROVEEDOR"
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TOEPFERDB.TBLEMPRESAS_INS", strParams)
	rtrn = sp_ret("IDPROVEEDOR")
	grabarProforma = rtrn
	
End Function
'--------------------------------------------------------------------------------------------------
Function Var2Number(pVar)
    Dim rtrn
    rtrn = pVar
	if (rtrn = "") then rtrn = 0
    var2Number = rtrn	
End Function
'--------------------------------------------------------------------------------------------------
Function grabarProveedor()
	Dim strSQL,rs,conn, newPro

	strSQL = "select max(NROPRO) NROPRO from " & TABLA_MAESTRO_PROVEEDORES 
	Call executeQuery(rs, "OPEN", strSQL)
	newPro = 1
	if not rs.EoF then newPro = cdbl(rs("NROPRO"))+1	
	fecalt = mid(session("MmtoDato"),3,6)
	
	strSQL = "insert into " & TABLA_MAESTRO_PROVEEDORES  & " ("
	strSQL = strSQL & "     NROPRO,RAZSOC,DOMICI,"
	strSQL = strSQL & "		LOCALI,CODPOS,CODPRO,"
	strSQL = strSQL & "		TELEFO,TIPROV,SECTOR,"
	strSQL = strSQL & "		SUCURS,NOMAMP,TIPDOC,"
	strSQL = strSQL & "		NRODOC,CODIGA,CODIVA,"
	strSQL = strSQL & "		INSIGA,NROIBR,NROCML,"
	strSQL = strSQL & "		DIREXT,LOCEXT,EX3125,"
	strSQL = strSQL & "		VISIMP,MARMAL,OPCION,"
	strSQL = strSQL & "		FECALT,FECBAJ,ESTADO,"
	strSQL = strSQL & "		SECTO2,SECTO3,SECTO4,"
	strSQL = strSQL & "		SECTO5,SECTO6,SECTO7,"
	strSQL = strSQL & "		SECTO8,SECTO9,SECT10,"
	strSQL = strSQL & "		PERIBR,FILLER) "
	strSQL = strSQL & "	values ("
	strSQL = strSQL & "		 "&newPro&" ,'"&ucase(razsoc)&"','"&ucase(domici)&"',"
	strSQL = strSQL & "		'"&left(dslocali,14)&"', "&codpos&" ,'"&codpro&"',"
	strSQL = strSQL & "		 "&Var2Number(telefo)&" ,'"&tiprov&"','"&sector&"',"
	strSQL = strSQL & "		'"&sucurs&"','"&ucase(nomamp)&"', "&tipdoc&" ,"
	strSQL = strSQL & "		 "&nrodoc&" ,'"&codiga&"','"&codiva&"',"
	strSQL = strSQL & "		'"&insiga&"', "&nroibr&" , "&nrocml&" ,"
	strSQL = strSQL & "		'"&ucase(dicext)&"','"&ucase(locext)&"','" & profor & "',"
	strSQL = strSQL & "		'"&visimp&"','"&marmal&"','',"
	strSQL = strSQL & "		 "&fecalt&" , "&fecbaj&" ,'"&estado&"',"
	strSQL = strSQL & "		'"&secto(2)&"','"&secto(3)&"','"&secto(4)&"',"
	strSQL = strSQL & "		'"&secto(5)&"','"&secto(6)&"','"&secto(7)&"',"
	strSQL = strSQL & "		'"&secto(8)&"','"&secto(9)&"','" & habContratos & "',"
	strSQL = strSQL & "		'"&peribr&"','') "
	Call executeQuery(rs, "EXEC", strSQL)
		
	if (cdbl(tipdoc) = TIPO_CUIT_80) then
		'inserto los datos del convenio multilateral
		if (cantFilasMulti > 0) then Call actualizarMultilateral()
	end if

    Call actualizarDatosComplementarios(newPro)	
	
	'Se actualiza el nro de proveedor en la firma de legales.	
    Call executeQuery(rs, "EXEC", "Update TOEPFERDB.TBLEMPRESASFIRMAS set IDEMPRESA=" & newPro & " WHERE IDEMPRESA = " & nroPro)	
	
	Call grabarProveedorLog(newPro)
	
	Call grabarProveedorPuertos(TERMINAL_ARROYO, newPro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	Call grabarProveedorPuertos(TERMINAL_TRANSITO, newPro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	Call grabarProveedorPuertos(TERMINAL_PIEDRABUENA, newPro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	
	grabarProveedor = newPro
End Function
'--------------------------------------------------------------------------------------------------
Function habilitadoParaContratos(nroPro)
   
    habilitadoParaContratos = false    
    strSQL="Select * from MET001A where nroemp = " & nroPro
    Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)    
    if (not rs.eof) then habilitadoParaContratos = true    
    
End Function
'--------------------------------------------------------------------------------------------------
Function actualizarDatosComplementarios(p_pro)	
    'Grabar datos impositivos en PRV6A1F2
	strSQL = "Select count(*) as QUANTITY from PROVFL.PRV6A1F2 WHERE CDPRR2 = " & p_pro
	Call executeQuery(rs, "OPEN", strSQL)
	if clng(rs("QUANTITY")) = 0 then
		strSQL = "INSERT INTO PROVFL.PRV6A1F2 VALUES(" & p_pro & "," & _
				 "'" & emplea & "','" & usoRRHH & "','" & sochec & "','" & cooper & "')"
	else	
		strSQL = "UPDATE PROVFL.PRV6A1F2 SET ESEMR2='" & emplea & "'" & _
				", MASHR2='" & sochec & "', COOPER='" & cooper & "', MCTRR2='" & usoRRHH & "'" & _
				" WHERE CDPRR2 = " & p_pro
	end if
	Call executeQuery(rs, "EXEC", strSQL)
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Da de baja un proveedor
' Parametros:
'				[int] pNroPro
' Devuelve:
'				nada
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function bajaProveedor(pNroPro)
	Dim rs,mensaje

	if (esProforma(pNroPro) and (profor <> PROV_PROFORMA)) then		
		Call executeSP(rs, "TOEPFERDB.TBLEMPRESAS_DEL", pNroPro)
	else
	    Call GP_ConfigurarMomentos()
	    Call loadDataDB(pNroPro)
	    estado = ESTADO_DESHABILITADO
	    profor = ""
	    'Siempre que se da de baja a mano, se cambia a manual el control de datos de AFIP.
	    controlAFIP = PROV_AFIP_MANUAL
	    fecbaj = mid(session("MmtoDato"),3,6)
    	Call updateProveedor()	
	end if	
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Actualiza el proveedor
' Parametros:
'				ninguno
' Devuelve:
'				nada
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function updateProveedor()
	Dim strParams,rs
	
	strParams = nropro & "||" & ucase(razsoc) & "||" & ucase(nomamp) & "||" & tiprov & "||" & codpro & "||" & left(dslocali,14) & "||" & ucase(domici) &_
	            "||" & codpos & "||" & sector & "||" & sucurs & "||" & codiga & "||" & codiva & "||" & insiga & "||" & nroibr & "||" & nrocml & "||" & controlAFIP &_	
	            "||" & profor & "||" & ucase(dirext) & "||" & ucase(locext) & "||" & visimp & "||" & marmal & "||" & fecalt & "||" & fecbaj & "||" & estado &_
	            "||" & secto(2) & "||" & secto(3) & "||" & secto(4) & "||" & secto(5) & "||" & secto(6) & "||" & secto(7) & "||" & secto(8) & "||" & secto(9) &_
	            "||" & habContratos & "||" & peribr		
	Call executeSP(rs, "PROVFL.QS36F_TG6A1F1_UPD", strParams)
	Call grabarProveedorLog(nropro)	

    if (profor <> PROV_PROFORMA) then
	    Call actualizarDatosComplementarios(nropro)	
	    Call grabarProveedorPuertos(TERMINAL_ARROYO, nropro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	    Call grabarProveedorPuertos(TERMINAL_TRANSITO, nropro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	    Call grabarProveedorPuertos(TERMINAL_PIEDRABUENA, nropro, ucase(razsoc), ucase(domici), getDsTipoDoc(tipdoc), nrodoc, estado)
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Ezequiel Bacarini
' Fecha: 18/10/2013
' Objetivo:
'				Actualiza el proveedor provisorio
' Parametros:
'				-
' Devuelve:
'				-
'--------------------------------------------------------------------------------------------------
Function updateProformaProveedor()
	Dim strSQL,rs,conn
	
	strSQL = "UPDATE " & TABLA_MAESTRO_PROFORMAS & " set "
	strSQL = strSQL & "DSEMPRESA = '"&ucase(razsoc)&"', "
	strSQL = strSQL & "DSLEGAL = '"&ucase(nomamp)&"', "
	strSQL = strSQL & "TIPOEMPRESA = '"&tiprov&"', "
	strSQL = strSQL & "PROVINCIA = '"&codpro&"', "
	strSQL = strSQL & "LOCALIDAD = '"&left(dslocali,14)&"', "
	strSQL = strSQL & "DOMICILIO = '"&ucase(domici)&"', "
	strSQL = strSQL & "CODIGOPOSTAL =  "&codpos&", "
	strSQL = strSQL & "SECTOR = '"&sector&"', "
	strSQL = strSQL & "EMPLEADOR = '"&emplea&"', "
	strSQL = strSQL & "SOCHECHO = '"&sochec&"', "
	strSQL = strSQL & "COOPERATIVA = '"&cooper&"', "			
	strSQL = strSQL & "SUCURSAL = '"&sucurs&"', "
	strSQL = strSQL & "CODIGOIGA = '"&codiga&"', "
	strSQL = strSQL & "CODIGOIVA = '"&codiva&"', "
	strSQL = strSQL & "INSCRIPCIONIGA = '"&insiga&"', "
	strSQL = strSQL & "NROIIBB =  "&nroibr&", "
	strSQL = strSQL & "NROCML =  "&nrocml&", "
	strSQL = strSQL & "FECHAALTA =  "&fecalt&", "
	strSQL = strSQL & "FECHABAJA =  "&fecbaj&", "
	strSQL = strSQL & "PERCEPCIONIIBB = '"&peribr&"', "
	strSQL = strSQL & "SECTO2 = '"&secto(2)&"', "
	strSQL = strSQL & "SECTO3 = '"&secto(3)&"', "
	strSQL = strSQL & "SECTO4 = '"&secto(4)&"', "
	strSQL = strSQL & "SECTO5 = '"&secto(5)&"', "
	strSQL = strSQL & "SECTO6 = '"&secto(6)&"', "
	strSQL = strSQL & "SECTO7 = '"&secto(7)&"', "
	strSQL = strSQL & "SECTO8 = '"&secto(8)&"', "
	strSQL = strSQL & "SECTO9 = '"&secto(9)&"', "
	strSQL = strSQL & "SECT10 = '" & habContratos & "', "
	strSQL = strSQL & "DOMICILIOEXT = '"&ucase(dirext)&"', "
	strSQL = strSQL & "LOCALIDADEXT = '"&ucase(locext)&"', "
	strSQL = strSQL & "VISTOIMPUESTOS = '"&visimp&"', "
	strSQL = strSQL & "CONPROBLEMAS = '"&marmal & "', "
	strSQL = strSQL & "CONTROLAFIP = '"&controlAFIP & "'"
	strSQL = strSQL & " where IDEMPRESA = '"&nropro&"'"
	
	Call executeQuery(rs, "EXEC", strSQL)		
    
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Actualiza el convenio Multilateral
' Parametros:
'				ninguno
' Devuelve:
'				nada
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function actualizarMultilateral()
	Dim strSQL,rs, conn, grabar
	
	    strSQL = "delete from provfl.prv6a1f4 where IIBCUI = " &  nrodoc
	    Call executeQuery(rs, "EXEC", strSQL)
    		
    	grabar = false
	    strSQL = "insert into provfl.prv6a1f4 (IIBCUI,IIBPRV,IIBCOF) values "	    
	    for i = 0 to  cantFilasMulti-1
		    if (trim(itemsMulti(i+1,1)) <> "") then
		        grabar = true
			    strSQL = strSQL & "("&nrodoc&",'"&itemsMulti(i+1,1)&"',"&itemsMulti(i+1,2)&"),"
		    end if
	    next    	
	    strSQL = left(strSQL, len(strSQL)-1) ' le quito la ultima coma
	    if (grabar) then Call executeQuery(rs, "EXEC", strSQL)
	
End Function
'------------------------------------------------------------------------------------------
Function isCuitEnabledAFIP(pNro)
dim rtrn
	rtrn = false
	strSQL = "Select * from DGI.DGI600F1 where NDOCR1 = '" & pNro & "'"
	Call executeQuery(rs, "OPEN", strSQL)
	if (not rs.EoF) then rtrn = true
    isCuitEnabledAFIP = rtrn
end function
'------------------------------------------------------------------------------------------
Function existeRegistrado(pNro, pCuit)
Dim strSQL,rs,conn
	rtrn = false
	strSQL = "select * from " & TABLA_MAESTRO_PROV_UNIFICADO & " where CUIT = " & pCuit & " and SUCURSAL not in ('" & PROV_BRANCH & "')"
	if (pNro <> 0) then strSQL = strSQL & " AND IDEMPRESA<>" & pNro
	Call executeQuery(rs, "OPEN", strSQL)
	if (not rs.EoF) then rtrn = true
    existeRegistrado = rtrn	
end function
'------------------------------------------------------------------------------------------
Function tieneNVA(pNro)
Dim strSQL,rs,conn
	rtrn = false
	strSQL = "select * from TOEPFERDB.TBLPROVEEDORESARCHIVOS where IDPROVEEDOR = " & pNro & " and SECUENCIA=1"
	Call executeQuery(rs, "OPEN", strSQL)
	if (not rs.EoF) then rtrn = true
tieneNVA = rtrn	
end function
'--------------------------------------------------------------------------------------------------
' Autor: Ezequiel Bacarini
' Fecha: 18/10/2013
' Objetivo:
'				Cargo toda la informacion de una proforma desde la base de datos
' Parametros:
'				[int] pNroPro
' Devuelve:
'				[boolean] booleano que indica la existancia del proveedor
'--------------------------------------------------------------------------------------------------
Function loadDataDB(pNroPro)
	Dim strSQL,rs
	rtrn = false
	strSQL = "select * from " & TABLA_MAESTRO_PROV_UNIFICADO & " where IDEMPRESA = " & pNroPro	
	Call executeQuery(rs, "OPEN", strSQL)
	if (not rs.EoF) then
	    Call loadSupplierGlobalData(rs)
		'Se levantan los datos de 
		if (cdbl(tipdoc) = TIPO_CUIT_80) then							
		    strSQL = "select IIBPRV idprov,IIBCOF coef,DESCPO dsprov from provfl.prv6a1f4 "
		    strSQL = strSQL & " inner join MERFL.MER1K2F1 dsprov on iibprv = codipo"
		    strSQL = strSQL & " where IIBCUI = " & nrodoc
		    strSQL = strSQL & " order by IIBCOF desc"
    				
		    Call executeQuery(rs,"OPEN", strSQL)
    		cantFilasMulti = 0	
		    while (not rs.EoF)
			    cantFilasMulti = cantFilasMulti + 1
			    'Response.Write "A(" & cantFilasMulti & ")1(" & ubound(itemsMulti) & ")"
    			Redim itemsMulti(cantFilasMulti,2)			        					
		        itemsMulti(cantFilasMulti,0) = trim(rs("dsprov"))
			    itemsMulti(cantFilasMulti,1) = trim(rs("idprov"))
			    itemsMulti(cantFilasMulti,2) = trim(rs("coef"))
			    rs.MoveNext			    
		    wend
	    end if
		rtrn = true
	end if	
	loadDataDB = rtrn
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Javier Scalisi
' Fecha: 22/09/2014
' Objetivo:
'				Se cargan los resultados de un recordset en las variables globales del proveedor.
' Parametros:
'				[int] rs    Recordset de datos.
'--------------------------------------------------------------------------------------------------
Function loadSupplierGlobalData(rs)
    if (not rs.EoF) then
        nroPro = rs("IDEMPRESA")
		razsoc = trim(rs("DSEMPRESA"))
		nomamp = trim(rs("DSLEGAL"))
		tiprov = trim(rs("TIPOEMPRESA"))
		nrodoc = trim(rs("CUIT"))
		tipdoc = trim(rs("TIPODOCUMENTO"))		
		codpro = trim(rs("PROVINCIA"))
		dslocali = trim(rs("LOCALIDAD"))
		domici = trim(rs("DOMICILIO"))
		codpos = trim(rs("CODIGOPOSTAL"))
		sector = trim(rs("SECTOR"))
		emplea = trim(rs("EMPLEADOR"))
		profor = trim(rs("PROFOR"))
		if emplea = "" then emplea = TIPO_NEGACION	
		sochec = trim(rs("SOCHECHO"))
		if sochec = "" then sochec = TIPO_NEGACION
		cooper = trim(rs("COOPERATIVA"))
		if cooper = "" then cooper = TIPO_NEGACION		
		sucurs = trim(rs("SUCURSAL"))
		nropro = trim(rs("IDEMPRESA"))
		codiga = trim(rs("CODIGOIGA"))
		codiva = trim(rs("CODIGOIVA"))
		insiga = trim(rs("INSCRIPCIONIGA"))
		nroibr = trim(rs("NROIIBB"))
		nrocml = trim(rs("NROCML"))
		dirext = trim(rs("DOMICILIOEXT"))
		locext = trim(rs("LOCALIDADEXT"))
		visimp = trim(rs("VISTOIMPUESTOS"))
		marmal = trim(rs("CONPROBLEMAS"))		
		fecalt = trim(rs("FECHAALTA"))
		fecbaj = trim(rs("FECHABAJA"))
		estado = trim(rs("ESTADO"))
		secto(2) = trim(rs("SECTO2"))
		secto(3) = trim(rs("SECTO3"))
		secto(4) = trim(rs("SECTO4"))
		secto(5) = trim(rs("SECTO5"))
		secto(6) = trim(rs("SECTO6"))
		secto(7) = trim(rs("SECTO7"))
		secto(8) = trim(rs("SECTO8"))
		secto(9) = trim(rs("SECTO9"))
		habContratos= trim(rs("SECT10"))
		peribr = trim(rs("PERCEPCIONIIBB"))
		controlAFIP = trim(rs("CONTROLAFIP"))	
		usoRRHH = trim(rs("USORRHH"))	
	end if
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Cargo los datos del proveedor desde la URL
' Parametros:
'				ninguno
' Devuelve:
'				nada
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function loadProvParameters()
	razsoc = ucase(replace(GF_PARAMETROS7("razsoc","",6),"'",""))
	nomamp = ucase(replace(GF_PARAMETROS7("nomamp","",6),"'",""))
	tiprov = GF_PARAMETROS7("tiprov","",6)
	nrodoc = GF_PARAMETROS7("nrodoc","",6)
	tipdoc = trim(GF_PARAMETROS7("tipdoc",0 ,6))
	codpro = GF_PARAMETROS7("codpro","",6)
	locali = GF_PARAMETROS7("locali","",6)
	dslocali= GF_PARAMETROS7("dslocali","",6)
	domici = ucase(replace(GF_PARAMETROS7("domici","",6),"'",""))
	codpos = GF_PARAMETROS7("codpos",0,6)
	sector = GF_PARAMETROS7("sector","",6)
	emplea = GF_PARAMETROS7("emplea","",6)
	if emplea = "" then emplea = TIPO_NEGACION
	sochec = GF_PARAMETROS7("sochec","",6)
	if sochec = "" then sochec = TIPO_NEGACION
	cooper = GF_PARAMETROS7("cooper","",6)
	if cooper = "" then cooper = TIPO_NEGACION
	'provincia = GF_PARAMETROS7("provincia","",6)
	sucurs = GF_PARAMETROS7("sucurs","",6)
	if (sucurs = "") then sucurs = PROV_HEAD_OFFICE
	nropro = GF_PARAMETROS7("nropro",0 ,6)
	codiga = GF_PARAMETROS7("codiga","",6)
	codiva = GF_PARAMETROS7("codiva","",6)
	insiga = GF_PARAMETROS7("insiga","",6)
	nroibr = GF_PARAMETROS7("nroibr",0 ,6)
	nrocml = GF_PARAMETROS7("nrocml",0 ,6)
	dirext = ucase(GF_PARAMETROS7("dirext","",6))
	locext = ucase(GF_PARAMETROS7("locext","",6))
	visimp = GF_PARAMETROS7("visimp","",6)
	marmal = GF_PARAMETROS7("marmal","",6) 
	if (marmal = "on") then marmal = ESTADO_DESHABILITADO	
	fecalt = GF_PARAMETROS7("fecalt",0 ,6)
	fecbaj = GF_PARAMETROS7("fecbaj",0 ,6)
	estado = GF_PARAMETROS7("estado","",6)
	profor = GF_PARAMETROS7("profor","",6)
	secto(2) = GF_PARAMETROS7("secto2","",6)
	secto(3) = GF_PARAMETROS7("secto3","",6)
	secto(4) = GF_PARAMETROS7("secto4","",6)
	secto(5) = GF_PARAMETROS7("secto5","",6)
	secto(6) = GF_PARAMETROS7("secto6","",6)
	secto(7) = GF_PARAMETROS7("secto7","",6)
	secto(8) = GF_PARAMETROS7("secto8","",6)
	secto(9) = GF_PARAMETROS7("secto9","",6)	
	habContratos= GF_PARAMETROS7("habContratos","",6)
	peribr = GF_PARAMETROS7("peribr","",6)
	controlAFIP = GF_PARAMETROS7("controlAFIP","",6)	
	usoRRHH = GF_PARAMETROS7("usoRRHH","",6)	
	cantFilasMulti = GF_PARAMETROS7("cantFilasMulti",0,6)
	reqLegales = GF_PARAMETROS7("reqLegales","",6)
	
	Redim itemsMulti(cantFilasMulti,2)
	
	for i = 1 to cantFilasMulti
		itemsMulti(i,0) = GF_PARAMETROS7("provMulti"&i,"",6)
		itemsMulti(i,1) = GF_PARAMETROS7("idProvMulti"&i,"",6)
		itemsMulti(i,2) = GF_PARAMETROS7("coefMulti"&i,"",6)
	next    
	'response.end
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Obtener un recorset generico
' Parametros:
'				[str]	pCodigo - campo de clave unica del registro
'				[str]	pDesc	- campo de la descripcion
'				[str]	pTabla	- La tabla donde se ejecutara la sql
' Devuelve:
'				[recordset]
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function rsGeneral(pCodigo,pDesc,pTabla)
	Dim rs,con,strSQL
		
	strSQL = "select "&pCodigo&" CODIGO, "&pDesc&" DESC from "&pTabla&" order by " & pDesc
	Call executeQuery(rs, "OPEN", strSQL)
	
	Set rsGeneral = rs
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Obtener un recordset con los Sectores
' Parametros:
'				ninguno
' Devuelve:
'				[recordset]
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function getRSSectores()
	if (not isnull(prov_rsSectores))then
		prov_rsSectores.MoveFirst
	else
		Set prov_rsSectores = rsGeneral("D1BGCD","D1HCTX","provfl.ACD1REP")
	end if
	
	Set	getRSSectores = prov_rsSectores 
End Function
'--------------------------------------------------------------------------------------------------
Function getRSProvincias()
	if (not isnull(prov_rsProvincias))then
		prov_rsProvincias.MoveFirst
	else
		Set prov_rsProvincias = rsGeneral("CODIPO","DESCPO","MERFL.MER1K2F1")
	end if
	
	Set	getRSProvincias = prov_rsProvincias 
End Function
'--------------------------------------------------------------------------------------------
Function getIIBB(pProveedor)
dim strSQL, rsIIBB, rtrn
rtrn = ""
strSQL = "SELECT NROIBR, NROCML FROM MERFL.TCB6A1F1 WHERE NROPRO=" & pProveedor
call GF_BD_COMPRAS(rsIIBB, con, "OPEN", strSQL)
if not rsIIBB.eof then
	if cdbl(rsIIBB("NROIBR")) <> 0 then 
		rtrn = rsIIBB("NROIBR")
	else	
		rtrn = rsIIBB("NROCML")
	end if	
end if
getIIBB = trim(rtrn)
End Function 
'--------------------------------------------------------------------------------------------
Function getRSTiposProv()
	Set	getRSTiposProv = rsGeneral("DJASST","DJGATX","provfl.acdjrep")
End Function 
'--------------------------------------------------------------------------------------------
Function getRSTiposDoc()
	Set	getRSTiposDoc = rsGeneral("DGCDOC","DGF7TX","provfl.acdgrep")
End Function 
'-------------------------------------------------------------------------------------------
Function getRSProvincias()
	set getRSProvincias = rsGeneral("CODIPO","DESCPO","MERFL.MER1K2F1")
End Function
'-------------------------------------------------------------------------------------------
Function getRSTipoIGA()
	set getRSTipoIGA = rsGeneral("CDIMR1","DESCR1","dgi.dgi601f1")
End Function 
'-------------------------------------------------------------------------------------------
Function getRSTipoIVA()
	set getRSTipoIVA = rsGeneral("CDIMR1","DESCR1","dgi.dgi601f1")	
End Function
'-------------------------------------------------------------------------------------------
Function getRSParentezcos()
	set getRSParentezcos = rsGeneral("id","descripcion","toepferdb.tblparentezcos")
End Function
'-------------------------------------------------------------------------------------------
Function getRSPaises()
	set getRSPaises = rsGeneral("id","descripcion","toepferdb.tblparentezcos")
End Function
'-------------------------------------------------------------------------------------------
Function getRsFamiliares(pNroPro)
	Dim strSQL,rs,con
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPROVEEDORESFAMILIA_GET_BY_IDPROVEEDOR", pNroPro)
	Set getRsFamiliares = rs	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli
' Fecha: 23/11/2011
' Objetivo:
'				Obtener un recordset con los archivos del proveedor
' Parametros:
'				ninguno
' Devuelve:
'				[recordset]
' Modificaciones:
'                                   --/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function getRsArchivos(pNroPro, pTipo)
	Dim rs
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLPROVEEDORESARCHIVOS_GET_BY_PARAMETERS", pNroPro & "||" & pTipo)
	Set getRsArchivos = rs
End Function
'-------------------------------------------------------------------------------------------
Function getDsSector(pId)
    Dim sp_p, rt, rs, sp_rt
    rt =""    
    if (pId <> "") then
        sp_p = "PROVFL.ACD1REP||D1BGCD||" & pId & "||T||D1HCTX$$DESC"
        Set sp_rt =  executeProcedureDb(DBSITE_SQL_INTRA, rs, "READ_DS_BY_KEY", sp_p)        
        rt = sp_rt("DESC")        
    end if
    getDsSector = rt
End Function
'-------------------------------------------------------------------------------------------
Function getDsTipoProv(pId)
    Dim sp_p, rt, rs, sp_rt
    rt =""    
    if (pId <> "") then
        sp_p = "PROVFL.ACDJREP||DJASST||" & pId & "||T||DJGATX$$DESC"        
        Set sp_rt =  executeProcedureDb(DBSITE_SQL_INTRA, rs, "READ_DS_BY_KEY", sp_p) 
        rt = sp_rt("DESC")        
    end if
	getDsTipoProv = rt
	
End Function
'-------------------------------------------------------------------------------------------
Function getDsTipoDoc(pId)
    Dim sp_p, rt, rs, sp_rt
    
    rt =""    
    if (pId <> "") then
        sp_p = "PROVFL.ACDGREP||DGCDOC||" & pId & "||N||DGF7TX$$DESC"                
        Set sp_rt =  executeSP(rs, "READ_DS_BY_KEY", sp_p)        
        rt = sp_rt("DESC")        
    end if
	getDsTipoDoc = rt
End Function
'-------------------------------------------------------------------------------------------
Function getDsProvincia(pId)
    Dim sp_p, rt, rs, sp_rt
    
    rt =""    
    if (pId <> "") then
        sp_p = "PROVFL.ACDGREP||DGCDOC||" & pId & "||T||DGF7TX$$DESC"
        Set sp_rt =  executeSP(rs, "READ_DS_BY_KEY", sp_p)
        rt = sp_rt("DESC")        
    end if
	getDsProvincia = rt
End Function
'-------------------------------------------------------------------------------------------
Function getDsLocalidad(pId,pIdProv)
	Dim rs,con,strSQL,rtrn
	
	strSQL = "select DESCPC DESC from MERFL.MER142F1 where CODIPC = " & pId & " and auxipc = " & pIdProv
	Call executeQuery(rs, "OPEN", strSQL)

	rtrn = ""
	if (not rs.EoF) then rtrn = rs("DESC")
	
	getDsLocalidad = trim(rtrn)
	
End Function
'-------------------------------------------------------------------------------------------
function getAlicuotaPercepcion(pCUIT, pFecha)
	Dim rs,con,strSQL,rtrn, myAnioMes
	myAnioMes = left(GF_DTE2FN(pFecha),6)
	strSQL = "SELECT PAALPE FROM PROVFL.PRV301F2 WHERE PACUIT = " & pCUIT & " AND PAFDES LIKE '" & myAnioMes & "%'"
	Call executeQuery(rs, "OPEN", strSQL)
	rtrn = "SIN DATOS"
	if (not rs.EoF) then rtrn = rs("PAALPE")
	getAlicuotaPercepcion = trim(rtrn)
end function
'-------------------------------------------------------------------------------------------
Function getDsTipoIGA(pId)
    Dim sp_p, rt, rs, sp_rt
    
    rt =""    
    if (pId <> "") then
         sp_p = "DGI.DGI601F1||CDIMR1||" & pId & "||T||DESCR1$$DESC"
        Set sp_rt =  executeProcedureDb(DBSITE_SQL_INTRA, rs, "READ_DS_BY_KEY", sp_p)
        rt = sp_rt("DESC")        
    end if
	getDsTipoIGA = rt    
End Function
'-------------------------------------------------------------------------------------------
Function validarExp(pExp,pTexto)
	Dim expReg,rtrn
	
	set expReg = New RegExp
	expReg.Pattern = pExp
	rtrn = expReg.Test(pTexto)
	set expReg = nothing
	
	validarExp = rtrn
End Function

'------------------------------------------------------------------------------------------
Function getDsParentezco(pIdParentezco)
	Dim strSQL,rs,con,rtrn
	
	strSQL = "select descripcion from tblparentezcos where id = " & pIdParentezco
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	rtrn = ""
	if (not rs.EoF) then
		rtrn = rs("descripcion")
	end if
	
	getDsParentezco = rtrn
End Function
'------------------------------------------------------------------------------------------
Function controlDatosProveedor(pDefinitivo)
dim rtrn 
rtrn = false
    '*** Controles obligatorios para proveedores y proformas.
    if (razsoc = "") then setError(SIN_NOMBRE)    
    if ((CDbl(tipdoc) = TIPO_CUIT_80) or (CDbl(tipdoc) = TIPO_CUIT_EX_83) or (CDbl(tipdoc) = TIPO_CUIL_86)) then
        if (nrodoc = "" or nrodoc = 0) then setError(SIN_DOCUMENTO)
	    'Se verifica que tenga el formato correcto (Validaci�n num�rica)
	    if (CDbl(tipdoc) <> TIPO_CUIT_EX_83) then
		    if (not GF_CONTROL_CUIT(nrodoc)) then 
			    Call setError(CUIT_ERRONEO)
		    else			
			    'El CUIT es correcto, controlar que este habilitado por AFIP
			    if (isCuitEnabledAFIP(nrodoc)) then
				    if (sucurs <> PROV_BRANCH) then 				
					    if (existeRegistrado(nroPro, nrodoc)) then
						    Call setError(PROV_CUIT_REGISTRADO) 'El CUIT esta registrado
					    end if
				    end if	
			    else
				    Call setError(PROV_CUIT_INHABILITADO) 'No habilitado por AFIP
			    end if	
		    end if
		end if	
	else
	    if (cint(tipdoc) = cint(SIN_SELECCION)) then setError(SIN_TIPO_DOCUMENTO)
	end if
    '*** Controles obligatorio solo para proveedores
    if (not esProforma(nropro)) or (pDefinitivo = TIPO_AFIRMACION) then
	    if (cstr(tiprov) = cstr(SIN_SELECCION)) then Call setError(SIN_TIPO_PROVEEDOR)	    
		if (domici = "") then setError(SIN_DOMICILIO)		
		if (dslocali = "") then setError(SIN_LOCALIDAD)		
		if (cstr(trim(sector)) = cstr(trim(SIN_SELECCION))) then Call setError(SIN_SECTOR)
		'EAB
		if (cstr(trim(codiga)) = cstr(trim(SIN_SELECCION)))  then Call setError(SIN_TIPO_IGA)
		if (CDbl(tipdoc) = TIPO_CUIT_EX_83) then
			if (cstr(trim(codiga)) = cstr(trim(SIN_SELECCION)))  then codiga = EXCENTO 
			if (cstr(trim(codiva)) = cstr(trim(SIN_SELECCION)))  then codiva = NO_CUIT
			if (cstr(trim(insiga)) = cstr(trim(SIN_SELECCION)))  then insiga = NO_INSCRIPTO
		end if	
		'Campo Provincias debe seleccionar una opcion
		if (cstr(trim(codpro)) = cstr(trim(SIN_SELECCION)) and (CDbl(tipdoc) <> TIPO_CUIT_EX_83)) then Call setError(SIN_PROVINCIA)
		
		'Los c�digos Impositivos deben coincidir con el definido en AFIP, salvo que se pueda operar el proveedor Manualmente
		if ((CDbl(tipdoc) = TIPO_CUIT_80) or (CDbl(tipdoc) = TIPO_CUIT_EX_83) or (CDbl(tipdoc) = TIPO_CUIL_86)) then
			if (controlAFIP <> PROV_AFIP_MANUAL) then Call controlCodigosAFIP(nrodoc, codiva, insiga, emplea, sochec)
		end if
		
		'controles de convenios multilaterales
		if (cstr(trim(nrocml)) <> "" and cstr(trim(nrocml)) <> "0" ) then
			if (cdbl(trim(tipdoc)) <> TIPO_CUIT_80) then Call setError(C_MULTI_SIN_CUIT)
		end if
	
		if (peribr = PROV_MULTILATERAL and (cstr(trim(nrocml))="" or cdbl(nrocml) = 0 ) ) then Call setError(SIN_C_MULTI)
		
		'Controlo que no se repita la misma provincia
		if (cantFilasMulti > 0) then
			for i = 1 to cantFilasMulti
				for j = i+1 to cantFilasMulti
					if (trim(itemsMulti(i,1))<>"") then
						if (itemsMulti(i,1) = itemsMulti(j,1)) then
							Call setError(PROV_REPETIDA)
						end if 
					end if
				next
			next
		end if
	
		for i = 1 to cantFilasMulti
			if ( cdbl(itemsMulti(i,2)) > 9.999 ) then
				Call setError(COEF_SUPERADO)
			end if
			
			'Si se cargo una provincia compruebo que tenga coeficiente
			if ((cstr(itemsMulti(i,2)) = "" or cdbl(itemsMulti(i,2)) = 0) and itemsMulti(i,1)<>"" ) then
				Call setError(PROV_SIN_COEF)
			end if
		next
	end if
	if (not hayError()) then rtrn = true
    
    controlDatos = rtrn
End Function
'------------------------------------------------------------------------------------------
'Funcion responsable de comprobar si los �digos de situacion impositiva coinciden con los indicados por la AFIP.
'Fecha: 15/01/2014
'Autor: Javier A. Scalisi 
Function controlCodigosAFIP(p_cuit, p_codiva, p_insiga, p_emplea, p_sochec)
    Dim strSQL, rs
    strSQL =    "Select B.CDIMR2 CDIGA, C.CDIMR2 CDIVA, D.CDIMR2 CDMONO, EMPLR1, ISOCR1 from DGI.DGI600F1 A " &_
                "inner join DGI.DGI601F2 B on A.IIGAR1 = B.CDEXR2 " &_
                "inner join DGI.DGI601F2 C on A.IIVAR1 = C.CDEXR2 " &_
                "inner join DGI.DGI601F2 D on A.MONOR1 = D.CDEXR2 " &_
                "where NDOCR1 = '" & p_cuit & "'"
    Call executeQuery(rs, "OPEN", strSQL)    
    if (not rs.eof) then
        if (rs("CDMONO") = TIPO_NEGACION) then
            'Valen los codigos de cada tipo de retencion
            if (cstr(trim(p_insiga)) <> cstr(trim(rs("CDIGA")))) then Call setError(COD_IGA_INCORRECTO)			
            if (cstr(trim(p_codiva)) <> cstr(trim(rs("CDIVA")))) then Call setError(COD_IVA_INCORRECTO)		
        else
            'Es monotributista.
            if (cstr(trim(p_insiga)) <> PROV_COD_MONOTRIBUTO) then Call setError(COD_IGA_INCORRECTO)			
            if (cstr(trim(p_codiva)) <> PROV_COD_MONOTRIBUTO) then Call setError(COD_IVA_INCORRECTO)		
        end if        
        if (cstr(trim(p_emplea)) <> cstr(trim(rs("EMPLR1")))) then Call setError(PROV_EMPELA_INCORRECTO)			
        if (cstr(trim(p_sochec)) <> cstr(trim(rs("ISOCR1")))) then Call setError(PROV_SOCHEC_INCORRECTO)

    end if
End function    
'------------------------------------------------------------------------------------------
'Actualiza los datos del proveedor en el puerto especificado.
Function grabarProveedorPuertos(pPTO, pNroPro, pDsPro, pDsDomicilio, pTDoc, pCUIT, pEstado)
    
    Dim strSQL, rs
    
    strSQL="Delete from CORREDORES where CDCORREDOR=" & pNroPro
    Call GF_BD_Puertos(pPTO, rs, "EXEC", strSQL)
      
    strSQL="Insert into CORREDORES(CDCORREDOR,	DSCORREDOR,	DSDOMICILIO, CDTIPODOC,	NUCUIT,	CDESTADO) values(" & pNroPro & ", '" & Left(pDsPro, 50) & "', '" & Left(pDsDomicilio, 40) & "', '" & Left(pTDoc, 10) & "', '" & pCUIT & "', '" & pEstado & "')"
    Call GF_BD_Puertos(pPTO, rs, "EXEC", strSQL)
    'Response.Write "<br>1" & strSQL & " , " & err.Description 
    
    strSQL="Delete from VENDEDORES where CDVENDEDOR=" & pNroPro
    Call GF_BD_Puertos(pPTO, rs, "EXEC", strSQL)

    strSQL="Insert into VENDEDORES(CDVENDEDOR,	DSVENDEDOR,	DSDOMICILIO, NUTELEFONO, CDTIPODOC,	NUDOCUMENTO, DSOBSERVACIONES, CDSUCURSAL, CDESTADO) values(" & pNroPro & ", '" & Left(pDsPro, 50) & "', '" & Left(pDsDomicilio, 40) & "',null, '" & Left(pTDoc, 10) & "', '" & pCUIT & "',null,null, '" & pEstado & "')"
    Call GF_BD_Puertos(pPTO, rs, "EXEC", strSQL)
    'Response.Write "<br>2" & strSQL & " , " & err.Description 
    'Response.End 
End Function
'------------------------------------------------------------------------------------------
Function proveedorADefinitivo(pNroOld,pNroNew)
	
	Dim strSQL,conn,rs,rs2, flag
	
	flag=false
	
		'verifico los PCT
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_PCT_CABECERA)
		'verifico los CTZ
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_CTZ_CABECERA)
		'verifico los PCP
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_PCP_DETALLE)
		'verifico los PCTCotizaciones
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_PCT_COTIZACIONES)
		'verifico los PCTPROVEEDORES
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_PCT_PROVEEDORES)
		'verifico los REMITOS
		'*****************************************************************
		Call ActualizarProveedorTablas(pNroOld,pNroNew,TABLA_REM_CABECERA)
		'verifico los Mails
		'*****************************************************************
		
		'Se actualizan los mails del proveedor si figura con dos c�digos.
		strSQL =          "SELECT a.idempresa idNuevo, a.email mailNuevo, b.idempresa idviejo, b.email mailViejo FROM toepferdb.TBLMAILSCOMPRAS a "
		strSQL = strSQl & " LEFT JOIN toepferdb.TBLMAILSCOMPRAS b "
		strSQl = strSQl & "    ON b.idempresa = " & pNroOld
		strSQl = strSQl & " WHERE a.idempresa = " & pNroNew
		Call GF_BD_COMPRAS(rs2, conn, "OPEN", strSQL)			
		if (not rs2.eof) then				
			'El proveedor tiene 2 codigos (uno nuevo < 100K y otro viejo > 100K)
			strSQL = "UPDATE toepferdb.TBLMAILSCOMPRAS Set email = '" & rs2("mailViejo") & "' WHERE idempresa = " & pNroNew
			Call GF_BD_COMPRAS(rs2, conn, "EXEC", strSQL)
			strSQl = "DELETE FROM toepferdb.TBLMAILSCOMPRAS WHERE idempresa = " & pNroOld			
			Call GF_BD_COMPRAS(rs2, conn, "EXEC", strSQL)					
		else
			'El proveedor tiene solo un c�digo
			strSQL = "UPDATE toepferdb.TBLMAILSCOMPRAS Set idempresa = " & pNroNew & " WHERE idempresa = " & pNroOld
			Call GF_BD_COMPRAS(rs2, conn, "EXEC", strSQL)			
		end if			
						
End Function
'------------------------------------------------------------
Function ActualizarProveedorTablas(pIdViejo, pIdNuevo, pTabla)
	Dim strSQL,rs,rsAux,conn,rtrn

	rtrn = 0
	strSQL = "SELECT * "
	strSQL = strSQL & "FROM   toepferdb." & pTabla & " "
	strSQL = strSQL & "WHERE  idproveedor = " & pIdViejo	
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	if (not rs.eof) then
	'reemplazo por el nuevo ID
		strSQL =          "UPDATE toepferdb." & pTabla & " "
		strSQL = strSQL & " SET    idproveedor  = " & pIdNuevo
		strSQL = strSQL & " WHERE  idproveedor = " & pIdViejo
		Call GF_BD_COMPRAS(rsAux, conn, "EXEC", strSQL)
		rtrn = 1
	end if
	ActualizarProveedorTablas = rtrn		
End Function
'------------------------------------------------------------------------------------------
Function getIndicadorEstado(pMostrarTexto)
	Dim img, txt, txtMostrar
	    
	if (esProforma(nropro)) then
	    img = "activo.png"
	    txt = "Proforma"
	else
	    if (estado = ESTADO_HABILITADO)  then    
	        img = "enable-16.png"
	        txt = "Habilitado"
	    else
            img = "disable-16.png"
            txt = "Dado de Baja"
        end if            
    end if
    if (pMostrarTexto) then txtMostrar = txt
	getIndicadorEstado = "<img style='vertical-align:middle;' src='../images/" & img & "' title='" & txt & "'> " & txtMostrar
End Function

%>
