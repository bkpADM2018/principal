<%
    Const CYPHER_NONE = "NONE-"
    Const CYPHER_VIGENERE = "A001-"

	'Claves de Algoritmos de Encripcion.
    Const VIGENERE_KEY = "RomuLAN Empire Strikes AGAIN in 2064"

	Const ADS_SECURE_AUTHENTICATION=1
	
	
	const PASSPORT_INF_VALID	= 0
	const PASSPORT_NOT_VALID	= 1
	const PASSPORT_OUT_OF_DATE	= 2
	const PASSPORT_BAD_OWNER	= 3
	const PASSPORT_BAD_SYSTEM	= 4

	Const PASSPORT_DATA_TOKEN = "&"
	Const PASSPORT_FIELD_TOKEN = "="
    
    Const DIVISION_ALL = 0
    
    '/* CONTANTES DE SEGURIDAD DE ACCESO */    
	Const TASK_LVL_VW_ACCESS = 1 'VIEW ACCESS
	Const TASK_LVL_WK_ACCESS = 2 'WIRITE/MODIFY ACCESS
    	
    '/***** TAREAS SOBRE LAS CUALES SE APLICA SEGURIDAD  ******/
    '/  Nomeclatura:
    '       TASK_<SISTEMA>_<DESCRIPCION>
    '   Donde:
    '       <SISTEMA>: Abreviatura del sistema al cual pertenece la tarea. No más de 4 letras. (ej: PROV, MER, SYS, FAC, TES, etc)
    '       <DESCRIPCION>: Texto representativo de la tarea, no más de 2 o 3 paralabras o 10-15 caracteres. (usar abreviaturas y en lo posible que sea en ingles!)
    Const TASK_EJE_PROVISIONS = 15
	Const TASK_POS_MT_ZAR_Y_SEC  = 16
	Const TASK_POS_BZA_EMB  = 17
    Const TASK_POS_MERMA_VOLATIL = 18
    Const TASK_POS_CONSULTA_CUPO_PATENTE = 22
    Const TASK_POS_ASOCIAR_CUPO_PATENTE = 23
    Const TASK_POS_LIBERAR_CUPO_PATENTE = 24
    Const TASK_POS_MODIFICACION_HISTORICA = 25
    Const TASK_POS_DESCARGA_TERCEROS = 27
    Const TASK_COM_AUTH_REASSIGNING_BUDGET = 28
    Const TASK_POS_PANEL_PUERTO = 29
    Const TASK_POS_INFO_ANALISIS = 30
    Const TASK_POS_ADMIN_CUPOS = 31
	Const TASK_SMW_REPO_CUMPLIMIENTO = 33
	Const TASK_POS_REPO_CCPP = 34
	Const TASK_POS_CUPOS_TYL = 35
	Const TASK_POS_ADM_AJUSTES = 36
	Const TASK_BZA_CAM_CTRL_PESO = 37
	Const TASK_POS_CTRL_MERMA = 39
	Const TASK_POS_DESC_X_ENT = 40
    '/***** MAIL ******/
    Const TASK_PROV_MAIL_ALERT = 6
    Const TASK_FAC_MAIL_ALERT  = 7
    
    Const LCK_LOGISTICA = "LCKCUPOS"
    '/*********************************************************/   
    
    dim cypherKey() 'Clave del Algoritmo de encripcion.
    dim cypherReady 'Indica si la clave del Algoritmo de Encripcion ya fue genereda.	
    dim PASSPORT_DECRYPTION_KEY
    
    PASSPORT_DECRYPTION_KEY = Chr(210) & Chr(182) & Chr(115) & Chr(116) & Chr(126) & Chr(39) & Chr(7) & Chr(75) & Chr(130) & Chr(191) & Chr(140) & Chr(6) & Chr(175) & Chr(94) & Chr(129) & Chr(249)	
'------------------------------------------------------------------------------------------------
function Encrypt(p_str)
	if (not cypherReady) then GenerateKey(VIGENERE_KEY)
    Encrypt = CYPHER_VIGENERE & Encrypt_Vigenere(p_str)
end function
'------------------------------------------------------------------------------------------------
function Decrypt(p_str)
	dim code, buffer, output

	if (not cypherReady) then GenerateKey(VIGENERE_KEY)

	    buffer = CStr(p_str)
	    code = left(buffer,5)
	    buffer = right(buffer,len(buffer)-5)
	    'code = buffer.Substring(0, 5)
	    'buffer = buffer.Remove(0, 5)
	    select case code
	        case CYPHER_NONE 'No tiene encripcion alguna, pero paso por el encriptor.                    
	            output = buffer
	        case CYPHER_VIGENERE 'Algoritmo Vigenere
	            output = Decrypt_Vigenere(buffer)
	        case Else
	            output = p_str
	    end select

	Decrypt = output
end function
'------------------------------------------------------------------------------------------------
sub GenerateKey(p_key)
Dim str, i, k
str = p_key
k=len(str)

redim cypherKey(k)
while (i < k)
    cypherKey(i) = Asc(left(str,1))
    str = right(str,len(str)-1)
    i = i + 1
wend
cypherReady = True
end sub
'------------------------------------------------------------------------------------------------
function Encrypt_Vigenere(p_str)
dim buffer, i, value, result, k
buffer = p_str
k = len(buffer)
'Se leen todos los caracteres del buffer.
while (i < k)	
    value = webChar((Asc(left(buffer, 1)) + cypherKey(i Mod ubound(cypherKey))) Mod 255)    
    buffer = right(buffer,len(buffer)-1)
    result = result & value
    i = i + 1
wend
Encrypt_Vigenere = result
end function
'------------------------------------------------------------------------------------------------
function Decrypt_Vigenere(p_str)
dim buffer, i, value, result, k
buffer = CStr(p_str)
k = len(buffer)
'Se leen todos los caracteres del buffer.
while (i < k)
    value = Chr((Asc(left(buffer,1)) - cypherKey(i Mod ubound(cypherKey)) + 255) Mod 255)        
    buffer = right(buffer,len(buffer)-1)    
    result = result & value
    i = i + 1
wend
Decrypt_Vigenere = result
end function
'------------------------------------------------------------------------------------------------
'Funcion para autenticar 
Function authenticateWindowsUser(username, password)

On Error Resume Next
	
	'Conecto al dominio
	Set objRootDSE = GetObject("LDAP://rootDSE")	
	strADsPath = "LDAP://" & objRootDSE.Get("defaultNamingContext")
	'Se autentica el ususario	
	szUserId= username
	szPasswd= password

	Set oDSObj = GetObject("LDAP:")
	Set oAuth = oDSObj.OpenDSObject(strADsPath, szUserId, szPasswd, ADS_SECURE_AUTHENTICATION)
	
	if (Err.Number = 0) then
		authenticateWindowsUser = True
	else
		authenticateWindowsUser = False
	End if

End Function
'------------------------------------------------------------------------------------------------
'Funcion que reemplaza al Chr tradicional para evitar que se utilicen caracteres no imprimibles en la encriptacion
Function webChar(p_ascii)
	if ((p_ascii < 33) or _
		(p_ascii > 126) or (p_ascii = 34) or _
		(p_ascii = 39) or (p_ascii = 60) or _
		(p_ascii = 62) or (p_ascii = 96)) then
		value = "%" & Hex(p_ascii)
	else
		value = Chr(p_ascii)
	end if
	webChar = value
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	???
' Fecha: 	--/--/--
' Objetivo:	
'			Decodifica y devuelve la informacion de un pasaporte
' Parametros:
'			pasaporte 	[str] 	pasaporte a leer
'			idTitular 	[int] 	
'			mmtoEmison	[int]	
'			mmtoVto		[int] 	
'			payLoad		[int] 	datos incluidos en el pasaporte para uso del programa receptor.
'			crc			[int] 	codificacion MD5 de los datos'			
' Devuelve:
'			parametros por referencia
' Modificaciones:
'			05/11/10 - GFG
'--------------------------------------------------------------------------------------------------
Function leerPassporte(pasaporte, ByRef sistema, ByRef idTitular, ByRef mmtoEmision, ByRef mmtoVto, ByRef payLoad, ByRef crc)
	Dim fields, pass, readOk, plainPassport
	
	readOk = false
	
	if (pasaporte <> "") then
		'pasaporte2 = Decrypt(pasaporte)
		 pasaporte2 = algorithmRC4(DecodeRC4(Trim(pasaporte)), PASSPORT_DECRYPTION_KEY)
		
		pass = Split(pasaporte2,"$")
		if (UBound(pass) = 1) then crc = pass(1)
		fields = Split(pass(0),"|")	
		if (UBound(fields) = 4) then	
			sistema = fields(0)
			idTitular = fields(1)
			if (idTitular = "") then idTitular=0
			mmtoEmision = fields(2)
			if (mmtoEmision = "") then mmtoEmision = "20000101000000"
			mmtoVto = fields(3)	
			if (mmtoVto = "") then mmtoVto = "20000101000000"
			payLoad = fields(4)						
			readOk = true
		end if
	end if
	if (not readOk) then
		sistema = "ERROR"
		idTitular = 0
		payLoad = ""			
		mmtoEmision = "20000101000000"
		mmtoVto = "20000101000000"
		crc = ""
	end if
End Function
'------------------------------------------------------------------------------------------------
Function generarCRC(sistema, idTitular, validaDesde, validaHasta, payLoad)	
	generarCRC = MD5(idTitular & payLoad & validaDesde & validaHasta & sistema)	
End Function 
'------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	--/--/----
' Objetivo:	
'			Genera un pasaporte nuevo
' Parametros:
'			sistema 	[str] 	Nombre del sistema que emite el pasaporte
'			idTitular 	[int] 	
'			validaDesde	[int]	
'			validaHasta	[int] 	
' Devuelve:
'			Un pasaporte valido.
' Modificaciones:
'			09/11/10 - JAS
Function emitirPasaporte(sistema, idTitular, validaDesde, validaHasta, payLoad) 
	Dim passString, passport , fd, fh, enc
		
	fd=validaDesde	
	if (inStr(validaDesde, "/") <> 0) then fd = GF_DTE2FN(validaDesde)	
	fh=validaHasta	
	if (inStr(validaHasta, "/") <> 0) then fh = GF_DTE2FN(validaHasta)
	
	passString = sistema & "|" & idTitular & "|" & fd & "|" & fh & "|" & payload
	'Response.Write passString
	passport = passString & "$" & generarCRC(sistema, idTitular, fd, fh, payLoad)		
	'Para la web se deben adaptar los caracteres no imprimibles a codigos hexadecimales.
	'emitirPasaporte = Encrypt(passport)	
	emitirPasaporte = encodeRC4(algorithmRC4(passport, PASSPORT_DECRYPTION_KEY))	
End Function
'------------------------------------------------------------------------------------------------
Function validarPasaporte(pSistema, pIdTitular, pasaporte)
	Dim idTitular, validaDesde, validaHasta, sistema
	Dim result, rs, strSQL, codeCRC, dataCRC, payLoad, myRol
	Call leerPassporte(pasaporte, sistema, idTitular, validaDesde, validaHasta, payLoad, codeCRC)
	'result = false	
	'Valido los datos obtenidos del pasaporte
	dataCRC = generarCRC(sistema, idTitular, validaDesde, validaHasta, payLoad)		
	if (codeCRC = dataCRC) then	
		'2.- Se valida el sistema
		if (sistema = pSistema) then
			'3.- Se valida que el pasaporte sea para el proveedor/usuario				
			if (CLng(idTitular) = CLng(pIdTitular)) then
				'4.- Se valida que el pasaporte no haya expirado.
				if (GF_DTEDIFF(session("MmtoSistema"), validaHasta, "D") >= 0) and (GF_DTEDIFF(session("MmtoSistema"), validaDesde, "D") =< 0) then 		
					result = PASSPORT_INF_VALID
				else
					result = PASSPORT_OUT_OF_DATE	
				end if					
			else
				result = PASSPORT_BAD_OWNER 	
			end if	
		else
			result = PASSPORT_BAD_SYSTEM 	
		end if
	else
		result = PASSPORT_NOT_VALID 	
	end if

	validarPasaporte = result
End Function
'---------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	09/11/2010
' Objetivo:	
'			Codigica datos a incluir en un pasaporte.
' Parametros:
'			payLoad [str] 	String donde se incluyen los nuevos datos
'			key 	[str]	Clave para el nuevo dato 	
'			val		[str]	El nuevo dato
' Devuelve:
'			String con formato valido para los datos a incluir en un pasaporte.
'			La respuesta se obtiene en el parametro "payLoad".
Function addPayloadData(ByRef payLoad, key, val)
	if (payLoad <> "") then	payLoad = payLoad & PASSPORT_DATA_TOKEN
	payLoad = payLoad & key & PASSPORT_FIELD_TOKEN & val	
End Function
'---------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	09/11/2010
' Objetivo:	
'			Recupera los datos incluidos en el pasaporte
' Parametros:
'			passport [str] 			Pasaporte.
'			data 	 [dictionary]	Diccionario con los datos recuperados
' Devuelve:
'			Si todo esta OK, en el parametro "data" davuelve los datos y la función retorna true. 
'			Si algo falla devuelve un diccionario vacio y retorna false.
Function retrievePayload(passport, ByRef data)
	Dim ret, temp, temp1, k
	Set data = Server.CreateObject("Scripting.Dictionary")
	Call leerPassporte(pasaporte, sistema, idTitular, validaDesde, validaHasta, payLoad, codeCRC)
	'result = false	
	'Valido los datos obtenidos del pasaporte
	retrievePayload = false
	dataCRC = generarCRC(sistema, idTitular, validaDesde, validaHasta, payLoad)		
	if (codeCRC = dataCRC) then	
		'Los datos son validos, obtengo los datos
		temp = split(payLoad, PASSPORT_DATA_TOKEN)		
		for k = 0 to UBound(temp)
			temp1 = split(temp(k), PASSPORT_FIELD_TOKEN)
			data.Add temp1(0), temp1(1)
		next
		retrievePayload = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Genera una firma MD5 del archivo inficado.
Function generateFileSignature(pPath)
	Dim fso
	
	Set fso = CreateObject("Scripting.FileSystemObject")			
	if (fso.FileExists(pPath)) Then
		text = fso.OpenTextFile(pPath).ReadAll
		generateFileSignature = MD5(text)
	else
		generateFileSignature = ""
	end if
	Set fso = nothing
End Function
'---------------------------------------------------------------------------------------------
Function initTaskAccessInfo(p_IdTask, p_IdDivision)    
    if (not CheckAccess(p_IdTask, p_IdDivision)) then response.redirect SITE_ROOT & "comprasAccesoDenegado.asp"	
End Function
'---------------------------------------------------------------------------------------------
'Esta funcion sirve para leer todos los datos de seguridad del usuario que luego sera accedida en las aplicaciones.
'Cada vez que se llame se recarga la info de seguridad. Conveniente que se haga en la p{agina principal para que solo se ejecute una vez.
Function LoadAccessInfo(pUsername)
	Dim Sql, rs, oConn, params, myTarea, flagAny
				
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLUSUARIOTAREAS_GET_BY_NOMBREUSUARIO", pUsername)    
	if (not rs.eof) then
		flagAny = false	
		myTarea = rs("IDTarea")
		while (not rs.eof)	
			if (CLng(myTarea) <> CLng(rs("IDTarea"))) then
				session(myTarea & "_ANY") = "0"										'Flag de permiso para alguna division.
				if (flagAny) then session(myTarea & "_ANY") = "1"
				myTarea = rs("IDTarea")
				flagAny = false
			end if		
			session(rs("IDTarea") & "_" & rs("IDDivision")) = rs("NivelPermiso")    'Permiso para una division especifica
			if (CInt(rs("NivelPermiso")) > 0) then flagAny = True
			rs.MoveNext()
		wend
		'Controlo la ultima
		session(myTarea & "_ANY") = "0"												'Flag de permiso para alguna division.
		if (flagAny) then session(myTarea & "_ANY") = "1"
	end if
End Function
'---------------------------------------------------------------------------------------------
'Funcion que permite saber si el usuario tiene algun tipo de acceso a una determinada tarea/division.
'(No importa que acceso, solo improta saber si tenga alguno)
Function CheckAccess(pIdTask, pIdDivision)  
    CheckAccess = False	
	if (pIdDivision = "") then 	
		if (session(pIdTask & "_ANY") = "1") then CheckAccess = True
    else
		'Primero verifico para la division solicitada, sino verifico para todas las divisiones (esto ultimo se da si el permiso no diferencia por division.
		if (session(pIdTask & "_" & pIdDivision) <> "") then  
			  CheckAccess = True
		else          
			if (session(pIdTask & "_" & DIVISION_ALL) <> "") then CheckAccess = True			
		end if            
	end if
End Function
'---------------------------------------------------------------------------------------------
'Funcion que permite saber si el usuario tiene acceso de Lectura/visualizacion a una determinada tarea/division.
Function hasReadAcess(pIdTask, pIdDivision)  
    hasReadAcess = False	
	if (pIdDivision <> "") then		
		'Primero verifico para la division solicitada, sino verifico para todas las divisiones (esto ultimo se da si el permiso no diferencia por division.
		if (session(pIdTask & "_" & pIdDivision) <> "") then  
			  if (CInt(session(pIdTask & "_" & pIdDivision)) >= TASK_LVL_VW_ACCESS) then hasReadAcess = True
		else          
			if (session(pIdTask & "_" & DIVISION_ALL) <> "") then 
				if (CInt(session(pIdTask & "_" & DIVISION_ALL)) >= TASK_LVL_VW_ACCESS) then hasReadAcess = True
			end if
		end if            
	end if
End Function
'---------------------------------------------------------------------------------------------
'Funcion que permite saber si el usuario tiene acceso de Escritura/Actualizacion a una determinada tarea/division.
Function hasWriteAcess(pIdTask, pIdDivision)  
    hasWriteAcess = False	
	if (pIdDivision <> "") then		
		'Primero verifico para la division solicitada, sino verifico para todas las divisiones (esto ultimo se da si el permiso no diferencia por division.
		if (session(pIdTask & "_" & pIdDivision) <> "") then  
			  if (CInt(session(pIdTask & "_" & pIdDivision)) >= TASK_LVL_WK_ACCESS) then hasWriteAcess = True
		else          
			if (session(pIdTask & "_" & DIVISION_ALL) <> "") then 
				if (CInt(session(pIdTask & "_" & DIVISION_ALL)) >= TASK_LVL_WK_ACCESS) then hasWriteAcess = True
			end if
		end if            
	end if
End Function
'---------------------------------------------------------------------------------------------
Function GetTaskDivisionAccessList(pIdTask, pUsename)
	Dim strSQL, rs, ret
	
	GetTaskDivisionAccessList = ""
	strSQL="Select IDDIVISION from TBLUSUARIOTAREAS where IDTAREA=" & pIdTask & " and NombreUsuario='" & pUsename & "' and NivelPermiso > 0"		
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	if (not rs.eof) then 
		ret = rs.GetString(2,,"",", ")
		GetTaskDivisionAccessList = Left(ret, Len(ret) - 2)
	end if
End Function
'---------------------------------------------------------------------------------------------
Function getLckUser(pSeed)
    getLckUser = pSeed & "-" & session.SessionID 
End Function
'---------------------------------------------------------------------------------------------
Function checkLckKey(pDbSite, pKey, pUsr)
    Dim arr, diff, ret    
    ret=false
    strSQL="Select * from PARAMETROS where CDPARAMETRO='" & pKey & "'"
    Call executequeryDb(pDbSite, rs, "OPEN", strSQL)
    if (not rs.eof) then
        if (rs("VLPARAMETRO") <> "") then
            arr = Split(rs("VLPARAMETRO"), "|")
            diff = GF_DTEDIFF(arr(1), session("MmtoDato"), "S")                
            if (diff > 60) then 
                ret = true
            else
                if (pUsr = arr(0)) then ret = true
            end if                
        else
            ret = true            
        end if       
        if (ret) then 
            strSQL="Update PARAMETROS set VLPARAMETRO='" & pUsr & "|" & session("MmtoDato") & "' where CDPARAMETRO='" & pKey & "'"
            Call executequeryDb(pDbSite, rs, "EXEC", strSQL)
        end if            
    else
        ret = true
        strSQL="Insert into PARAMETROS values('" & pKey & "', 'Lock Key Logistica', '" & pUsr & "|" & session("MmtoDato") & "')"             
        Call executequeryDb(pDbSite, rs, "EXEC", strSQL)
    end if
    checkLckKey = ret
End Function
'---------------------------------------------------------------------------------------------
Function releaseLckKey(pDbSite, pUsr, pKey)
	Dim flagDo
	
	flagDo = false
	strSQL="Select * from PARAMETROS where CDPARAMETRO='" & pKey & "'"
    Call executequeryDb(pDbSite, rs, "OPEN", strSQL)
    if (not rs.eof) then
		if (rs("VLPARAMETRO") <> "") then
            arr = Split(rs("VLPARAMETRO"), "|")
			if (pUsr = arr(0)) then flagDo = true	
		end if
	end if
	if (flagDo) then
		strSQL="Update PARAMETROS set VLPARAMETRO='' where CDPARAMETRO='" & pKey & "'"
		Call executequeryDb(pDbSite, rs, "EXEC", strSQL)
	end if
End Function
%>